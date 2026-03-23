"""
1_get_token.py
==============
Lấy Bearer token từ Microsoft Teams Web.

Cách dùng:
    python 1_get_token.py              # SSO (mặc định)
    python 1_get_token.py --method sso
    python 1_get_token.py --method password
    python 1_get_token.py --browser chrome
"""

import argparse
import base64
import json
import os
import getpass
from datetime import datetime, timezone
from pathlib import Path

TOKEN_FILE = "token.json"


# ── JWT helpers ───────────────────────────────────────────────────────────────

def decode_token_expiry(token: str) -> datetime | None:
    """Đọc thời gian hết hạn từ JWT payload (không cần thư viện jwt)."""
    try:
        payload_b64 = token.split(".")[1]
        payload_b64 += "=" * (4 - len(payload_b64) % 4)
        payload = json.loads(base64.b64decode(payload_b64))
        exp = payload.get("exp", 0)
        return datetime.fromtimestamp(exp, tz=timezone.utc)
    except Exception:
        return None


def save_token(token: str):
    """Lưu token + metadata vào token.json."""
    expiry = decode_token_expiry(token)
    data = {
        "token": token,
        "fetched_at": datetime.now(timezone.utc).isoformat(),
        "expires_at": expiry.isoformat() if expiry else None,
    }
    Path(TOKEN_FILE).write_text(json.dumps(data, indent=2), encoding="utf-8")

    if expiry:
        remaining = int((expiry - datetime.now(timezone.utc)).total_seconds() / 60)
        print(f"⏱  Token hết hạn lúc: {expiry.strftime('%H:%M:%S')} "
              f"(còn ~{remaining} phút)")
    print(f"💾 Đã lưu token vào {TOKEN_FILE}")


def load_token() -> str | None:
    """
    Đọc token từ file.
    Trả về None nếu file không tồn tại hoặc token còn < 5 phút.
    """
    if not Path(TOKEN_FILE).exists():
        return None

    data = json.loads(Path(TOKEN_FILE).read_text(encoding="utf-8"))
    token = data.get("token")
    expires_at = data.get("expires_at")

    if expires_at:
        exp = datetime.fromisoformat(expires_at)
        remaining = (exp - datetime.now(timezone.utc)).total_seconds()
        if remaining < 300:  # Còn dưới 5 phút → coi như hết hạn
            print(f"⚠️  Token hết hạn (còn {int(remaining)}s), cần lấy mới.")
            return None
        print(f"✅ Dùng token đã lưu (còn ~{int(remaining / 60)} phút)")

    return token


# ── Profile path helpers ──────────────────────────────────────────────────────

def get_profile_path(browser: str) -> str:
    """Trả về đường dẫn profile Windows cho Edge hoặc Chrome."""
    if browser == "edge":
        return os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\User Data")
    else:
        return os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data")


# ── Method C: SSO ─────────────────────────────────────────────────────────────

def get_token_sso(browser: str = "edge") -> str | None:
    """
    Mở browser với profile Windows hiện tại → SSO tự động đăng nhập →
    bắt Bearer token từ network request.

    Hoạt động tốt nhất khi:
    - Máy đã join domain công ty
    - Đã đăng nhập Teams trên máy (Windows Credential Manager có sẵn session)
    """
    from playwright.sync_api import sync_playwright

    captured = {"token": None}
    profile_dir = get_profile_path(browser)

    print(f"🌐 Mở {browser.capitalize()} với profile: {profile_dir}")

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=profile_dir,
            channel=f"ms{browser}" if browser == "edge" else browser,
            headless=False,           # Phải hiện browser để SSO xử lý
            args=["--no-sandbox", "--disable-dev-shm-usage"],
            ignore_default_args=["--enable-automation"],  # Tránh banner "controlled by automation"
        )

        page = context.new_page()

        # Bắt token từ mọi request tới graph.microsoft.com
        def on_request(request):
            auth = request.headers.get("authorization", "")
            if (auth.startswith("Bearer ")
                    and "graph.microsoft.com" in request.url
                    and not captured["token"]):
                token = auth.removeprefix("Bearer ")
                if len(token) > 200:  # Token thật thường > 500 ký tự
                    captured["token"] = token
                    print("✅ Đã bắt được Bearer token!")

        page.on("request", on_request)

        print("⏳ Đang load Teams, chờ SSO xử lý...")
        try:
            page.goto("https://teams.microsoft.com", timeout=60_000)
            # Chờ sidebar Teams xuất hiện = đã đăng nhập xong
            page.wait_for_selector('[data-tid="team-channel-list"]', timeout=45_000)
            print("✅ Teams đã load xong!")
        except Exception as e:
            print(f"⚠️  Timeout/lỗi: {e}")

        # Nếu chưa bắt được token (Teams load nhưng chưa call API),
        # click vào channel đầu tiên để trigger request
        if not captured["token"]:
            print("🖱  Thử click vào channel để trigger API call...")
            try:
                page.locator('[data-tid="channel-list-item"]').first.click()
                page.wait_for_timeout(4_000)
            except Exception:
                page.wait_for_timeout(4_000)

        context.close()

    return captured["token"]


# ── Method B: Email/Password ──────────────────────────────────────────────────

def get_token_password(browser: str = "edge") -> str | None:
    """
    Tự động điền email + password vào trang đăng nhập Microsoft.

    ⚠️  Không hoạt động khi tổ chức bật:
        - MFA / Authenticator app
        - Conditional Access (chỉ cho phép thiết bị được quản lý)
        - ADFS / custom SSO provider

    Password được nhập trực tiếp từ terminal, không lưu vào file.
    """
    from playwright.sync_api import sync_playwright

    email = input("📧 Nhập email Teams: ").strip()
    password = getpass.getpass("🔑 Nhập password (ẩn): ")

    captured = {"token": None}

    with sync_playwright() as p:
        browser_obj = p.chromium.launch(
            channel=f"ms{browser}" if browser == "edge" else browser,
            headless=False,
        )
        context = browser_obj.new_context()
        page = context.new_page()

        def on_request(request):
            auth = request.headers.get("authorization", "")
            if (auth.startswith("Bearer ")
                    and "graph.microsoft.com" in request.url
                    and not captured["token"]):
                token = auth.removeprefix("Bearer ")
                if len(token) > 200:
                    captured["token"] = token
                    print("✅ Đã bắt được Bearer token!")

        page.on("request", on_request)

        print("🌐 Mở trang đăng nhập Microsoft...")
        page.goto("https://login.microsoftonline.com", timeout=30_000)

        # Điền email
        page.wait_for_selector('input[type="email"]', timeout=15_000)
        page.fill('input[type="email"]', email)
        page.click('input[type="submit"]')

        # Điền password
        page.wait_for_selector('input[type="password"]', timeout=15_000)
        page.fill('input[type="password"]', password)
        page.click('input[type="submit"]')

        # Bỏ qua "Stay signed in?"
        try:
            page.wait_for_selector('#idBtn_Back', timeout=5_000)
            page.click('#idBtn_Back')  # Click "No"
        except Exception:
            pass

        # Chuyển sang Teams
        print("⏳ Chuyển sang Teams...")
        page.goto("https://teams.microsoft.com", timeout=60_000)
        try:
            page.wait_for_selector('[data-tid="team-channel-list"]', timeout=45_000)
        except Exception:
            pass

        if not captured["token"]:
            try:
                page.locator('[data-tid="channel-list-item"]').first.click()
                page.wait_for_timeout(4_000)
            except Exception:
                page.wait_for_timeout(4_000)

        browser_obj.close()

    return captured["token"]


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Lấy Bearer token từ Teams")
    parser.add_argument("--method", choices=["sso", "password"], default="sso",
                        help="Phương thức lấy token (mặc định: sso)")
    parser.add_argument("--browser", choices=["edge", "chrome"], default="edge",
                        help="Trình duyệt sử dụng (mặc định: edge)")
    parser.add_argument("--force", action="store_true",
                        help="Bỏ qua token đã lưu, lấy token mới")
    args = parser.parse_args()

    # Thử dùng token cũ trước (trừ khi --force)
    if not args.force:
        token = load_token()
        if token:
            return token

    # Lấy token mới
    print(f"\n🔄 Lấy token mới bằng phương thức: {args.method.upper()}")
    if args.method == "sso":
        token = get_token_sso(browser=args.browser)
    else:
        token = get_token_password(browser=args.browser)

    if token:
        save_token(token)
        print("\n✅ Hoàn tất! Bạn có thể chạy tiếp 3_export.py")
    else:
        print("\n❌ Không lấy được token!")
        print("   Thử: tăng timeout, kiểm tra browser đã cài driver chưa")

    return token


if __name__ == "__main__":
    main()
