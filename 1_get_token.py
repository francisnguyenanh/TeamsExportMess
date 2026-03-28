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

# Danh sách client ID của các Microsoft first-party apps
# (luôn có service principal trong mọi Microsoft 365 tenant)
# Thử lần lượt — app nào được tenant cho phép sẽ dùng app đó.
_CANDIDATE_CLIENT_IDS = [
    ("Microsoft Teams",          "1fec8e78-bce4-4aaf-ab1b-5451cc387264"),
    ("Microsoft Teams (Mobile)", "d3590ed6-52b3-4102-aeff-aad2292ab01c"),
    ("Azure CLI",                "04b07795-8ddb-461a-bbee-02f9e1bf7b46"),
    ("Microsoft Office",         "d3590ed6-52b3-4102-aeff-aad2292ab01c"),
]

TEAMS_CLIENT_ID  = "1fec8e78-bce4-4aaf-ab1b-5451cc387264"   # mặc định, ghi đè bởi auto-detect
GRAPH_SCOPES     = [
    "https://graph.microsoft.com/Team.ReadBasic.All",
    "https://graph.microsoft.com/Channel.ReadBasic.All",
    "https://graph.microsoft.com/ChannelMessage.Read.All",
    "https://graph.microsoft.com/Chat.Read",
    "https://graph.microsoft.com/Chat.ReadBasic",
    "https://graph.microsoft.com/ChatMessage.Read",
]
# Dùng tenant ID cụ thể thay vì "organizations" để tránh lỗi AADSTS1001010
# Sẽ được thay thế bằng tenant ID thực tế trong get_token_device_code()
AUTHORITY        = "https://login.microsoftonline.com/organizations"

TOKEN_FILE = "token.json"


# ── Helpers: tự detect client ID phù hợp với tenant ─────────────────────────

def _detect_tenant_id(email: str) -> str | None:
    """
    Lấy tenant ID từ email bằng cách query OpenID metadata.
    Ví dụ: user@company.com → tenant ID của company.com
    """
    import urllib.request
    domain = email.split("@")[-1] if "@" in email else None
    if not domain:
        return None
    try:
        url  = f"https://login.microsoftonline.com/{domain}/.well-known/openid-configuration"
        with urllib.request.urlopen(url, timeout=5) as r:
            data = json.loads(r.read())
        # token_endpoint = https://login.microsoftonline.com/{tenant_id}/oauth2/token
        token_ep = data.get("token_endpoint", "")
        parts    = token_ep.split("/")
        idx      = parts.index("microsoftonline.com") if "microsoftonline.com" in parts else -1
        if idx >= 0 and idx + 1 < len(parts):
            return parts[idx + 1]
    except Exception:
        pass
    return None


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


# ── Method D: MSAL Device Code Flow ──────────────────────────────────────────

def get_token_device_code(email: str = "", scopes: list = None,
                          on_code: callable = None) -> str | None:
    """
    Lấy token qua MSAL Device Code Flow.
    Tự động thử nhiều client ID cho đến khi tìm được app được tenant cho phép.

    Tham số:
        email    : email người dùng — dùng để detect tenant ID chính xác
        scopes   : danh sách scope Graph API cần
        on_code  : callback(user_code, url) — gọi khi có mã để hiển thị
    """
    import msal

    if scopes is None:
        scopes = GRAPH_SCOPES

    # Xác định authority: dùng tenant ID cụ thể nếu biết email
    tenant_id = _detect_tenant_id(email) if email else None
    authority = (
        f"https://login.microsoftonline.com/{tenant_id}"
        if tenant_id
        else "https://login.microsoftonline.com/organizations"
    )
    if tenant_id:
        print(f"🏢 Detected tenant: {tenant_id}")

    last_error = ""
    for app_name, client_id in _CANDIDATE_CLIENT_IDS:
        print(f"🔑 Thử client: {app_name} ({client_id[:8]}…)")
        app = msal.PublicClientApplication(client_id, authority=authority)

        # Thử cache trước
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(scopes, account=accounts[0])
            if result and "access_token" in result:
                print(f"✅ Dùng token từ MSAL cache ({app_name}).")
                return result["access_token"]

        flow = app.initiate_device_flow(scopes=scopes)

        if "user_code" not in flow:
            err = flow.get("error_description", flow.get("error", "unknown"))
            # Lỗi service principal = tenant không có app này → thử app khác
            if "AADSTS1001010" in err or "AADSTS700016" in err or "does not exist" in err.lower():
                print(f"   ⚠️  Tenant không có {app_name} — thử app tiếp theo…")
                last_error = err
                continue
            print(f"❌ Lỗi khởi tạo device flow: {err}")
            return None

        # Có mã rồi — hiển thị cho user
        if on_code:
            on_code(flow["user_code"], flow["verification_uri"])
        else:
            print("\n" + "="*60)
            print("📱 ĐĂNG NHẬP THEO BƯỚC SAU:")
            print(f"   1. Mở trình duyệt: {flow['verification_uri']}")
            print(f"   2. Nhập mã       : {flow['user_code']}")
            print(f"   3. Đăng nhập tài khoản Microsoft / Teams")
            print(f"   ⏰ Hết hạn sau   : {flow.get('expires_in', 900)//60} phút")
            print("="*60 + "\n")

        result = app.acquire_token_by_device_flow(flow)

        if "access_token" in result:
            print(f"✅ Đăng nhập thành công! (via {app_name})")
            return result["access_token"]

        err = result.get("error_description", result.get("error", ""))
        print(f"❌ Thất bại ({app_name}): {err}")
        return None   # user đã thấy code rồi → không thử app khác nữa

    print(f"❌ Không có client ID nào được tenant chấp nhận.\nLỗi cuối: {last_error}")
    return None


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Lấy Bearer token từ Teams")
    parser.add_argument("--method", choices=["sso", "password", "device"], default="device",
                        help="Phương thức lấy token (mặc định: device)")
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
    if args.method == "device":
        token = get_token_device_code()
    elif args.method == "sso":
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
