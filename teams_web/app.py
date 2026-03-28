"""
teams_web/app.py
================
Flask web interface for the Microsoft Teams Export Tool.

Imports and calls existing CLI modules (1_get_token.py, 2_list_channels.py, 3_export.py).
No logic is rewritten here — only web-layer wiring.

Run:
    cd teams_web
    python app.py
  or:
    flask --app teams_web/app.py run
"""

import importlib.util
import io
import json
import os
import re
import sys
import threading
import time
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from flask import (Flask, Response, flash, jsonify, redirect,
                   render_template, request, send_file, url_for)

# ── Bootstrap: resolve base dir and load CLI modules ──────────────────────────
BASE_DIR = Path(__file__).parent.parent   # project root (where CLI scripts live)
os.chdir(BASE_DIR)                        # ensure relative paths in CLI modules work

sys.path.insert(0, str(BASE_DIR))


def _load_module(alias: str, filename: str):
    """Load a .py file with a numeric prefix (can't be imported via normal import)."""
    spec = importlib.util.spec_from_file_location(alias, BASE_DIR / filename)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[alias] = mod
    spec.loader.exec_module(mod)
    return mod


_load_module("get_token",     "1_get_token.py")
_load_module("list_channels", "2_list_channels.py")
_load_module("export_mod",    "3_export.py")

# Import from the CLI modules
from get_token import (decode_token_expiry, get_token_sso,   # noqa: E402
                       load_token, save_token)
from get_token_helper import get_valid_token                  # noqa: E402
import list_channels as lc                                    # noqa: E402
import export_mod    as em                                    # noqa: E402

# export_docx lives in BASE_DIR
sys.path.insert(0, str(BASE_DIR))
import export_docx                                            # noqa: E402


# ── Helpers ────────────────────────────────────────────────────────────────────

def _safe_sheet_name(name: str, max_len: int = 31) -> str:
    """Sanitize a string for use as an Excel sheet name.
    Excel forbids these characters in sheet names: \\ / * ? : [ ]
    and limits length to 31 characters."""
    import re as _re
    cleaned = _re.sub(r'[\\/*?:\[\]]', '_', name)
    return cleaned[:max_len] if cleaned else "Sheet1"


def _safe_filename(name: str, max_len: int = 80) -> str:
    """Sanitize a string for use as a filename (Windows-safe)."""
    import re as _re
    return _re.sub(r'[\\/*?:"<>|\[\]]', '_', name)[:max_len]


# ── Beta-endpoint fallback for message fetching ───────────────────────────────
# GET /v1.0/teams/{id}/channels/{id}/messages requires ChannelMessage.Read.All
# with admin consent.  The /beta endpoint tolerates delegated Teams-app tokens.

def _fetch_messages_beta(team_id: str, channel_id: str, headers: dict,
                         date_from: str = "", date_to: str = "",
                         include_replies: bool = True) -> list[dict]:
    """
    Same logic as em.fetch_messages() but hits the /beta endpoint.
    Used as automatic fallback when /v1.0 returns 403.
    """
    import requests as _req

    def _api_get(url):
        resp = _req.get(url, headers=headers, timeout=30)
        if resp.status_code == 401:
            raise PermissionError("TOKEN_EXPIRED")
        if resp.status_code == 403:
            raise PermissionError(f"ACCESS_DENIED: {url}")
        resp.raise_for_status()
        return resp.json()

    url = (f"https://graph.microsoft.com/beta"
           f"/teams/{team_id}/channels/{channel_id}/messages")
    rows = []
    page_num = 0

    while url:
        page_num += 1
        data = _api_get(url)

        for msg in data.get("value", []):
            if msg.get("messageType") != "message":
                continue
            created = msg.get("createdDateTime", "")
            if date_from and created < date_from:
                continue
            if date_to and created > date_to + "T23:59:59Z":
                continue

            rows.append(em.parse_msg(msg, is_reply=False))

            if include_replies:
                replies_url = (f"https://graph.microsoft.com/beta"
                               f"/teams/{team_id}/channels/{channel_id}"
                               f"/messages/{msg['id']}/replies")
                try:
                    rdata = _api_get(replies_url)
                    for r in sorted(rdata.get("value", []),
                                    key=lambda x: x.get("createdDateTime", "")):
                        if r.get("messageType") == "message":
                            rows.append(em.parse_msg(r, is_reply=True))
                except PermissionError:
                    raise
                except Exception:
                    pass

        url = data.get("@odata.nextLink")

    return rows


# ── Group Chat message fetching ───────────────────────────────────────────────

def _fetch_chat_messages(chat_id: str, headers: dict,
                         date_from: str = "", date_to: str = "") -> list[dict]:
    """
    Lấy tin nhắn từ Group Chat hoặc 1:1 Chat.
    Thử Graph API trước, fallback sang Teams chatsvc API.
    """
    import requests as _req

    # ── Strategy 1: Graph API ────────────────────────────────────────────
    for api_ver in ("v1.0", "beta"):
        url = f"https://graph.microsoft.com/{api_ver}/me/chats/{chat_id}/messages?$top=50"
        try:
            resp = _req.get(url, headers=headers, timeout=30)
            if resp.status_code in (401, 403):
                continue  # Thử version khác hoặc fallback
            resp.raise_for_status()
            data = resp.json()
            rows = []

            # Nếu có data → parse hết → convert to rich format
            while True:
                for msg in data.get("value", []):
                    if msg.get("messageType") != "message":
                        continue
                    created = msg.get("createdDateTime", "")
                    if date_from and created < date_from:
                        continue
                    if date_to and created > date_to + "T23:59:59Z":
                        continue
                    parsed = export_docx.parse_chatsvc_message(msg)
                    if parsed:
                        rows.append(parsed)

                page_url = data.get("@odata.nextLink")
                if not page_url:
                    break
                resp2 = _req.get(page_url, headers=headers, timeout=30)
                if resp2.status_code in (401, 403):
                    break
                resp2.raise_for_status()
                data = resp2.json()

            return rows
        except Exception:
            continue

    # ── Strategy 2: Teams chatsvc API (dùng captured token) ──────────────
    return _fetch_chat_messages_chatsvc(chat_id, date_from=date_from, date_to=date_to)


def _fetch_chat_messages_chatsvc(chat_id: str,
                                 date_from: str = "",
                                 date_to: str = "") -> list[dict]:
    """
    Lấy tin nhắn chat qua Teams chatsvc internal API.
    Dùng token đã capture từ Playwright.

    Returns list of rich message dicts (parsed by export_docx.parse_chatsvc_message).
    Each dict has: sender, datetime, content_raw, segments, images, attachments, message_id.
    """
    import requests as _req

    # Validate chat_id format — phải là thread ID (19:xxx@thread.v2)
    if not chat_id or "19:" not in chat_id or "@thread" not in chat_id:
        raise PermissionError(
            f"INVALID_CHAT_ID: '{chat_id[:30]}…' không phải thread ID hợp lệ. "
            "Cần format '19:xxx@thread.v2'. Hãy chạy lại 'Scrape Chats (Playwright)' ở Step 2."
        )

    # Refresh token nếu hết hạn
    _refresh_chatsvc_token_if_needed()

    # Load chatsvc token
    chatsvc_file = BASE_DIR / "chatsvc_token.json"
    if not chatsvc_file.exists():
        raise PermissionError(
            "ACCESS_DENIED: Không có chatsvc token. "
            "Hãy chạy 'Scrape Chats (Playwright)' ở Step 2 trước."
        )
    td = json.loads(chatsvc_file.read_text(encoding="utf-8"))
    token = td.get("token", "")
    base_url = td.get("base_url", "https://teams.microsoft.com/api/chatsvc/apac/v1")

    if not token:
        raise PermissionError("ACCESS_DENIED: chatsvc token rỗng.")

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    # URL encode chat_id nếu cần (19:xxx@thread.v2)
    from urllib.parse import quote
    encoded_id = quote(chat_id, safe="")

    # Fetch messages — chatsvc endpoint
    # Thử nhiều endpoint format
    endpoints = [
        f"{base_url}/users/ME/conversations/{encoded_id}/messages?pageSize=200&view=msnp24Equivalent",
        f"{base_url}/users/ME/conversations/{encoded_id}/messages?pageSize=200",
    ]

    messages_raw = []

    for ep_url in endpoints:
        try:
            url = ep_url
            page_num = 0
            while url and page_num < 100:
                page_num += 1
                resp = _req.get(url, headers=headers, timeout=30)
                if resp.status_code in (401, 403):
                    break
                if resp.status_code == 404:
                    break
                resp.raise_for_status()
                data = resp.json()

                msgs = data.get("messages", [])
                if not msgs:
                    break
                messages_raw.extend(msgs)

                # Pagination: _metadata.backwardLink
                meta = data.get("_metadata", {})
                url = meta.get("backwardLink") or meta.get("syncState")
                if not url:
                    break
                # backwardLink có thể là relative URL
                if url.startswith("/"):
                    url = f"https://teams.microsoft.com{url}"

            if messages_raw:
                break  # Đã lấy được messages
        except Exception:
            continue

    if not messages_raw:
        # Thử thêm threads endpoint
        try:
            thread_url = f"{base_url}/threads/{encoded_id}/messages?pageSize=200"
            resp = _req.get(thread_url, headers=headers, timeout=30)
            if resp.status_code == 200:
                data = resp.json()
                messages_raw = data.get("messages", [])
        except Exception:
            pass

    if not messages_raw:
        raise PermissionError(
            f"ACCESS_DENIED: chatsvc API không trả về messages cho {chat_id[:30]}… "
            "Token có thể đã hết hạn — chạy lại Playwright Scrape."
        )

    # ── Parse raw messages → rich format via export_docx ─────────────────
    rows = []
    for msg in messages_raw:
        # Date filter trước khi parse
        composed = (msg.get("composetime")
                    or msg.get("createdDateTime")
                    or msg.get("originalarrivaltime", ""))
        if date_from and composed and composed < date_from:
            continue
        if date_to and composed and composed > date_to + "T23:59:59Z":
            continue

        parsed = export_docx.parse_chatsvc_message(msg)
        if parsed:
            rows.append(parsed)

    # Sort by datetime (oldest first)
    rows.sort(key=lambda r: r.get("datetime", ""))

    return rows


def _refresh_chatsvc_token_if_needed():
    """Check if chatsvc token is expired and refresh via quick Playwright session."""
    chatsvc_file = BASE_DIR / "chatsvc_token.json"
    if not chatsvc_file.exists():
        return False

    try:
        td = json.loads(chatsvc_file.read_text(encoding="utf-8"))
        token = td.get("token", "")
        if not token:
            return False

        # Decode JWT to check exp
        import base64 as _b64
        parts = token.split(".")
        if len(parts) >= 2:
            padded = parts[1] + "=" * (-len(parts[1]) % 4)
            payload = json.loads(_b64.urlsafe_b64decode(padded))
            exp = payload.get("exp", 0)
            remaining = exp - time.time()
            if remaining > 300:  # > 5 min left
                return True  # Token still valid
    except Exception:
        pass

    # Token expired or about to expire → refresh
    print("[chatsvc] Token expired, refreshing via Playwright…", flush=True)
    try:
        from playwright.sync_api import sync_playwright
        captured_new = {}

        with sync_playwright() as pw:
            context = pw.chromium.launch_persistent_context(
                user_data_dir=str(_PW_SESSION_DIR),
                headless=True,  # Chạy ẩn
                channel="chromium",
                args=["--disable-blink-features=AutomationControlled", "--no-first-run"],
                ignore_default_args=["--enable-automation"],
                viewport={"width": 1280, "height": 900},
            )

            page = context.pages[0] if context.pages else context.new_page()

            def _on_req(req):
                url = req.url
                auth = req.headers.get("authorization", "")
                if "chatsvc" in url and auth.startswith("Bearer ") and not captured_new.get("token"):
                    captured_new["token"] = auth.removeprefix("Bearer ").strip()
                    import re as _re
                    m = _re.search(r"(https://teams\.microsoft\.com/api/chatsvc/[^/]+/v\d+)", url)
                    if m:
                        captured_new["base_url"] = m.group(1)

            page.on("request", _on_req)

            try:
                page.goto("https://teams.microsoft.com", timeout=30_000)
            except Exception:
                pass

            # Wait for token capture
            for _ in range(15):
                time.sleep(1)
                if captured_new.get("token"):
                    break

            context.close()

        if captured_new.get("token"):
            chatsvc_file.write_text(json.dumps({
                "token": captured_new["token"],
                "base_url": captured_new.get("base_url", td.get("base_url", "")),
                "fetched_at": datetime.now(timezone.utc).isoformat(),
            }, ensure_ascii=False, indent=2), encoding="utf-8")
            print("[chatsvc] Token refreshed!", flush=True)
            return True
    except Exception as e:
        print(f"[chatsvc] Refresh failed: {e}", flush=True)

    return False


# ── File paths ────────────────────────────────────────────────────────────────
CONFIG_FILE       = BASE_DIR / "config.json"
EXPORT_STATE_FILE = BASE_DIR / "export_state.json"
ALL_CHANNELS_FILE = BASE_DIR / "all_channels.json"
OUTPUT_DIR        = BASE_DIR / "output"

# ── Flask app ─────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "teams-export-dev-secret-change-in-prod")

# Background job registry: { job_id: {"status": str, "progress": list, "files": list} }
jobs: dict[str, dict] = {}


# ── Config helpers ────────────────────────────────────────────────────────────

def load_config() -> dict:
    if CONFIG_FILE.exists():
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    return {
        "browser": "edge",
        "channels": [],
        "export": {
            "date_from": "",
            "date_to": "",
            "output_dir": "output",
            "include_replies": True,
        },
    }


def save_config(cfg: dict):
    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def load_export_state() -> dict:
    if EXPORT_STATE_FILE.exists():
        return json.loads(EXPORT_STATE_FILE.read_text(encoding="utf-8"))
    return {}


def save_export_state(state: dict):
    EXPORT_STATE_FILE.write_text(
        json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def load_all_channels() -> list:
    if ALL_CHANNELS_FILE.exists():
        return json.loads(ALL_CHANNELS_FILE.read_text(encoding="utf-8"))
    return []


def get_token_status() -> dict:
    """Returns a dict describing current token validity for templates."""
    token_file = BASE_DIR / "token.json"
    if not token_file.exists():
        return {"valid": False, "reason": "Token file not found"}
    try:
        data = json.loads(token_file.read_text(encoding="utf-8"))
        token = data.get("token")
        expires_at = data.get("expires_at")

        if not token:
            return {"valid": False, "reason": "No token in file"}

        if expires_at:
            exp = datetime.fromisoformat(expires_at)
            remaining = (exp - datetime.now(timezone.utc)).total_seconds()
            if remaining < 300:
                return {
                    "valid": False,
                    "reason": "Expired" if remaining < 0 else "Expires in < 5 min",
                }
            minutes = int(remaining / 60)
            return {
                "valid": True,
                "expires_at": exp.strftime("%H:%M:%S"),
                "minutes_remaining": minutes,
                "fetched_at": data.get("fetched_at", ""),
            }

        return {"valid": True, "expires_at": None, "minutes_remaining": None}
    except Exception as e:
        return {"valid": False, "reason": str(e)}


def group_channels_by_team(all_channels: list) -> list:
    """Group flat channel list into [{team_id, team_name, channels:[...]}]."""
    teams: dict[str, dict] = {}
    for ch in all_channels:
        tid = ch["team_id"]
        if tid not in teams:
            teams[tid] = {
                "team_id": tid,
                "team_name": ch["team_name"],
                "channels": [],
            }
        teams[tid]["channels"].append(ch)
    return list(teams.values())


# ── Web-compatible password token fetcher ────────────────────────────────────

def _get_token_password_web(email: str, password: str, browser: str = "edge"):
    """
    Playwright-based token fetch using form-provided credentials.
    This is the web adaptation of get_token_password() which normally reads stdin.
    """
    from playwright.sync_api import sync_playwright  # noqa: PLC0415

    captured = {"token": None}

    with sync_playwright() as p:
        browser_obj = p.chromium.launch(
            channel=f"ms{browser}" if browser == "edge" else browser,
            headless=False,
        )
        context = browser_obj.new_context()
        page = context.new_page()

        def on_request(request_obj):
            auth = request_obj.headers.get("authorization", "")
            if (auth.startswith("Bearer ")
                    and "graph.microsoft.com" in request_obj.url
                    and not captured["token"]):
                token_val = auth.removeprefix("Bearer ")
                if len(token_val) > 200:
                    captured["token"] = token_val

        page.on("request", on_request)

        page.goto("https://login.microsoftonline.com", timeout=30_000)
        page.wait_for_selector('input[type="email"]', timeout=15_000)
        page.fill('input[type="email"]', email)
        page.click('input[type="submit"]')

        page.wait_for_selector('input[type="password"]', timeout=15_000)
        page.fill('input[type="password"]', password)
        page.click('input[type="submit"]')

        try:
            page.wait_for_selector('#idBtn_Back', timeout=5_000)
            page.click('#idBtn_Back')
        except Exception:
            pass

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


# ── Cách 3: Dùng Chrome profile thật (decrypt cookies → inject vào Playwright) ─

def _decrypt_chrome_cookies(profile_dir: str) -> list[dict]:
    """
    Đọc và decrypt toàn bộ cookies của Chrome (Windows DPAPI + AES-256-GCM).
    Trả về list dict tương thích với Playwright context.add_cookies().
    Chỉ lấy cookies liên quan đến Microsoft / Teams.
    """
    import base64
    import shutil
    import sqlite3
    import tempfile
    import json as _json
    import win32crypt                          # pywin32
    from Crypto.Cipher import AES             # pycryptodome

    profile_path = Path(profile_dir)

    # ── Lấy encryption key từ Local State ───────────────────────────────────
    local_state_path = profile_path / "Local State"
    encrypted_key = None
    if local_state_path.exists():
        ls = _json.loads(local_state_path.read_text(encoding="utf-8"))
        key_b64 = ls.get("os_crypt", {}).get("encrypted_key", "")
        if key_b64:
            key_bytes = base64.b64decode(key_b64)
            # Bỏ prefix "DPAPI" (5 bytes)
            encrypted_key = win32crypt.CryptUnprotectData(
                key_bytes[5:], None, None, None, 0
            )[1]

    # ── Tìm Cookies DB ──────────────────────────────────────────────────────
    cookies_path = profile_path / "Default" / "Network" / "Cookies"
    if not cookies_path.exists():
        cookies_path = profile_path / "Default" / "Cookies"
    if not cookies_path.exists():
        raise FileNotFoundError(
            f"Không tìm thấy Cookies DB trong {profile_path}.\n"
            "Hãy kiểm tra Chrome profile directory."
        )

    # ── Copy Cookies DB sang temp file (tránh lock + tránh lỗi URI path) ────
    tmp_cookies = Path(tempfile.mktemp(suffix=".db", prefix="chrome_cookies_"))
    try:
        # Thử nhiều cách copy — Chrome lock file rất chặt trên Windows
        copied = False
        import subprocess

        # Cách 1: shutil.copy2 (nhanh nhất, chỉ hoạt động nếu Chrome ko lock)
        if not copied:
            try:
                shutil.copy2(cookies_path, tmp_cookies)
                copied = True
            except (PermissionError, OSError):
                pass

        # Cách 2: esentutl /y (Windows built-in, bypass exclusive file lock!)
        if not copied:
            try:
                r = subprocess.run(
                    ["esentutl", "/y", str(cookies_path), "/d", str(tmp_cookies)],
                    capture_output=True, timeout=15,
                )
                if tmp_cookies.exists() and tmp_cookies.stat().st_size > 0:
                    copied = True
            except Exception:
                pass

        # Cách 3: sqlite3 backup API with URI (nolock)
        if not copied:
            try:
                # Escape backslashes for URI
                uri_path = str(cookies_path).replace("\\", "/")
                src_conn = sqlite3.connect(f"file:///{uri_path}?mode=ro&nolock=1", uri=True, timeout=3)
                dst_conn = sqlite3.connect(str(tmp_cookies))
                src_conn.backup(dst_conn)
                src_conn.close()
                dst_conn.close()
                if tmp_cookies.exists() and tmp_cookies.stat().st_size > 0:
                    copied = True
            except Exception:
                pass

        # Cách 4: Đọc raw bytes
        if not copied:
            try:
                with open(cookies_path, "rb") as f:
                    data = f.read()
                with open(tmp_cookies, "wb") as f:
                    f.write(data)
                copied = True
            except (PermissionError, OSError):
                pass

        # Cách 5: PowerShell Copy-Item (dùng .NET FileStream share mode)
        if not copied:
            try:
                ps_cmd = (
                    f'[IO.File]::Copy("{cookies_path}", "{tmp_cookies}", $true)'
                )
                subprocess.run(
                    ["powershell", "-NoProfile", "-Command", ps_cmd],
                    capture_output=True, timeout=10,
                )
                if tmp_cookies.exists() and tmp_cookies.stat().st_size > 0:
                    copied = True
            except Exception:
                pass

        # Cách 6: PowerShell FileStream with share ReadWrite+Delete (bypass exclusive lock)
        if not copied:
            try:
                ps_script = (
                    f"$src = [System.IO.FileStream]::new('{cookies_path}', 'Open', 'Read', 'ReadWrite,Delete');"
                    f"$dst = [System.IO.FileStream]::new('{tmp_cookies}', 'Create', 'Write');"
                    f"$src.CopyTo($dst); $src.Close(); $dst.Close()"
                )
                subprocess.run(
                    ["powershell", "-NoProfile", "-Command", ps_script],
                    capture_output=True, timeout=15,
                )
                if tmp_cookies.exists() and tmp_cookies.stat().st_size > 0:
                    copied = True
            except Exception:
                pass

        # Cách 7: Win32 API CopyFile (share read)
        if not copied:
            try:
                import ctypes
                kernel32 = ctypes.windll.kernel32
                # CopyFileW(src, dst, failIfExists=False)
                result = kernel32.CopyFileW(str(cookies_path), str(tmp_cookies), False)
                if result and tmp_cookies.exists() and tmp_cookies.stat().st_size > 0:
                    copied = True
            except Exception:
                pass

        if not copied:
            raise PermissionError(
                f"Không thể đọc Cookies DB: {cookies_path}\n"
                "Chrome đang lock file quá chặt. Hãy thử:\n"
                "1. Đóng Chrome hoàn toàn (kiểm tra Task Manager)\n"
                "2. Hoặc dùng cách CDP (kết nối Chrome đang chạy) thay thế"
            )

        conn = sqlite3.connect(str(tmp_cookies), timeout=5)
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            "SELECT host_key, name, value, path, expires_utc, "
            "       is_secure, is_httponly, encrypted_value "
            "FROM cookies "
            "WHERE host_key LIKE '%.microsoft.com' "
            "   OR host_key LIKE '%.live.com' "
            "   OR host_key LIKE '%.microsoftonline.com'"
        ).fetchall()
        conn.close()
    finally:
        tmp_cookies.unlink(missing_ok=True)

    # ── Webkit epoch → Unix epoch (Chrome dùng microseconds từ 1601-01-01) ──
    WEBKIT_TO_UNIX = 11_644_473_600

    pw_cookies = []
    for row in rows:
        # Decrypt cookie value
        value = row["value"]
        enc   = row["encrypted_value"]
        if enc:
            try:
                if enc[:3] == b"v10" or enc[:3] == b"v11":
                    # AES-256-GCM với app-bound key
                    if encrypted_key:
                        nonce      = enc[3:15]
                        ciphertext = enc[15:]
                        cipher     = AES.new(encrypted_key, AES.MODE_GCM, nonce=nonce)
                        value      = cipher.decrypt(ciphertext)[:-16].decode("utf-8", errors="replace")
                else:
                    # DPAPI trực tiếp (Chrome cũ)
                    value = win32crypt.CryptUnprotectData(enc, None, None, None, 0)[1].decode("utf-8")
            except Exception:
                value = ""

        if not value:
            continue

        exp = row["expires_utc"]
        expires = (exp / 1_000_000 - WEBKIT_TO_UNIX) if exp else -1

        pw_cookies.append({
            "name":     row["name"],
            "value":    value,
            "domain":   row["host_key"],
            "path":     row["path"] or "/",
            "expires":  expires,
            "httpOnly": bool(row["is_httponly"]),
            "secure":   bool(row["is_secure"]),
            "sameSite": "None",
        })

    return pw_cookies


def _get_token_chrome_profile(profile_dir: str = None) -> str | None:
    """
    Decrypt cookies từ Chrome profile → inject vào Playwright context mới →
    mở Teams (tự SSO nhờ cookies) → bắt Graph API token.
    Cũng inject JS để extract token từ Teams Web nội bộ.
    """
    import shutil
    import tempfile
    from playwright.sync_api import sync_playwright  # noqa: PLC0415

    if profile_dir is None:
        profile_dir = os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data")

    # Decrypt cookies từ profile thật
    cookies = _decrypt_chrome_cookies(profile_dir)

    # Dùng thư mục tạm trống (không copy profile cũ → tránh DPAPI conflict)
    tmp_dir = Path(tempfile.mkdtemp(prefix="teams_chrome_"))
    captured = {"token": None}

    try:
        with sync_playwright() as p:
            context = p.chromium.launch_persistent_context(
                user_data_dir=str(tmp_dir),
                channel="chrome",
                headless=False,
                args=["--no-sandbox", "--disable-dev-shm-usage"],
                ignore_default_args=["--enable-automation"],
            )

            # Inject cookies vào context trước khi navigate
            if cookies:
                try:
                    context.add_cookies(cookies)
                except Exception:
                    pass  # Một số cookie có thể invalid → bỏ qua

            page = context.new_page()

            def on_request(req):
                auth = req.headers.get("authorization", "")
                if (
                    auth.startswith("Bearer ")
                    and ("graph.microsoft.com" in req.url
                         or "api.spaces.skype.com" in req.url
                         or "teams.microsoft.com/api" in req.url)
                    and not captured["token"]
                ):
                    token_val = auth.removeprefix("Bearer ")
                    if len(token_val) > 200:
                        captured["token"] = token_val

            def on_response(resp):
                """Bắt token từ auth endpoint responses."""
                if captured["token"]:
                    return
                url = resp.url
                if "login.microsoftonline.com" in url and "/oauth2" in url:
                    try:
                        body = resp.json()
                        if "access_token" in body:
                            captured["token"] = body["access_token"]
                    except Exception:
                        pass

            page.on("request", on_request)
            page.on("response", on_response)

            # Mở Teams — cookies đã inject → tự SSO
            page.goto("https://teams.microsoft.com", timeout=60_000)

            # Chờ trang load xong (nhận biết bằng nhiều selector khác nhau)
            try:
                page.wait_for_selector(
                    '[data-tid="team-channel-list"], [data-tid="chat-list"], '
                    '#app-bar-chat-button, [data-tid="app-bar-chat-button"], '
                    'button[title="Chat"], div[class*="leftRail"]',
                    timeout=45_000,
                )
            except Exception:
                pass

            # Chờ thêm cho Teams fetch data
            page.wait_for_timeout(5_000)

            # ── Cách 2: Inject JS để lấy token từ Teams internals ──────────
            if not captured["token"]:
                js_extract = """
                () => {
                    // Teams v2 lưu token trong sessionStorage / localStorage
                    const stores = [sessionStorage, localStorage];
                    for (const store of stores) {
                        for (let i = 0; i < store.length; i++) {
                            const key = store.key(i);
                            try {
                                const val = store.getItem(key);
                                // Token thường trong JSON objects
                                if (val && val.includes('accessToken')) {
                                    const obj = JSON.parse(val);
                                    // MSAL cache format
                                    if (obj.secret && obj.secret.length > 200) {
                                        return obj.secret;
                                    }
                                    if (obj.accessToken && obj.accessToken.length > 200) {
                                        return obj.accessToken;
                                    }
                                }
                                // Direct token value (eyJ...)
                                if (val && val.startsWith('eyJ') && val.length > 200) {
                                    return val;
                                }
                            } catch (e) {}
                        }
                    }
                    return null;
                }
                """
                try:
                    js_token = page.evaluate(js_extract)
                    if js_token and len(js_token) > 200:
                        captured["token"] = js_token
                except Exception:
                    pass

            # ── Cách 3: Navigate đến chat để trigger thêm API calls ─────────
            if not captured["token"]:
                try:
                    # Click vào Chat tab
                    chat_btn = page.locator(
                        '#app-bar-chat-button, [data-tid="app-bar-chat-button"], '
                        'button[title="Chat"], button[aria-label="Chat"]'
                    ).first
                    chat_btn.click(timeout=5_000)
                    page.wait_for_timeout(5_000)
                except Exception:
                    pass

            # ── Cách 4: Thử fetch Graph API trực tiếp từ page context ──────
            if not captured["token"]:
                try:
                    # Dùng fetch trong page context — browser sẽ tự gắn auth
                    result = page.evaluate("""
                    async () => {
                        try {
                            const resp = await fetch(
                                'https://graph.microsoft.com/v1.0/me',
                                { credentials: 'include' }
                            );
                            // Không quan trọng response — ta cần bắt request header
                            return resp.status;
                        } catch(e) {
                            return -1;
                        }
                    }
                    """)
                except Exception:
                    pass
                page.wait_for_timeout(2_000)

            context.close()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    return captured["token"]


@app.route("/debug/cdp_launch")
def debug_cdp_launch():
    """Test: Launch Chrome with CDP manually and check."""
    import subprocess
    import time
    import socket
    import urllib.request
    info = []

    chrome_exe = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    profile_dir = os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data")
    port = _find_free_port()
    info.append(f"Port: {port}")

    # Kill existing Chrome
    subprocess.run(["taskkill", "/F", "/IM", "chrome.exe"], capture_output=True, timeout=10)
    time.sleep(3)

    # Check port is free
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.5)
        result = s.connect_ex(("127.0.0.1", port))
        info.append(f"Port {port} before launch: {'LISTENING' if result == 0 else 'FREE'}")

    # Launch Chrome
    cmd = [chrome_exe, f"--remote-debugging-port={port}",
           f"--user-data-dir={profile_dir}", "--no-first-run",
           "https://teams.microsoft.com"]
    info.append(f"CMD: {' '.join(cmd)}")
    proc = subprocess.Popen(cmd)
    info.append(f"PID: {proc.pid}")

    # Wait and check
    for i in range(15):
        time.sleep(2)
        # Check port
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.5)
            result = s.connect_ex(("127.0.0.1", port))
        port_status = "LISTENING" if result == 0 else "FREE"

        # Check CDP
        cdp_ok = False
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/version", timeout=2) as r:
                cdp_ok = True
                data = r.read().decode()[:200]
                info.append(f"  {i*2+2}s: port={port_status}, CDP=✅ {data}")
                break
        except Exception as e:
            info.append(f"  {i*2+2}s: port={port_status}, CDP=❌ {type(e).__name__}: {e}")

        # Check netstat for port
        if i == 2:
            ns = subprocess.run(
                ["netstat", "-ano", "-p", "TCP"],
                capture_output=True, text=True, timeout=5,
            )
            port_lines = [l for l in ns.stdout.splitlines() if f":{port}" in l]
            info.append(f"  netstat for port {port}: {port_lines[:5]}")

    # Check chrome command line
    wmic = subprocess.run(
        ["wmic", "process", "where", "name='chrome.exe'", "get", "CommandLine"],
        capture_output=True, text=True, timeout=10,
    )
    chrome_cmds = [l.strip() for l in wmic.stdout.splitlines() if "remote-debugging" in l.lower()]
    info.append(f"\nChrome with debug flag: {len(chrome_cmds)}")
    for c in chrome_cmds[:3]:
        info.append(f"  {c[:200]}")

    if not chrome_cmds:
        info.append("\n⚠️ Chrome is running but NO process has --remote-debugging-port!")
        info.append("Possible causes:")
        info.append("  1. Group Policy blocks --remote-debugging-port")
        info.append("  2. Chrome reattached to existing instance (ignoring new flags)")
        info.append("  3. Antivirus blocked the flag")

    return "<pre>" + "\n".join(info) + "</pre>"


@app.route("/debug/cdp_test")
def debug_cdp_test():
    """Test CDP: tìm Chrome, tìm port, khởi động, kiểm tra."""
    import subprocess
    import socket
    info = []

    # 1. Tìm Chrome
    chrome_exe = None
    for candidate in [
        os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ]:
        exists = os.path.exists(candidate)
        info.append(f"  {candidate} → {'✅' if exists else '❌'}")
        if exists and not chrome_exe:
            chrome_exe = candidate
    info.insert(0, f"Chrome exe: {chrome_exe or 'NOT FOUND'}")

    # 2. Check Chrome processes
    r = subprocess.run(["tasklist", "/FI", "IMAGENAME eq chrome.exe"],
                       capture_output=True, text=True, timeout=5)
    chrome_count = r.stdout.lower().count("chrome.exe")
    info.append(f"\nChrome processes: {chrome_count}")

    # 3. Check ports 9222-9230
    info.append("\nPort scan:")
    for p in range(9222, 9231):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.5)
            result = s.connect_ex(("127.0.0.1", p))
            status = "LISTENING" if result == 0 else "FREE"
            info.append(f"  Port {p}: {status}")

    # 4. Check CDP on common ports
    import urllib.request
    info.append("\nCDP check:")
    for p in range(9222, 9231):
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{p}/json/version", timeout=1) as resp:
                data = resp.read().decode()
                info.append(f"  Port {p}: ✅ CDP active — {data[:100]}")
        except Exception as e:
            info.append(f"  Port {p}: ❌ {type(e).__name__}")

    # 5. Free port
    try:
        fp = _find_free_port()
        info.append(f"\nFree port found: {fp}")
    except Exception as e:
        info.append(f"\nFree port error: {e}")

    return "<pre>" + "\n".join(info) + "</pre>"


@app.route("/debug/chrome_profile")
def debug_chrome_profile():
    """Tạm thời — kiểm tra Chrome profile location và Cookies DB."""
    import sqlite3 as _sql
    import shutil
    import subprocess
    import tempfile
    info = []
    profile_dir = os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data")
    info.append(f"Chrome dir: {profile_dir}")
    info.append(f"Exists: {os.path.exists(profile_dir)}")

    cookies_path = os.path.join(profile_dir, "Default", "Network", "Cookies")
    if not os.path.exists(cookies_path):
        cookies_path = os.path.join(profile_dir, "Default", "Cookies")
    info.append(f"\nCookies path: {cookies_path}")
    info.append(f"Cookies exists: {os.path.exists(cookies_path)}")
    if os.path.exists(cookies_path):
        info.append(f"Cookies size: {os.path.getsize(cookies_path)} bytes")

    tmp = os.path.join(tempfile.gettempdir(), "test_cookies_copy.db")

    # Cleanup trước
    if os.path.exists(tmp):
        try: os.unlink(tmp)
        except: pass

    # Cách 1: shutil.copy2
    try:
        shutil.copy2(cookies_path, tmp)
        info.append("\n✅ Cách 1 (shutil.copy2): OK")
        os.unlink(tmp)
    except Exception as e:
        info.append(f"\n❌ Cách 1 (shutil.copy2): {e}")

    # Cách 2: esentutl /y (Windows built-in, bypass exclusive lock!)
    try:
        r = subprocess.run(
            ["esentutl", "/y", cookies_path, "/d", tmp],
            capture_output=True, text=True, timeout=15,
        )
        if os.path.exists(tmp) and os.path.getsize(tmp) > 0:
            info.append(f"✅ Cách 2 (esentutl): OK ({os.path.getsize(tmp)} bytes)")
            os.unlink(tmp)
        else:
            info.append(f"❌ Cách 2 (esentutl): Failed. RC={r.returncode}\n   stdout={r.stdout[:200]}\n   stderr={r.stderr[:200]}")
    except Exception as e:
        info.append(f"❌ Cách 2 (esentutl): {e}")

    # Cách 3: sqlite3 backup with URI (corrected URI format)
    try:
        uri_path = cookies_path.replace("\\", "/")
        src = _sql.connect(f"file:///{uri_path}?mode=ro&nolock=1", uri=True, timeout=3)
        dst = _sql.connect(tmp)
        src.backup(dst)
        src.close()
        dst.close()
        info.append(f"✅ Cách 3 (sqlite3 URI backup): OK")
        os.unlink(tmp)
    except Exception as e:
        info.append(f"❌ Cách 3 (sqlite3 URI backup): {e}")

    # Cách 4: raw bytes
    try:
        with open(cookies_path, "rb") as f:
            data = f.read()
        with open(tmp, "wb") as f:
            f.write(data)
        info.append(f"✅ Cách 4 (raw bytes): OK ({len(data)} bytes)")
        os.unlink(tmp)
    except Exception as e:
        info.append(f"❌ Cách 4 (raw bytes): {e}")

    # Cách 5: Win32 CopyFileW
    try:
        import ctypes
        result = ctypes.windll.kernel32.CopyFileW(cookies_path, tmp, False)
        if result and os.path.exists(tmp) and os.path.getsize(tmp) > 0:
            info.append(f"✅ Cách 5 (Win32 CopyFileW): OK")
            os.unlink(tmp)
        else:
            err_code = ctypes.GetLastError()
            info.append(f"❌ Cách 5 (Win32 CopyFileW): Failed. GetLastError={err_code}")
    except Exception as e:
        info.append(f"❌ Cách 5 (Win32 CopyFileW): {e}")

    # Cách 6: PowerShell .NET FileStream (share ReadWrite)
    try:
        ps_script = f"""
$src = [System.IO.FileStream]::new('{cookies_path}', 'Open', 'Read', 'ReadWrite,Delete')
$dst = [System.IO.FileStream]::new('{tmp}', 'Create', 'Write')
$src.CopyTo($dst)
$src.Close()
$dst.Close()
Write-Host 'OK'
"""
        r = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_script],
            capture_output=True, text=True, timeout=15,
        )
        if os.path.exists(tmp) and os.path.getsize(tmp) > 0:
            info.append(f"✅ Cách 6 (PS FileStream share): OK ({os.path.getsize(tmp)} bytes)")
            os.unlink(tmp)
        else:
            info.append(f"❌ Cách 6 (PS FileStream share): {r.stderr[:300]}")
    except Exception as e:
        info.append(f"❌ Cách 6 (PS FileStream share): {e}")

    # Test decrypt nếu có cách nào thành công
    info.append("\n--- Test Decrypt ---")
    try:
        cookies = _decrypt_chrome_cookies(profile_dir)
        ms_cookies = [c for c in cookies if "microsoft" in c.get("domain", "").lower()
                      or "live.com" in c.get("domain", "").lower()]
        info.append(f"✅ Total cookies decrypted: {len(cookies)}")
        info.append(f"✅ Microsoft/Live cookies: {len(ms_cookies)}")
        for c in ms_cookies[:10]:
            info.append(f"   {c['domain']} | {c['name']} = {c['value'][:30]}...")
    except Exception as e:
        import traceback
        info.append(f"❌ Decrypt error: {e}")
        info.append(traceback.format_exc())

    return "<pre>" + "\n".join(info) + "</pre>"


_cdp_jobs: dict[str, dict] = {}

# Thư mục lưu session Playwright (persist giữa các lần chạy)
_PW_SESSION_DIR = Path(__file__).resolve().parent.parent / ".pw_session"


def _get_token_playwright_login(job: dict = None) -> str | None:
    """
    Mở Playwright Chromium (bundled — KHÔNG phải Chrome) → user đăng nhập Teams
    → bắt token. Session được lưu lại cho lần sau (không cần login lại).
    Không bị Group Policy của Chrome ảnh hưởng.
    """
    import time
    from playwright.sync_api import sync_playwright

    def log(msg):
        if job:
            job["log"].append(msg)
        print(f"[PW] {msg}")

    captured = {"token": None}
    session_dir = str(_PW_SESSION_DIR)

    with sync_playwright() as p:
        log("🌐 Mở Playwright Chromium…")

        context = p.chromium.launch_persistent_context(
            user_data_dir=session_dir,
            headless=False,
            args=["--no-sandbox", "--disable-dev-shm-usage"],
            ignore_default_args=["--enable-automation"],
            viewport={"width": 1280, "height": 900},
        )

        page = context.pages[0] if context.pages else context.new_page()

        def on_request(req):
            if captured.get("source") in ("graph_request", "graph_with_chat"):
                return
            url = req.url
            headers = req.headers
            auth = headers.get("authorization", "")

            # --- 1) Graph Bearer token ---
            if auth.startswith("Bearer "):
                token_val = auth.removeprefix("Bearer ")
                if len(token_val) > 200:
                    if "graph.microsoft.com" in url:
                        captured["token"] = token_val
                        captured["source"] = "graph_request"
                        log(f"✅ Graph token bắt được! ({url[:60]}…)")
                    elif not captured.get("fallback"):
                        captured["fallback"] = token_val
                        log(f"📌 Non-Graph token: {url[:60]}…")

            # --- 2) Teams/Skype Bearer token (chat service) ---
            if auth.startswith("Bearer "):
                token_val = auth.removeprefix("Bearer ")
                if len(token_val) > 200 and not captured.get("teams_token"):
                    teams_domains = [
                        "msg.teams.microsoft.com",
                        "chatsvcagg.teams.microsoft.com",
                        "teams.microsoft.com/api/csa",
                        "teams.microsoft.com/api/mt",
                    ]
                    if any(d in url for d in teams_domains):
                        captured["teams_token"] = token_val
                        captured["teams_token_url"] = url
                        log(f"📌 Teams chat Bearer token bắt được! ({url[:60]}…)")

            # --- 3) skypetoken= auth scheme (Teams internal chat API) ---
            if auth.lower().startswith("skypetoken=") and not captured.get("skype_token"):
                token_val = auth.split("=", 1)[1].strip()
                if len(token_val) > 50:
                    captured["skype_token"] = token_val
                    log(f"📌 Skype token (auth header) bắt được! ({url[:60]}…)")

            # --- 4) x-skypetoken header ---
            x_skype = headers.get("x-skypetoken", "")
            if x_skype and len(x_skype) > 50 and not captured.get("skype_token"):
                captured["skype_token"] = x_skype
                log(f"📌 Skype token (x-skypetoken) bắt được! ({url[:60]}…)")

            # --- 5) Bearer on Skype/msgapi domains ---
            if auth.startswith("Bearer "):
                token_val = auth.removeprefix("Bearer ")
                if len(token_val) > 200 and not captured.get("skype_token"):
                    skype_domains = [
                        "api.spaces.skype.com",
                        "msgapi.teams.live.com",
                    ]
                    if any(d in url for d in skype_domains):
                        captured["skype_token"] = token_val
                        log(f"📌 Skype Bearer token bắt được! ({url[:60]}…)")

            # --- 6) Log tất cả requests đến chat-related domains (debug) ---
            chat_debug_domains = [
                "msg.teams.microsoft.com",
                "chatsvcagg.teams.microsoft.com",
                "api.spaces.skype.com",
                "msgapi.teams.live.com",
            ]
            if any(d in url for d in chat_debug_domains):
                auth_type = "none"
                if auth.startswith("Bearer "):
                    auth_type = "Bearer"
                elif auth.lower().startswith("skypetoken"):
                    auth_type = "skypetoken"
                if x_skype:
                    auth_type += "+x-skypetoken"
                if not captured.get("_debug_chat_logged"):
                    captured["_debug_chat_logged"] = True
                    log(f"🔍 Chat domain request: auth={auth_type}, url={url[:80]}…")

        def on_response(resp):
            # --- Bắt Skype token từ Teams authsvc response ---
            try:
                url = resp.url
                # Teams fetches skypetoken via authsvc or api/mt/emea
                if ("authsvc" in url or "api/mt" in url or "api/csa" in url) and "teams.microsoft.com" in url:
                    if resp.status == 200:
                        try:
                            body = resp.json()
                            # authsvc response contains { "tokens": { "skypeToken": "...", ... } }
                            tokens_obj = body.get("tokens", {})
                            skype_val = tokens_obj.get("skypeToken", "") or body.get("skypeToken", "") or body.get("skype_token", "")
                            if skype_val and len(skype_val) > 50 and not captured.get("skype_token"):
                                captured["skype_token"] = skype_val
                                log(f"📌 Skype token từ Teams authsvc response!")
                            # Also check for chatSvcAggToken
                            chat_svc = tokens_obj.get("chatSvcAggToken", "") or body.get("chatSvcAggToken", "")
                            if chat_svc and len(chat_svc) > 50 and not captured.get("teams_token"):
                                captured["teams_token"] = chat_svc
                                captured["teams_token_url"] = url
                                log(f"📌 ChatSvcAgg token từ Teams authsvc response!")
                        except Exception:
                            pass
            except Exception:
                pass

            if captured.get("source") in ("graph_request", "graph_with_chat"):
                return
            if "login.microsoftonline.com" in resp.url and "/oauth2" in resp.url:
                try:
                    body = resp.json()
                    if "access_token" in body:
                        token_val = body["access_token"]
                        # Decode để check audience
                        try:
                            import base64 as _b64
                            parts = token_val.split(".")
                            if len(parts) >= 2:
                                padded = parts[1] + "=" * (-len(parts[1]) % 4)
                                payload = json.loads(_b64.urlsafe_b64decode(padded))
                                aud = payload.get("aud", "")
                                if "graph.microsoft.com" in aud or "00000003-0000-0000-c000-000000000000" in aud:
                                    captured["token"] = token_val
                                    captured["source"] = "oauth_graph"
                                    log(f"✅ Graph token từ OAuth response! (aud={aud})")
                                    return
                                else:
                                    log(f"📌 OAuth token bỏ qua (aud={aud[:50]})")
                        except Exception:
                            pass
                        # Nếu không decode được, lưu fallback
                        if not captured.get("fallback"):
                            captured["fallback"] = token_val
                except Exception:
                    pass

        page.on("request", on_request)
        page.on("response", on_response)

        log("📄 Navigate đến Teams…")
        page.goto("https://teams.microsoft.com", timeout=120_000)

        # Chờ Teams load hoặc login page
        log("⏳ Chờ Teams load (nếu cần đăng nhập, hãy đăng nhập trên cửa sổ vừa mở)…")

        # Chờ tối đa 120 giây cho user đăng nhập + Teams load
        for i in range(60):
            if captured.get("source") in ("graph_request", "graph_with_chat"):
                break
            time.sleep(2)
            if i % 5 == 4:
                current_url = page.url
                log(f"  ⏳ {(i+1)*2}s — {current_url[:80]}")

            # Kiểm tra xem Teams đã load chưa (hoặc user đã login xong)
            try:
                # Kiểm tra Teams UI loaded
                is_teams = page.evaluate("""
                () => {
                    return window.location.hostname === 'teams.microsoft.com'
                        && !window.location.pathname.includes('/login')
                        && !window.location.pathname.includes('/common/oauth2')
                        && (document.querySelector('[data-tid="team-channel-list"]') !== null
                            || document.querySelector('[data-tid="chat-list"]') !== null
                            || document.querySelector('#app-bar-chat-button') !== null
                            || document.querySelector('button[title="Chat"]') !== null
                            || document.querySelector('app-bar') !== null
                            || document.querySelector('[class*="app-bar"]') !== null
                            || document.querySelector('[data-tid]') !== null);
                }
                """)
                if is_teams and not captured.get("token"):
                    log("✅ Teams đã load! Thử mở Chat tab để bắt chat token…")
                    # Click vào Chat tab để trigger chat API requests
                    try:
                        chat_btn = page.query_selector(
                            '#app-bar-chat-button, button[title="Chat"], '
                            '[data-tid="app-bar-chat-button"], '
                            'button[data-tid="app-bar-2"], '
                            '[aria-label="Chat"], '
                            'button:has-text("Chat")'
                        )
                        if chat_btn:
                            chat_btn.click()
                            log("📌 Đã click Chat tab, chờ chat API requests…")
                            # Chờ tối đa 15s cho chat API tokens
                            for _w in range(15):
                                time.sleep(1)
                                if captured.get("teams_token") or captured.get("skype_token"):
                                    log("✅ Đã bắt được chat token!")
                                    break
                            if not captured.get("teams_token") and not captured.get("skype_token"):
                                log("⚠️ Chưa bắt được chat token từ Chat tab, tiếp tục…")
                        else:
                            log("⚠️ Không tìm thấy Chat button, tiếp tục…")
                    except Exception as e_chat:
                        log(f"⚠️ Lỗi khi click Chat: {e_chat}")
                    log("🔄 Chuyển sang Graph Explorer để lấy Graph token…")
                    break  # Chuyển sang phase 2
            except Exception:
                pass

        # ── Phase 2: Graph Explorer → consent Chat.Read → bắt token ─────────
        # Strategy:
        #   A) Mở GE, chờ load, bắt Graph token từ auto-requests
        #   B) Nếu thiếu Chat.Read → dùng Playwright click "Modify permissions"
        #      → consent Chat.Read → chạy query /me/chats → bắt token mới
        #   C) Fallback — Auth Code + PKCE với prompt=consent
        if not captured.get("token") or captured.get("source") != "graph_with_chat":
            log("🔄 Mở Graph Explorer để lấy Graph API token có Chat.Read…")
            try:
                from urllib.parse import urlparse, parse_qs, quote
                import base64 as _b64_ge
                graph_page = context.new_page()

                # Helper: decode JWT token → (audience, scopes, has_chat_read)
                def _decode_token_scopes(token_val):
                    try:
                        parts = token_val.split(".")
                        padded = parts[1] + "=" * (-len(parts[1]) % 4)
                        payload = json.loads(_b64_ge.urlsafe_b64decode(padded))
                        aud = payload.get("aud", "")
                        scp = payload.get("scp", "")
                        has_chat = any(s in scp.lower() for s in ["chat.read", "chat.readbasic"])
                        return aud, scp, has_chat
                    except Exception:
                        return "", "", False

                # Intercept Graph requests trên graph_page
                def on_graph_page_request(req):
                    if captured.get("source") == "graph_with_chat":
                        return  # Đã có token tốt nhất
                    auth = req.headers.get("authorization", "")
                    if auth.startswith("Bearer ") and "graph.microsoft.com" in req.url:
                        token_val = auth.removeprefix("Bearer ")
                        if len(token_val) > 200:
                            _, _, has_chat = _decode_token_scopes(token_val)
                            if has_chat:
                                captured["token"] = token_val
                                captured["source"] = "graph_with_chat"
                                log(f"✅ Graph token VỚI Chat.Read! ({req.url[:60]}…)")
                            elif captured.get("source") != "graph_request":
                                captured["token"] = token_val
                                captured["source"] = "graph_request"
                                log(f"📌 Graph token (chưa có Chat.Read): {req.url[:60]}…")

                graph_page.on("request", on_graph_page_request)

                # Step A: Mở Graph Explorer
                log("📄 Navigate đến Graph Explorer…")
                graph_page.goto(
                    "https://developer.microsoft.com/en-us/graph/graph-explorer",
                    timeout=60_000,
                )

                # Chờ Graph Explorer load + auto-SSO (tối đa 40s)
                log("⏳ Chờ Graph Explorer load & auto-SSO…")
                for gi in range(20):
                    if captured.get("source") == "graph_with_chat":
                        break
                    time.sleep(2)
                    if gi % 5 == 4:
                        log(f"  ⏳ GE: {(gi+1)*2}s…")

                # Step B: Nếu có Graph token nhưng thiếu Chat.Read → consent qua GE UI
                if captured.get("source") == "graph_request":
                    log("🔑 Có Graph token nhưng thiếu Chat.Read — consent qua GE UI…")

                    # B1: Click "Modify permissions" tab
                    try:
                        perms_tab = graph_page.locator(
                            'button:has-text("Modify permissions"), '
                            '[role="tab"]:has-text("Modify permissions"), '
                            'button:has-text("Permissions")'
                        )
                        if perms_tab.count() > 0:
                            perms_tab.first.click()
                            log("📋 Clicked 'Modify permissions' tab")
                            time.sleep(3)

                            # Find Chat.Read row and its Consent button
                            chat_row = graph_page.locator('tr:has-text("Chat.Read"), div:has-text("Chat.Read")')
                            if chat_row.count() > 0:
                                consent_btn = chat_row.first.locator('button:has-text("Consent")')
                                if consent_btn.count() > 0:
                                    log("🔓 Found Chat.Read consent button — clicking…")
                                    consent_btn.first.click()
                                    time.sleep(3)
                                    for ci in range(15):
                                        if captured.get("source") == "graph_with_chat":
                                            break
                                        time.sleep(2)
                                        if ci % 5 == 4:
                                            log(f"  ⏳ Consent: {(ci+1)*2}s…")
                                else:
                                    log("  ℹ️ Chat.Read found but no Consent button (already consented?)")
                            else:
                                any_consent = graph_page.locator('button:has-text("Consent")')
                                if any_consent.count() > 0:
                                    log(f"  📋 {any_consent.count()} Consent button(s) found, clicking first…")
                                    any_consent.first.click()
                                    time.sleep(5)
                                else:
                                    log("  ℹ️ No Chat.Read or Consent buttons found")
                        else:
                            log("  ℹ️ 'Modify permissions' tab not found")
                    except Exception as e:
                        log(f"  ⚠️ GE UI consent: {e}")

                # B2: set query /me/chats rồi Run
                if captured.get("source") != "graph_with_chat":
                    try:
                        log("📝 Thử query /me/chats…")
                        query_input = graph_page.locator(
                            'input[aria-label*="request URL"], '
                            'input[placeholder*="/me"], '
                            'input[type="url"], '
                            'input[role="combobox"]'
                        )
                        if query_input.count() > 0:
                            query_input.first.fill("https://graph.microsoft.com/v1.0/me/chats")
                            time.sleep(1)
                            run_btn = graph_page.locator(
                                'button:has-text("Run query"), button[aria-label*="Run"]'
                            )
                            if run_btn.count() > 0:
                                run_btn.first.click()
                                log("▶️ Clicked Run query for /me/chats")
                                time.sleep(5)
                    except Exception as e:
                        log(f"  ⚠️ Run query: {e}")

                # Step C: Auth Code + PKCE fallback
                if captured.get("source") != "graph_with_chat":
                    log("🔄 Auth Code + PKCE with Chat.Read + consent…")
                    try:
                        import hashlib
                        import secrets

                        code_verifier = secrets.token_urlsafe(64)
                        code_challenge = _b64_ge.urlsafe_b64encode(
                            hashlib.sha256(code_verifier.encode()).digest()
                        ).rstrip(b"=").decode()

                        GE_CLIENT_ID = "de8bc8b5-d9f9-48b1-a8ad-b748da725064"
                        REDIRECT_URI = "https://developer.microsoft.com/en-us/graph/graph-explorer"
                        SCOPES = "User.Read Chat.Read Chat.ReadBasic Team.ReadBasic.All Channel.ReadBasic.All ChannelMessage.Read.All openid profile offline_access"
                        auth_url = (
                            f"https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
                            f"?client_id={GE_CLIENT_ID}"
                            f"&response_type=code"
                            f"&redirect_uri={quote(REDIRECT_URI, safe='')}"
                            f"&scope={quote(SCOPES, safe='')}"
                            f"&code_challenge={code_challenge}"
                            f"&code_challenge_method=S256"
                            f"&prompt=consent"
                        )
                        log("📄 Consent page — bấm Accept nếu thấy popup…")
                        consent_page = context.new_page()
                        consent_page.goto(auth_url, timeout=60_000, wait_until="commit")

                        for wi in range(45):
                            time.sleep(2)
                            cur_url = consent_page.url
                            if "code=" in cur_url:
                                qs = parse_qs(urlparse(cur_url).query)
                                auth_code = qs.get("code", [""])[0]
                                if auth_code:
                                    log("📦 Got auth code, exchanging…")
                                    import requests as _req
                                    token_resp = _req.post(
                                        "https://login.microsoftonline.com/common/oauth2/v2.0/token",
                                        data={
                                            "client_id": GE_CLIENT_ID,
                                            "grant_type": "authorization_code",
                                            "code": auth_code,
                                            "redirect_uri": REDIRECT_URI,
                                            "code_verifier": code_verifier,
                                            "scope": SCOPES,
                                        },
                                        timeout=30,
                                    )
                                    if token_resp.status_code == 200:
                                        td = token_resp.json()
                                        captured["token"] = td["access_token"]
                                        captured["source"] = "graph_with_chat"
                                        log("✅ Graph token VỚI Chat.Read!")
                                    else:
                                        log(f"  ❌ Exchange: {token_resp.status_code} — {token_resp.text[:200]}")
                                break
                            if "error=" in cur_url:
                                frag = urlparse(cur_url).fragment
                                q_e = parse_qs(urlparse(cur_url).query)
                                f_e = parse_qs(frag) if frag else {}
                                err = q_e.get("error", f_e.get("error", [""]))[0]
                                desc = q_e.get("error_description", f_e.get("error_description", [""]))[0][:150]
                                log(f"  ❌ {err} — {desc}")
                                break
                            if wi % 5 == 4:
                                log(f"  ⏳ {(wi+1)*2}s — {cur_url[:80]}")
                        try:
                            consent_page.close()
                        except Exception:
                            pass
                    except Exception as e:
                        log(f"  ⚠️ Auth Code: {e}")

                # Final status
                if captured.get("source") == "graph_with_chat":
                    log("🎉 Graph token VỚI Chat.Read!")
                elif captured.get("source") == "graph_request":
                    _, scp, _ = _decode_token_scopes(captured.get("token", ""))
                    log(f"⚠️ Graph token nhưng KHÔNG có Chat.Read. Scopes: {scp[:120]}")
                    log("💡 Tip: Vào graph-explorer → Modify permissions → consent Chat.Read")

                try:
                    graph_page.close()
                except Exception:
                    pass
            except Exception as e:
                log(f"⚠️ Graph Explorer flow: {e}")

        # ── Phase 3: Fallback — thử lấy token bằng JS trong Teams page ───────
        if not captured.get("token") or captured.get("source") not in ("graph_request", "graph_with_chat"):
            log("� Thử lấy Graph token từ MSAL cache trong Teams…")
            try:
                graph_token_from_cache = page.evaluate("""
                () => {
                    // Teams new client uses MSAL — tokens are in sessionStorage/localStorage
                    // Keys look like: {"homeAccountId"-"login.microsoftonline.com"-"accesstoken"-"..."-"https://graph.microsoft.com/..."}
                    const stores = [sessionStorage, localStorage];
                    const graphAudiences = ['graph.microsoft.com', '00000003-0000-0000-c000-000000000000'];

                    // Pass 1: Look for MSAL access token cache entries for Graph
                    for (const store of stores) {
                        for (let i = 0; i < store.length; i++) {
                            const key = store.key(i);
                            const keyLower = key.toLowerCase();
                            if (keyLower.includes('accesstoken')) {
                                try {
                                    const val = store.getItem(key);
                                    const obj = JSON.parse(val);
                                    if (obj && obj.secret && obj.secret.length > 200) {
                                        const target = (obj.target || '').toLowerCase();
                                        const realm  = (obj.realm || '').toLowerCase();
                                        const env    = (obj.environment || '').toLowerCase();
                                        // Check if this token has Graph scopes
                                        if (target.includes('user.read') || target.includes('chat.read')
                                            || target.includes('channelmessage') || target.includes('team.readbasic')
                                            || target.includes('group.read')
                                            || keyLower.includes('graph.microsoft.com')
                                            || keyLower.includes('00000003-0000-0000-c000-000000000000')) {
                                            return { token: obj.secret, source: 'msal_graph', key: key.substring(0, 80) };
                                        }
                                    }
                                } catch (e) {}
                            }
                        }
                    }

                    // Pass 2: Any MSAL token (fallback)
                    for (const store of stores) {
                        for (let i = 0; i < store.length; i++) {
                            const key = store.key(i);
                            if (key.toLowerCase().includes('accesstoken')) {
                                try {
                                    const val = store.getItem(key);
                                    const obj = JSON.parse(val);
                                    if (obj && obj.secret && obj.secret.length > 200) {
                                        return { token: obj.secret, source: 'msal_any', key: key.substring(0, 80) };
                                    }
                                } catch (e) {}
                            }
                        }
                    }

                    // Pass 3: Dump all storage keys for debugging
                    const allKeys = [];
                    for (const store of stores) {
                        for (let i = 0; i < store.length; i++) {
                            const key = store.key(i);
                            const val = store.getItem(key) || '';
                            allKeys.push(key.substring(0, 100) + ' (' + val.length + ' chars)');
                        }
                    }
                    return { token: null, source: 'none', keys: allKeys.slice(0, 30) };
                }
                """)

                if graph_token_from_cache and graph_token_from_cache.get("token"):
                    tok = graph_token_from_cache["token"]
                    src = graph_token_from_cache.get("source", "?")
                    key_info = graph_token_from_cache.get("key", "")
                    log(f"✅ Token từ MSAL cache ({src}): {key_info}")
                    if not captured.get("token") or src == "msal_graph":
                        captured["token"] = tok
                        captured["source"] = src
                else:
                    keys_preview = graph_token_from_cache.get("keys", []) if graph_token_from_cache else []
                    log(f"⚠️ Không tìm thấy token trong MSAL cache. Storage keys ({len(keys_preview)}):")
                    for k in keys_preview[:15]:
                        log(f"  📦 {k}")
            except Exception as e:
                log(f"⚠️ MSAL cache scan: {e}")

        # Nếu chưa có Graph token nhưng có fallback token
        if not captured.get("token") and captured.get("fallback"):
            captured["token"] = captured["fallback"]
            log("📌 Dùng backup token (Teams API — không phải Graph API, có thể không hoạt động với Graph endpoints)")

        # ── Validate token audience ───────────────────────────────────────────
        if captured.get("token"):
            try:
                import base64
                parts = captured["token"].split(".")
                if len(parts) >= 2:
                    padded = parts[1] + "=" * (-len(parts[1]) % 4)
                    payload = json.loads(base64.urlsafe_b64decode(padded))
                    aud = payload.get("aud", "")
                    scp = payload.get("scp", "")
                    log(f"🔍 Token audience: {aud}")
                    log(f"🔍 Token scopes: {scp[:100]}")
                    if "graph.microsoft.com" not in aud and "00000003-0000-0000-c000-000000000000" not in aud:
                        log(f"⚠️ Token KHÔNG phải cho Graph API (audience={aud})")
                        log("⚠️ Sẽ thử dùng nhưng có thể bị 401 khi gọi Graph endpoints.")
            except Exception:
                pass

        # Lưu session (cookies, localStorage, etc.) cho lần sau
        context.close()
        log("💾 Session đã được lưu cho lần sau.")

        # Lưu teams_token riêng (cho Teams internal chat API)
        if captured.get("teams_token") or captured.get("skype_token"):
            try:
                teams_token_file = BASE_DIR / "teams_token.json"
                teams_data = {
                    "fetched_at": datetime.now(timezone.utc).isoformat(),
                }
                if captured.get("teams_token"):
                    teams_data["token"] = captured["teams_token"]
                    teams_data["token_url"] = captured.get("teams_token_url", "")
                if captured.get("skype_token"):
                    teams_data["skype_token"] = captured["skype_token"]
                teams_token_file.write_text(
                    json.dumps(teams_data, ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )
                log(f"💾 Teams chat token(s) saved to teams_token.json")
            except Exception as e:
                log(f"⚠️ Lưu teams_token: {e}")

    return captured["token"]


# ── Playwright Chat Scraping ─────────────────────────────────────────────────
# Dùng Playwright để mở Teams → Chat tab → intercept response chứa danh sách
# conversations. Teams tự handle auth nên KHÔNG cần token/scope nào cả.
# ─────────────────────────────────────────────────────────────────────────────

_pw_chat_jobs: dict[str, dict] = {}


def _scrape_chats_via_playwright(job: dict | None = None):
    """
    Mở Teams trong Playwright, navigate đến Chat tab, intercept response
    chứa danh sách conversations/chats. Trả về list[dict].
    """
    from playwright.sync_api import sync_playwright

    def log(msg):
        print(f"[PW-Chat] {msg}", flush=True)
        if job:
            job["log"].append(msg)

    chats_data = []
    conversations_raw = []
    chat_list_captured = threading.Event()
    captured_tokens = {}   # { "chatsvc_bearer": "...", "chatsvc_url_base": "..." }

    log("🌐 Mở Playwright Chromium để lấy chats từ Teams…")

    with sync_playwright() as pw:
        context = pw.chromium.launch_persistent_context(
            user_data_dir=str(_PW_SESSION_DIR),
            headless=False,
            channel="chromium",
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-first-run",
                "--no-default-browser-check",
            ],
            ignore_default_args=["--enable-automation"],
            viewport={"width": 1280, "height": 900},
        )

        page = context.pages[0] if context.pages else context.new_page()

        def on_chat_response(resp):
            """Intercept Teams API responses chứa chat/conversation data."""
            url = resp.url
            status = resp.status

            # Log mọi chatsvc response
            if "chatsvc" in url:
                log(f"📡 chatsvc response: HTTP {status} {url[:100]}…")

            try:
                # ── Pattern 1: chatsvc responses ──
                if "chatsvc" in url and status == 200:
                    try:
                        body_text = resp.text()
                        if body_text:
                            body = json.loads(body_text)
                            # List response: { conversations: [...] }
                            convs = body.get("conversations", [])
                            if convs:
                                conversations_raw.extend(convs)
                                log(f"📌 Intercepted {len(convs)} conversations (list) từ chatsvc!")
                                chat_list_captured.set()
                            # Single conversation: has "id" field
                            elif body.get("id"):
                                conversations_raw.append(body)
                                log(f"📌 Intercepted single conv: {body.get('id', '')[:40]}… keys={list(body.keys())[:6]}")
                            else:
                                # Log keys for debugging
                                log(f"📡 chatsvc body keys: {list(body.keys())[:10]}")
                    except Exception as e:
                        log(f"⚠️ chatsvc parse error: {e}")

                # ── Pattern 2: chatsvcagg / msg.teams.microsoft.com ──
                if ("chatsvcagg" in url or "msg.teams.microsoft.com" in url) and "/conversations" in url:
                    if resp.status == 200:
                        try:
                            body = resp.json()
                            convs = body.get("conversations", [])
                            if convs:
                                conversations_raw.extend(convs)
                                log(f"📌 Intercepted {len(convs)} conversations từ {url[:70]}…")
                                chat_list_captured.set()
                        except Exception:
                            pass

                # ── Pattern 3: teams.microsoft.com generic ──
                if "teams.microsoft.com" in url and "/conversations" in url and "/messages" not in url:
                    if resp.status == 200:
                        try:
                            body = resp.json()
                            convs = body.get("conversations", []) or body.get("chats", [])
                            if isinstance(body, list):
                                convs = body
                            if convs and len(convs) > 1:
                                conversations_raw.extend(convs)
                                log(f"📌 Intercepted {len(convs)} chats từ {url[:70]}…")
                                chat_list_captured.set()
                        except Exception:
                            pass

                # ── Pattern 4: Graph API /me/chats ──
                if "graph.microsoft.com" in url and "/chats" in url and "/messages" not in url:
                    if resp.status == 200:
                        try:
                            body = resp.json()
                            items = body.get("value", [])
                            if items:
                                for item in items:
                                    conversations_raw.append({
                                        "_graph_chat": True,
                                        "id": item.get("id", ""),
                                        "chatType": item.get("chatType", ""),
                                        "topic": item.get("topic", ""),
                                        "members": item.get("members", []),
                                    })
                                log(f"📌 Intercepted {len(items)} Graph chats")
                                chat_list_captured.set()
                        except Exception:
                            pass

                # ── Pattern 5: Teams /api/mt chat endpoints ──
                if "teams.microsoft.com/api/mt" in url and ("chat" in url.lower() or "conversation" in url.lower()):
                    if resp.status == 200:
                        try:
                            body = resp.json()
                            if isinstance(body, dict):
                                convs = body.get("conversations", []) or body.get("chats", []) or body.get("value", [])
                                if convs:
                                    conversations_raw.extend(convs)
                                    log(f"📌 Intercepted {len(convs)} items từ MT API")
                                    chat_list_captured.set()
                        except Exception:
                            pass
            except Exception:
                pass

        def on_chat_request(req):
            """Capture Bearer token từ chatsvc requests + log."""
            url = req.url
            auth = req.headers.get("authorization", "")

            # Capture chatsvc Bearer token
            if "chatsvc" in url and auth.startswith("Bearer ") and not captured_tokens.get("chatsvc_bearer"):
                captured_tokens["chatsvc_bearer"] = auth.removeprefix("Bearer ").strip()
                # Extract base URL: https://teams.microsoft.com/api/chatsvc/apac/v1
                import re as _re
                m = _re.search(r"(https://teams\.microsoft\.com/api/chatsvc/[^/]+/v\d+)", url)
                if m:
                    captured_tokens["chatsvc_url_base"] = m.group(1)
                log(f"📌 Captured chatsvc Bearer token! base={captured_tokens.get('chatsvc_url_base', '?')}")

            # Log chat-related requests
            if any(kw in url.lower() for kw in ["conversation", "/chats", "chatsvc", "recentchat"]):
                auth_type = "none"
                if "skypetoken" in auth.lower():
                    auth_type = "skypetoken"
                elif auth.startswith("Bearer "):
                    auth_type = "Bearer"
                log(f"🔍 {req.method} {url[:100]}… (auth={auth_type})")

        page.on("response", on_chat_response)
        page.on("request", on_chat_request)

        log("📄 Navigate đến Teams…")
        try:
            page.goto("https://teams.microsoft.com", timeout=60_000)
        except Exception as e:
            log(f"⚠️ Teams navigate timeout (OK nếu Teams đã load): {e}")

        # Chờ Teams UI load
        log("⏳ Chờ Teams load…")
        teams_loaded = False
        for i in range(45):
            time.sleep(2)
            try:
                is_teams = page.evaluate("""
                () => {
                    return window.location.hostname === 'teams.microsoft.com'
                        && !window.location.pathname.includes('/login')
                        && (document.querySelector('[data-tid]') !== null
                            || document.querySelector('app-bar') !== null);
                }
                """)
                if is_teams:
                    teams_loaded = True
                    log("✅ Teams đã load!")
                    break
            except Exception:
                pass
            if i % 5 == 4:
                log(f"  ⏳ {(i+1)*2}s…")

        if not teams_loaded:
            log("⚠️ Teams chưa load hoàn toàn, thử tiếp…")

        # Navigate đến Chat tab
        log("📌 Thử navigate đến Chat tab…")
        chat_clicked = False

        # Approach 1: Click Chat button trên app bar
        for selector in [
            '#app-bar-chat-button',
            'button[data-tid="app-bar-2"]',
            '[data-tid="app-bar-chat-button"]',
            'button[title="Chat"]',
            '[aria-label="Chat"]',
            'button:has-text("Chat")',
            'a[href*="/chat"]',
            'nav button:nth-child(2)',  # Chat thường là button thứ 2
        ]:
            try:
                el = page.query_selector(selector)
                if el:
                    el.click()
                    log(f"✅ Clicked Chat button ({selector})")
                    chat_clicked = True
                    break
            except Exception:
                continue

        # Approach 2: Navigate trực tiếp qua URL
        if not chat_clicked:
            log("📌 Thử navigate trực tiếp đến /chat…")
            try:
                page.goto("https://teams.microsoft.com/v2/#/conversations", timeout=30_000)
                chat_clicked = True
                log("✅ Navigated đến /v2/#/conversations")
            except Exception:
                try:
                    page.goto("https://teams.microsoft.com/#/conversations", timeout=30_000)
                    chat_clicked = True
                    log("✅ Navigated đến /#/conversations")
                except Exception as e:
                    log(f"⚠️ Navigate đến Chat URL thất bại: {e}")

        # Chờ chat data được intercepted
        log("⏳ Chờ Teams load chat data (tối đa 20s)…")
        chat_list_captured.wait(timeout=20)

        # ── Strategy A: Dùng captured chatsvc token để gọi list API ──────
        if not chat_list_captured.is_set() and captured_tokens.get("chatsvc_bearer"):
            log("📌 Có chatsvc token! Thử gọi conversations list API trực tiếp…")
            base = captured_tokens.get("chatsvc_url_base", "")
            token = captured_tokens["chatsvc_bearer"]
            if base:
                # Gọi API trong browser context (để tận dụng cookies + CORS)
                try:
                    api_result = page.evaluate("""
                    async ([baseUrl, bearerToken]) => {
                        const endpoints = [
                            baseUrl + '/users/ME/conversations?view=msnp24Equivalent&pageSize=200',
                            baseUrl + '/users/ME/conversations?pageSize=200',
                        ];
                        for (const url of endpoints) {
                            try {
                                const resp = await fetch(url, {
                                    headers: {
                                        'Authorization': 'Bearer ' + bearerToken,
                                        'Content-Type': 'application/json',
                                    }
                                });
                                if (resp.ok) {
                                    const data = await resp.json();
                                    const convs = data.conversations || data.chats || data.value || [];
                                    if (convs.length > 0) {
                                        return { ok: true, conversations: convs, url: url, count: convs.length };
                                    }
                                }
                            } catch(e) {}
                        }
                        return { ok: false };
                    }
                    """, [base, token])

                    if api_result and api_result.get("ok"):
                        convs = api_result.get("conversations", [])
                        conversations_raw.extend(convs)
                        log(f"✅ API trả về {api_result.get('count', 0)} conversations!")
                        chat_list_captured.set()
                    else:
                        log("⚠️ API list không trả về data, thử fetch ngoài browser…")
                except Exception as e:
                    log(f"⚠️ In-browser fetch error: {e}")

                # Fallback: gọi từ Python requests
                if not chat_list_captured.is_set():
                    import requests as _req_pw
                    list_urls = [
                        f"{base}/users/ME/conversations?view=msnp24Equivalent&pageSize=200",
                        f"{base}/users/ME/conversations?pageSize=200",
                    ]
                    for list_url in list_urls:
                        try:
                            resp = _req_pw.get(list_url, headers={
                                "Authorization": f"Bearer {token}",
                                "Content-Type": "application/json",
                            }, timeout=15)
                            log(f"  → {list_url[:80]}… → HTTP {resp.status_code}")
                            if resp.status_code == 200:
                                data = resp.json()
                                convs = data.get("conversations", []) or data.get("chats", []) or data.get("value", [])
                                if convs:
                                    conversations_raw.extend(convs)
                                    log(f"✅ Python requests: {len(convs)} conversations!")
                                    chat_list_captured.set()
                                    break
                        except Exception as e:
                            log(f"  ⚠️ {e}")

        # ── Strategy B: Scroll + wait thêm ──────────────────────────────────
        if not conversations_raw:
            log("⚠️ Chưa có chat data, thử scroll chat list…")
            try:
                page.evaluate("""
                () => {
                    const containers = document.querySelectorAll('[role="list"], [data-tid="chat-list"], .chat-list');
                    for (const c of containers) {
                        c.scrollTop = c.scrollHeight;
                    }
                    const recent = document.querySelector('[data-tid="recent-chat-list"]');
                    if (recent) recent.click();
                }
                """)
            except Exception:
                pass
            chat_list_captured.wait(timeout=10)

        # ── Strategy C: DOM scraping (chỉ khi chưa có data từ API) ────────
        if not conversations_raw:
            log("⚠️ Thử lấy chat list từ DOM trực tiếp…")
            try:
                # Khi scrape DOM, PHẢI click vào từng chat để get thread ID
                dom_chats = page.evaluate("""
                () => {
                    const results = [];
                    const selectors = [
                        '[data-tid*="chat-list-item"]',
                        '[role="treeitem"]',
                        '[role="listitem"]',
                        '.chat-list-item',
                        '[data-tid*="listitem"]',
                    ];
                    const seen = new Set();
                    for (const sel of selectors) {
                        const items = document.querySelectorAll(sel);
                        for (const item of items) {
                            const allText = item.textContent.trim();
                            if (!allText || allText.length < 3) continue;
                            if (seen.has(allText.substring(0, 50))) continue;
                            seen.add(allText.substring(0, 50));

                            let name = '';
                            const nameEls = item.querySelectorAll(
                                '[data-tid*="title"], [class*="title"], [class*="displayName"], h2, h3, span[title]'
                            );
                            for (const nel of nameEls) {
                                const t = nel.textContent.trim();
                                if (t && t.length > 1 && t.length < 200) { name = t; break; }
                            }
                            if (!name) {
                                const spans = item.querySelectorAll('span');
                                for (const sp of spans) {
                                    const t = sp.textContent.trim();
                                    if (t && t.length > 1 && t.length < 100 && !t.match(/^\\d+[smhd]? ago$/)) {
                                        name = t; break;
                                    }
                                }
                            }

                            // Tìm thread ID trong mọi attribute (bao gồm data-*, aria-*, href)
                            let chatId = '';
                            // 1. Tìm trong tất cả attributes của item và children
                            const allEls = [item, ...item.querySelectorAll('*')];
                            for (const el of allEls) {
                                for (const attr of el.attributes) {
                                    if (attr.value && attr.value.includes('19:') && attr.value.includes('@thread')) {
                                        const m = attr.value.match(/(19:[a-f0-9-]+@thread\\.v2)/);
                                        if (m) { chatId = decodeURIComponent(m[1]); break; }
                                    }
                                }
                                if (chatId) break;
                            }
                            // 2. Tìm trong href
                            if (!chatId) {
                                const links = item.querySelectorAll('a[href]');
                                for (const link of links) {
                                    const href = link.getAttribute('href') || '';
                                    const m = href.match(/(19:[a-f0-9-]+@thread\\.v2)/);
                                    if (m) { chatId = decodeURIComponent(m[1]); break; }
                                }
                            }
                            // 3. Tìm trong onclick / data attributes
                            if (!chatId) {
                                const dataStr = JSON.stringify(
                                    Object.fromEntries([...item.attributes].map(a => [a.name, a.value]))
                                );
                                const m = dataStr.match(/(19:[a-f0-9-]+@thread\\.v2)/);
                                if (m) chatId = m[1];
                            }

                            // CHỈ thêm nếu có thread ID hợp lệ hoặc tên rõ ràng
                            if (name && chatId) {
                                results.push({ name, chatId, source: 'dom' });
                            }
                            // Skip items không có thread ID — chúng vô dụng cho export
                        }
                    }
                    return results;
                }
                """)
                if dom_chats:
                    log(f"📌 Tìm thấy {len(dom_chats)} chats từ DOM (có thread ID)!")
                    for dc in dom_chats:
                        conversations_raw.append({
                            "_dom_chat": True,
                            "display_name": dc.get("name", ""),
                            "dom_id": dc.get("chatId", ""),
                        })
                else:
                    log("⚠️ DOM scraping: Không tìm thấy chat nào có thread ID hợp lệ.")
            except Exception as e:
                log(f"⚠️ DOM scraping error: {e}")

        # ── Save chatsvc token for future use ────────────────────────────────
        if captured_tokens.get("chatsvc_bearer"):
            try:
                chatsvc_file = BASE_DIR / "chatsvc_token.json"
                chatsvc_file.write_text(json.dumps({
                    "token": captured_tokens["chatsvc_bearer"],
                    "base_url": captured_tokens.get("chatsvc_url_base", ""),
                    "fetched_at": datetime.now(timezone.utc).isoformat(),
                }, ensure_ascii=False, indent=2), encoding="utf-8")
                log(f"💾 Saved chatsvc token to chatsvc_token.json")
            except Exception:
                pass

        # Close browser
        context.close()
        log(f"📊 Tổng cộng: {len(conversations_raw)} raw conversations")

    # ── Parse conversations → standardized format ────────────────────────
    for conv in conversations_raw:
        # Graph chat format
        if conv.get("_graph_chat"):
            chats_data.append({
                "chat_id":      conv.get("id", ""),
                "chat_type":    conv.get("chatType", "group"),
                "display_name": conv.get("topic", "") or f"Chat ({conv.get('chatType', '')})",
                "member_count": len(conv.get("members", [])),
                "members":      [],
                "_source":      "pw_graph",
            })
            continue

        # DOM-scraped chat
        if conv.get("_dom_chat"):
            chats_data.append({
                "chat_id":      conv.get("dom_id", f"dom_{len(chats_data)}"),
                "chat_type":    "group",
                "display_name": conv.get("display_name", "Unknown"),
                "member_count": 0,
                "members":      [],
                "_source":      "pw_dom",
                "_preview":     conv.get("preview", ""),
            })
            continue

        # Teams internal conversation format
        thread_props = conv.get("threadProperties", {})
        conv_id = conv.get("id", "")

        is_group = "@thread" in conv_id
        is_one_on_one = conv_id.startswith("8:")
        if not is_group and not is_one_on_one:
            continue

        topic = thread_props.get("topic", "")
        members_raw = thread_props.get("members", "")
        member_list = []
        if isinstance(members_raw, str) and members_raw:
            member_list = [m.strip() for m in members_raw.split(",") if m.strip()]
        elif isinstance(members_raw, list):
            member_list = members_raw

        chat_type = "group" if is_group else "oneOnOne"
        if topic:
            display_name = topic
        elif len(member_list) <= 5:
            names = [m.split(":")[-1][:15] for m in member_list[:5]]
            display_name = ", ".join(names) if names else f"Chat ({conv_id[:20]}…)"
        else:
            display_name = f"Group ({len(member_list)} members)"

        chats_data.append({
            "chat_id":      conv_id,
            "chat_type":    chat_type,
            "display_name": display_name,
            "member_count": len(member_list),
            "members":      member_list[:10],
            "_source":      "pw_intercept",
        })

    # Deduplicate by chat_id
    seen = set()
    unique_chats = []
    for c in chats_data:
        cid = c.get("chat_id", "")
        if cid and cid in seen:
            continue
        seen.add(cid)
        unique_chats.append(c)

    # Save to file for use by step2
    pw_chats_file = BASE_DIR / "pw_chats.json"
    pw_chats_file.write_text(
        json.dumps(unique_chats, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    log(f"💾 Saved {len(unique_chats)} chats to pw_chats.json")

    return unique_chats


@app.route("/step1/chrome_profile", methods=["POST"])
def step1_chrome_profile():
    """Khởi động Playwright Browser Login trong background thread."""
    job_id = str(uuid.uuid4())
    _cdp_jobs[job_id] = {
        "status": "starting",
        "log": [],
        "token": None,
        "error": None,
    }

    def _run():
        job = _cdp_jobs[job_id]
        try:
            job["log"].append("🌐 Đang mở Playwright Chromium — vui lòng đăng nhập Teams trên cửa sổ vừa mở…")
            token = _get_token_playwright_login(job=job)
            if token:
                save_token(token)
                job["token"] = token
                job["status"] = "done"
                job["log"].append("✅ Lấy token thành công!")
            else:
                job["status"] = "error"
                job["error"] = "Không bắt được token. Teams có thể chưa load xong hoặc chưa đăng nhập."
                job["log"].append("❌ Không tìm thấy token.")
        except Exception as e:
            job["status"] = "error"
            job["error"] = f"{type(e).__name__}: {e}"
            job["log"].append(f"❌ {e}")

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/step1/chrome_profile/poll/<job_id>")
def step1_chrome_profile_poll(job_id: str):
    job = _cdp_jobs.get(job_id)
    if not job:
        return jsonify({"status": "error", "error": "Job not found"}), 404
    return jsonify(job)


# ── Cách 2: Lấy token từ Chrome (CDP — tự khởi động lại Chrome) ──────────────

def _find_free_port(start: int = 9222, end: int = 9240) -> int:
    """Tìm port không có ai đang listen."""
    import socket
    for p in range(start, end):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.5)
            result = s.connect_ex(("127.0.0.1", p))
            if result != 0:
                # Không ai listen → port trống
                return p
            # result == 0 → có process đang listen → bỏ qua
    raise RuntimeError(f"Không tìm được port trống trong khoảng {start}-{end}")


def _get_token_cdp_auto(profile_dir: str = None, port: int = 0, job: dict = None) -> str | None:
    """
    Đóng Chrome → khởi động lại Chrome với --remote-debugging-port + profile gốc
    → mở Teams → bắt token qua CDP.
    User nên mở web app trên Edge (không bị ảnh hưởng).
    """
    import subprocess
    import time
    import urllib.request
    from playwright.sync_api import sync_playwright

    def log(msg):
        if job:
            job["log"].append(msg)
            job["status"] = msg[:20]
        print(f"[CDP] {msg}")

    if profile_dir is None:
        profile_dir = os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data")

    # Tự tìm port trống nếu chưa chỉ định
    if port == 0:
        port = _find_free_port()
        log(f"🔌 Free port: {port}")

    # Tìm Chrome executable
    chrome_exe = None
    for candidate in [
        os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ]:
        if os.path.exists(candidate):
            chrome_exe = candidate
            break

    if not chrome_exe:
        raise FileNotFoundError("Không tìm thấy Chrome. Hãy cài Google Chrome.")
    log(f"✅ Chrome: {chrome_exe}")

    # Kiểm tra xem Chrome đã chạy với CDP chưa
    cdp_already_running = False
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/version", timeout=2) as r:
            info = json.loads(r.read())
            if "webSocketDebuggerUrl" in info:
                cdp_already_running = True
                log(f"✅ CDP đã sẵn sàng tại port {port}")
    except Exception:
        pass

    if not cdp_already_running:
        # Đóng TẤT CẢ Chrome (Edge không bị ảnh hưởng)
        log("🔄 Đang đóng Chrome…")
        subprocess.run(
            ["taskkill", "/F", "/IM", "chrome.exe"],
            capture_output=True, timeout=10,
        )

        # Chờ Chrome đóng hoàn toàn — verify bằng tasklist
        for i in range(10):
            time.sleep(1)
            r = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq chrome.exe"],
                capture_output=True, text=True, timeout=5,
            )
            if "chrome.exe" not in r.stdout.lower():
                log(f"✅ Chrome đã đóng (sau {i+1}s)")
                break

        time.sleep(3)  # buffer thêm cho OS giải phóng port/file lock

        # Tìm port trống SAU khi Chrome đã đóng hoàn toàn
        port = _find_free_port()
        log(f"🔌 Port sau kill: {port}")

        # Khởi động Chrome mới với CDP + profile gốc + mở Teams
        chrome_cmd = [
            chrome_exe,
            f"--remote-debugging-port={port}",
            f"--user-data-dir={profile_dir}",
            "--no-first-run",
            "https://teams.microsoft.com",
        ]
        log(f"🚀 Launching Chrome trên port {port}…")
        subprocess.Popen(chrome_cmd)

        # Chờ Chrome khởi động và CDP endpoint sẵn sàng
        cdp_ready = False
        for i in range(30):
            time.sleep(1)
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/version", timeout=2) as r:
                    data = json.loads(r.read())
                    if "webSocketDebuggerUrl" in data:
                        cdp_ready = True
                        log(f"✅ CDP sẵn sàng (sau {i+1}s)")
                        break
            except Exception:
                if i % 5 == 4:
                    log(f"⏳ Chờ CDP… ({i+1}s)")
                continue

        if not cdp_ready:
            r = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq chrome.exe"],
                capture_output=True, text=True, timeout=5,
            )
            chrome_running = "chrome.exe" in r.stdout.lower()
            raise ConnectionError(
                f"Chrome CDP không sẵn sàng tại cổng {port}. "
                f"Chrome process running: {chrome_running}. "
                f"Thử mở thủ công: chrome.exe --remote-debugging-port={port}"
            )

        # Chờ thêm để Teams tab load
        time.sleep(5)

    captured = {"token": None}

    try:
        log("🔗 Kết nối Playwright → Chrome CDP…")
        with sync_playwright() as p:
            try:
                browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{port}")
                log("✅ Đã kết nối CDP")
            except Exception as e:
                raise ConnectionError(
                    f"Không thể kết nối Chrome CDP tại 127.0.0.1:{port}. "
                    f"Chi tiết: {e}"
                )

            # Chờ và tìm tab Teams (có thể đang load)
            log("🔍 Tìm tab Teams…")
            teams_page = None
            for attempt in range(10):
                for context in browser.contexts:
                    for page in context.pages:
                        url = page.url
                        if "teams.microsoft.com" in url:
                            teams_page = page
                            break
                    if teams_page:
                        break
                if teams_page:
                    break
                time.sleep(2)

            # Nếu chưa có tab Teams → mở mới trong context đầu tiên
            if not teams_page:
                log("📄 Không tìm thấy tab Teams → mở mới…")
                ctx = browser.contexts[0] if browser.contexts else browser.new_context()
                teams_page = ctx.new_page()
                teams_page.goto("https://teams.microsoft.com", timeout=60_000)
            else:
                log(f"✅ Tìm thấy tab Teams: {teams_page.url[:60]}…")

            def on_request(req):
                if captured["token"]:
                    return
                auth = req.headers.get("authorization", "")
                if auth.startswith("Bearer "):
                    url = req.url
                    if any(domain in url for domain in (
                        "graph.microsoft.com",
                        "api.spaces.skype.com",
                        "teams.microsoft.com/api",
                        "substrate.office.com",
                    )):
                        token_val = auth.removeprefix("Bearer ")
                        if len(token_val) > 200:
                            captured["token"] = token_val

            def on_response(resp):
                if captured["token"]:
                    return
                if "login.microsoftonline.com" in resp.url and "/oauth2" in resp.url:
                    try:
                        body = resp.json()
                        if "access_token" in body:
                            captured["token"] = body["access_token"]
                    except Exception:
                        pass

            teams_page.on("request", on_request)
            teams_page.on("response", on_response)

            # Chờ Teams load
            log("⏳ Chờ Teams load…")
            try:
                teams_page.wait_for_selector(
                    '[data-tid="team-channel-list"], [data-tid="chat-list"], '
                    '#app-bar-chat-button, button[title="Chat"], '
                    'div[class*="leftRail"]',
                    timeout=30_000,
                )
            except Exception:
                pass

            teams_page.wait_for_timeout(3_000)

            if captured["token"]:
                log("✅ Token bắt được từ network request!")

            # Inject JS để extract token từ Teams internals
            if not captured["token"]:
                log("🔍 Thử extract token từ localStorage/sessionStorage…")
                try:
                    js_token = teams_page.evaluate("""
                    () => {
                        const stores = [sessionStorage, localStorage];
                        for (const store of stores) {
                            for (let i = 0; i < store.length; i++) {
                                const key = store.key(i);
                                try {
                                    const val = store.getItem(key);
                                    if (val && val.includes('accessToken')) {
                                        const obj = JSON.parse(val);
                                        if (obj.secret && obj.secret.length > 200) return obj.secret;
                                        if (obj.accessToken && obj.accessToken.length > 200) return obj.accessToken;
                                    }
                                    if (val && val.startsWith('eyJ') && val.length > 200) return val;
                                } catch (e) {}
                            }
                        }
                        return null;
                    }
                    """)
                    if js_token and len(js_token) > 200:
                        captured["token"] = js_token
                except Exception:
                    pass

            # Click Chat tab để trigger thêm requests
            if not captured["token"]:
                try:
                    chat_btn = teams_page.locator(
                        '#app-bar-chat-button, [data-tid="app-bar-chat-button"], '
                        'button[title="Chat"], button[aria-label="Chat"]'
                    ).first
                    chat_btn.click(timeout=5_000)
                    teams_page.wait_for_timeout(5_000)
                except Exception:
                    pass

            # Click channel để trigger Graph request
            if not captured["token"]:
                try:
                    teams_page.locator('[data-tid="channel-list-item"]').first.click()
                    teams_page.wait_for_timeout(5_000)
                except Exception:
                    pass

            # Đóng Edge sau khi xong (Chrome không bị ảnh hưởng)
            try:
                browser.close()
            except Exception:
                pass

    except Exception:
        raise

    return captured["token"]


def _get_token_from_running_chrome(port: int = 9222) -> str | None:
    """
    Kết nối vào Chrome đang chạy qua CDP và bắt token từ Graph API request của Teams.
    Chrome phải được khởi động với flag: --remote-debugging-port=<port>
    và đang mở tab teams.microsoft.com.
    """
    from playwright.sync_api import sync_playwright  # noqa: PLC0415

    captured = {"token": None}

    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp(f"http://localhost:{port}")
        except Exception as e:
            raise ConnectionError(
                f"Không thể kết nối Chrome tại cổng {port}. "
                f"Hãy khởi động Chrome với flag --remote-debugging-port={port}. "
                f"Chi tiết: {e}"
            )

        # Tìm tab đang mở Teams
        teams_page = None
        for context in browser.contexts:
            for page in context.pages:
                if "teams.microsoft.com" in page.url:
                    teams_page = page
                    break
            if teams_page:
                break

        if not teams_page:
            browser.close()
            raise ValueError(
                "Không tìm thấy tab Teams trong Chrome đang chạy. "
                "Hãy mở https://teams.microsoft.com trong Chrome rồi thử lại."
            )

        def on_request(req):
            auth = req.headers.get("authorization", "")
            if (
                auth.startswith("Bearer ")
                and "graph.microsoft.com" in req.url
                and not captured["token"]
            ):
                token_val = auth.removeprefix("Bearer ")
                if len(token_val) > 200:
                    captured["token"] = token_val

        teams_page.on("request", on_request)

        # Click vào channel đầu tiên để trigger Graph API request
        try:
            teams_page.locator('[data-tid="channel-list-item"]').first.click()
            teams_page.wait_for_timeout(5_000)
        except Exception:
            teams_page.wait_for_timeout(5_000)

        browser.close()

    return captured["token"]


@app.route("/step1/from_chrome", methods=["POST"])
def step1_from_chrome():
    """Lấy token từ Chrome đang chạy qua Chrome DevTools Protocol (CDP)."""
    port_str = request.form.get("cdp_port", "9222").strip()
    try:
        port = int(port_str)
    except ValueError:
        flash("❌ CDP port phải là số nguyên (mặc định: 9222).", "error")
        return redirect(url_for("step1"))

    try:
        token = _get_token_from_running_chrome(port=port)
        if token:
            save_token(token)
            flash("✅ Token lấy từ Chrome thành công!", "success")
            return redirect(url_for("step2"))
        flash(
            "❌ Không bắt được token. "
            "Hãy thử click vào một channel trong Teams rồi bấm lại.",
            "error",
        )
    except ConnectionError as e:
        flash(f"❌ {e}", "error")
    except ValueError as e:
        flash(f"❌ {e}", "error")
    except Exception as e:
        flash(f"❌ Lỗi không xác định: {e}", "error")

    return redirect(url_for("step1"))


# ── Device Code Flow (MSAL) ───────────────────────────────────────────────────
# Lưu trạng thái device-code job: { job_id: {"status", "user_code", "url", "token", "error"} }
_dc_jobs: dict[str, dict] = {}


@app.route("/step1/device_code", methods=["POST"])
def step1_device_code_start():
    """Khởi động device code flow — tự động thử các client ID cho đến khi thành công."""
    import msal

    email = request.form.get("email", "").strip()

    # Danh sách Microsoft first-party app IDs (luôn pre-approved trong M365 tenants)
    # Azure CLI đặt đầu vì có preauthorization tốt nhất cho device code flow
    CANDIDATE_CLIENTS = [
        ("Azure CLI",                "04b07795-8ddb-461a-bbee-02f9e1bf7b46"),
        ("Microsoft Office",         "d3590ed6-52b3-4102-aeff-aad2292ab01c"),
        ("Microsoft Teams",          "1fec8e78-bce4-4aaf-ab1b-5451cc387264"),
    ]
    # Thử full scopes trước, nếu thất bại thì giảm dần
    SCOPE_SETS = [
        ("Full (channel + chat)", [
            "https://graph.microsoft.com/Team.ReadBasic.All",
            "https://graph.microsoft.com/Channel.ReadBasic.All",
            "https://graph.microsoft.com/ChannelMessage.Read.All",
            "https://graph.microsoft.com/Chat.Read",
            "https://graph.microsoft.com/Chat.ReadBasic",
            "https://graph.microsoft.com/ChatMessage.Read",
        ]),
        ("Chat only", [
            "https://graph.microsoft.com/Chat.Read",
            "https://graph.microsoft.com/Chat.ReadBasic",
            "https://graph.microsoft.com/ChatMessage.Read",
        ]),
        ("Channel only", [
            "https://graph.microsoft.com/Team.ReadBasic.All",
            "https://graph.microsoft.com/Channel.ReadBasic.All",
            "https://graph.microsoft.com/ChannelMessage.Read.All",
        ]),
        ("User.Read only", [
            "https://graph.microsoft.com/User.Read",
        ]),
    ]

    job_id = str(uuid.uuid4())
    _dc_jobs[job_id] = {"status": "waiting", "user_code": None, "url": None,
                        "token": None, "error": None, "log": []}

    def _run():
        job = _dc_jobs[job_id]
        try:
            # Detect tenant ID từ email để tránh lỗi "organizations"
            authority = "https://login.microsoftonline.com/organizations"
            if email and "@" in email:
                domain = email.split("@")[-1]
                try:
                    import urllib.request
                    url = f"https://login.microsoftonline.com/{domain}/.well-known/openid-configuration"
                    with urllib.request.urlopen(url, timeout=5) as r:
                        meta = json.loads(r.read())
                    tenant_ep = meta.get("token_endpoint", "")
                    parts = tenant_ep.split("/")
                    idx = next((i for i, p in enumerate(parts) if "microsoftonline.com" in p), -1)
                    if idx >= 0 and idx + 1 < len(parts):
                        tid = parts[idx + 1]
                        if len(tid) > 10:
                            authority = f"https://login.microsoftonline.com/{tid}"
                            job["log"].append(f"✅ Detected tenant: {tid}")
                except Exception as e:
                    job["log"].append(f"⚠️ Tenant detect failed: {e}")

            last_err = "Không tìm được client ID / scope phù hợp với tenant."
            for scope_label, scopes in SCOPE_SETS:
                job["log"].append(f"📋 Scope set: {scope_label}")
                for app_name, client_id in CANDIDATE_CLIENTS:
                    job["log"].append(f"  🔑 Trying: {app_name} ({client_id[:8]}…)")
                    app_msal = msal.PublicClientApplication(client_id, authority=authority)
                    flow = app_msal.initiate_device_flow(scopes=scopes)

                    if "user_code" not in flow:
                        err = flow.get("error_description", flow.get("error", ""))
                        short_err = err[:200] if err else "(no error message)"
                        job["log"].append(f"    ❌ {short_err}")
                        # Lỗi tenant/consent/app → thử client tiếp theo
                        if any(code in err for code in (
                            "AADSTS1001010", "AADSTS700016", "AADSTS70001",
                            "AADSTS65001", "AADSTS65002", "AADSTS70002", "AADSTS700027",
                        )):
                            last_err = f"{app_name} [{scope_label}]: {err}"
                            continue
                        # Lỗi khác (network, config…) → vẫn thử tiếp client khác
                        last_err = f"{app_name}: {err}"
                        continue

                    # Đã có mã — gửi về cho frontend hiển thị
                    job["user_code"] = flow["user_code"]
                    job["url"]       = flow["verification_uri"]
                    job["status"]    = "pending"
                    job["log"].append(f"    ✅ Got code: {flow['user_code']} via {app_name} [{scope_label}]")

                    result = app_msal.acquire_token_by_device_flow(flow)  # blocking

                    if "access_token" in result:
                        save_token(result["access_token"])
                        job["token"]  = result["access_token"]
                        job["status"] = "done"
                        return

                    # Lỗi sau khi nhập code
                    acq_err = result.get("error_description", result.get("error", ""))
                    job["log"].append(f"    ❌ Auth failed: {acq_err[:200]}")

                    # Lỗi consent/preauth → reset UI và thử combo tiếp
                    if any(code in acq_err for code in (
                        "AADSTS65002", "AADSTS65001", "AADSTS70011",
                        "AADSTS700027", "AADSTS50076",
                    )):
                        job["user_code"] = None
                        job["status"]    = "waiting"
                        last_err = f"{app_name} [{scope_label}]: {acq_err}"
                        continue

                    # Lỗi khác (user declined, expired…) → dừng
                    job["status"] = "error"
                    job["error"]  = acq_err
                    return

            # Hết danh sách mà không app nào được chấp nhận
            job["status"] = "error"
            job["error"]  = (
                f"{last_err}\n\n"
                "💡 Gợi ý: Nhập email công ty vào ô bên dưới nút để hệ thống "
                "tự nhận diện đúng tenant, hoặc liên hệ IT admin để được cấp quyền."
            )

        except Exception as e:
            job["status"] = "error"
            job["error"]  = str(e)

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/step1/device_code/poll/<job_id>")
def step1_device_code_poll(job_id: str):
    """Client poll để biết user_code và trạng thái xác thực."""
    job = _dc_jobs.get(job_id)
    if not job:
        return jsonify({"status": "error", "error": "Job not found"}), 404
    return jsonify({
        "status":    job["status"],
        "user_code": job["user_code"],
        "url":       job["url"],
        "error":     job["error"],
        "log":       job.get("log", []),
    })


# ── Step 1: Get Token ─────────────────────────────────────────────────────────

@app.route("/")
@app.route("/step1")
def step1():
    return render_template(
        "step1_token.html",
        token_status=get_token_status(),
        current_step=1,
    )


@app.route("/step1/sso", methods=["POST"])
def step1_sso():
    cfg = load_config()
    browser = cfg.get("browser", "edge")
    try:
        token = get_token_sso(browser=browser)
        if token:
            save_token(token)
            flash("✅ Token obtained successfully via SSO!", "success")
            return redirect(url_for("step2"))
        flash("❌ Failed to obtain token via SSO. Please try again.", "error")
    except Exception as e:
        flash(f"❌ SSO error: {e}", "error")
    return redirect(url_for("step1"))


@app.route("/step1/password", methods=["POST"])
def step1_password():
    email    = request.form.get("email", "").strip()
    password = request.form.get("password", "")
    if not email or not password:
        flash("❌ Please enter both email and password.", "error")
        return redirect(url_for("step1"))

    cfg = load_config()
    browser = cfg.get("browser", "edge")
    try:
        token = _get_token_password_web(email, password, browser=browser)
        if token:
            save_token(token)
            flash("✅ Token obtained via email/password!", "success")
            return redirect(url_for("step2"))
        flash("❌ Failed to obtain token. Check your credentials.", "error")
    except Exception as e:
        flash(f"❌ Error: {e}", "error")
    return redirect(url_for("step1"))


@app.route("/step1/manual", methods=["POST"])
def step1_manual():
    import base64 as _b64
    token = request.form.get("token", "").strip()
    if not token:
        flash("❌ Please paste a token.", "error")
        return redirect(url_for("step1"))
    if len(token) < 100:
        flash("❌ Token looks too short — make sure you copied the full Bearer token.", "error")
        return redirect(url_for("step1"))
    parts = token.split(".")
    if len(parts) < 3:
        flash("❌ Token format invalid — expected a JWT (3 dot-separated segments).", "error")
        return redirect(url_for("step1"))

    # Validate that this token is scoped for Microsoft Graph, not the Teams web app.
    # Graph tokens have aud == 'https://graph.microsoft.com' or
    # '00000003-0000-0000-c000-000000000000'.
    GRAPH_AUDIENCES = {
        "https://graph.microsoft.com",
        "00000003-0000-0000-c000-000000000000",
    }
    try:
        padded  = parts[1] + "=" * (4 - len(parts[1]) % 4)
        payload = json.loads(_b64.b64decode(padded))
        aud = payload.get("aud", "")
        if aud not in GRAPH_AUDIENCES:
            flash(
                f"❌ Wrong token audience: '{aud}'. "
                "This token is for the Teams web app, not Microsoft Graph. "
                "In DevTools, filter Network requests by 'graph.microsoft.com' and copy "
                "the Authorization header from one of those requests.",
                "error",
            )
            return redirect(url_for("step1"))
    except Exception:
        pass  # If we can’t decode it, let save_token handle the bad token later

    save_token(token)
    flash("✅ Token saved! (Graph API audience confirmed)", "success")
    return redirect(url_for("step2"))


# ── Step 2: Select Channels ───────────────────────────────────────────────────

@app.route("/step2")
def step2():
    cfg          = load_config()
    export_state = load_export_state()
    all_channels = load_all_channels()
    teams        = group_channels_by_team(all_channels)
    selected_ids = {
        f"{c['team_id']}_{c['channel_id']}"
        for c in cfg.get("channels", [])
    }
    selected_chat_ids = {
        c["chat_id"] for c in cfg.get("chats", [])
    }
    return render_template(
        "step2_channels.html",
        teams=teams,
        config=cfg,
        export_state=export_state,
        selected_ids=selected_ids,
        selected_chat_ids=selected_chat_ids,
        saved_chats=cfg.get("chats", []),
        token_status=get_token_status(),
        current_step=2,
    )


# ── Playwright Chat Scraping Route ──────────────────────────────────────────

@app.route("/step2/pw_scrape_chats", methods=["POST"])
def step2_pw_scrape_chats():
    """Khởi động Playwright để scrape chat list từ Teams web."""
    job_id = str(uuid.uuid4())
    _pw_chat_jobs[job_id] = {
        "status": "running",
        "log": [],
        "chats": [],
        "error": None,
    }

    def _run():
        job = _pw_chat_jobs[job_id]
        try:
            chats = _scrape_chats_via_playwright(job=job)
            job["chats"] = chats
            job["status"] = "done"
            job["log"].append(f"✅ Hoàn thành! Tìm thấy {len(chats)} chats.")
        except Exception as e:
            job["status"] = "error"
            job["error"] = f"{type(e).__name__}: {e}"
            job["log"].append(f"❌ Lỗi: {e}")

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/step2/pw_scrape_chats/poll/<job_id>")
def step2_pw_scrape_poll(job_id: str):
    job = _pw_chat_jobs.get(job_id)
    if not job:
        return jsonify({"status": "error", "error": "Job not found"}), 404
    return jsonify(job)


@app.route("/step2/refresh", methods=["POST"])
def step2_refresh():
    try:
        token   = get_valid_token()
        headers = lc.make_headers(token)

        # ── Decode token audience for diagnostics ────────────────────────────
        token_aud = ""
        token_scp = ""
        try:
            import base64 as _b64
            parts = token.split(".")
            if len(parts) >= 2:
                padded = parts[1] + "=" * (-len(parts[1]) % 4)
                payload = json.loads(_b64.urlsafe_b64decode(padded))
                token_aud = payload.get("aud", "")
                token_scp = payload.get("scp", "")
        except Exception:
            pass

        is_graph_token = (
            "graph.microsoft.com" in token_aud
            or "00000003-0000-0000-c000-000000000000" in token_aud
        )

        # ── Lấy Teams + Channels (cần Graph token) ──────────────────────────
        result       = []
        flat_channels = []
        team_errors  = []

        try:
            teams_list = lc.fetch_joined_teams(headers)
            for team in teams_list:
                channels = lc.fetch_channels(team["id"], headers)
                team_data = {
                    "team_id":   team["id"],
                    "team_name": team["displayName"],
                    "channels": [
                        {"channel_id": ch["id"], "channel_name": ch["displayName"]}
                        for ch in channels
                    ],
                }
                result.append(team_data)
                for ch in channels:
                    flat_channels.append({
                        "team_name":    team["displayName"],
                        "team_id":      team["id"],
                        "channel_name": ch["displayName"],
                        "channel_id":   ch["id"],
                    })
        except Exception as e:
            err_msg = str(e)
            team_errors.append(err_msg)
            if not is_graph_token:
                team_errors.append(
                    f"⚠️ Token audience = '{token_aud}' — không phải Graph API token. "
                    f"Cần token với audience 'https://graph.microsoft.com'. "
                    f"Hãy thử lấy lại token bằng cách khác (Manual Token từ DevTools hoặc Graph Explorer)."
                )

        ALL_CHANNELS_FILE.write_text(
            json.dumps(flat_channels, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        # ── Lấy danh sách Group Chats ──────────────────────────────────────
        import requests as _req

        chats_result = []
        chat_errors  = []

        # Thử v1.0 trước, fallback sang /beta nếu lỗi
        for api_ver in ("v1.0", "beta"):
            chats_url = f"https://graph.microsoft.com/{api_ver}/me/chats?$top=50"
            page = 0
            while chats_url:
                page += 1
                try:
                    resp = _req.get(chats_url, headers=headers, timeout=30)
                    if resp.status_code in (401, 403):
                        chat_errors.append(
                            f"{api_ver}: HTTP {resp.status_code} — {resp.text[:200]}"
                        )
                        chats_url = None
                        continue
                    resp.raise_for_status()
                    data = resp.json()

                    for chat in data.get("value", []):
                        chat_type = chat.get("chatType", "")
                        if chat_type not in ("group", "oneOnOne", "meeting"):
                            continue

                        # Tạo tên hiển thị
                        topic = chat.get("topic") or ""
                        chat_id = chat["id"]

                        if topic:
                            display_name = topic
                        else:
                            display_name = f"Chat ({chat_id[:12]}…)"

                        chats_result.append({
                            "chat_id":      chat_id,
                            "chat_type":    chat_type,
                            "display_name": display_name,
                            "member_count": 0,
                            "members":      [],
                        })

                    chats_url = data.get("@odata.nextLink")
                except Exception as e:
                    chat_errors.append(f"{api_ver} page {page}: {e}")
                    chats_url = None

            if chats_result:
                break  # Đã lấy được → không cần thử endpoint khác

        # ── Fallback: Dùng Teams internal API nếu Graph bị 403 ────────────
        if not chats_result and chat_errors:
            teams_token_file = BASE_DIR / "teams_token.json"
            teams_token = None
            skype_token = None
            if teams_token_file.exists():
                try:
                    td = json.loads(teams_token_file.read_text(encoding="utf-8"))
                    teams_token = td.get("token")
                    skype_token = td.get("skype_token")
                except Exception:
                    pass

            # Thử từng token: teams_token, skype_token
            tokens_to_try = []
            if teams_token:
                tokens_to_try.append(("teams", teams_token))
            if skype_token:
                tokens_to_try.append(("skype", skype_token))

            if tokens_to_try:
                chat_errors.append("⏳ Thử Teams internal API (Skype)…")
                try:
                    conversations = []

                    # ─── Strategy 1: msg.teams.microsoft.com (classic) ──────
                    regions = ["apac.ng", "emea.ng", "amer.ng", ""]
                    for token_name, token_val in tokens_to_try:
                        if conversations:
                            break
                        # Thử cả skypetoken= và Bearer auth schemes
                        auth_schemes = [
                            ("skypetoken", {"Authorization": f"skypetoken={token_val}", "Content-Type": "application/json"}),
                            ("Bearer", {"Authorization": f"Bearer {token_val}", "Content-Type": "application/json"}),
                        ]
                        for auth_name, skype_headers in auth_schemes:
                            if conversations:
                                break
                            for region_prefix in regions:
                                if region_prefix:
                                    base_url = f"https://{region_prefix}.msg.teams.microsoft.com"
                                else:
                                    base_url = "https://msg.teams.microsoft.com"

                                conv_url = f"{base_url}/v1/users/ME/conversations?view=msnp24Equivalent&pageSize=200"
                                try:
                                    resp = _req.get(conv_url, headers=skype_headers, timeout=15)
                                    if resp.status_code == 200:
                                        data = resp.json()
                                        conversations = data.get("conversations", [])
                                        if conversations:
                                            chat_errors.append(f"✅ Teams API ({token_name}/{auth_name}/{region_prefix or 'global'}): {len(conversations)} conversations")
                                            break
                                    elif resp.status_code in (401, 403):
                                        chat_errors.append(f"❌ {token_name}/{auth_name}/{region_prefix or 'global'}: {resp.status_code}")
                                        continue
                                    else:
                                        chat_errors.append(f"⚠️ {token_name}/{auth_name}/{region_prefix or 'global'}: HTTP {resp.status_code}")
                                        continue
                                except Exception:
                                    continue

                    # ─── Strategy 2: chatsvcagg.teams.microsoft.com (new Teams) ──
                    if not conversations:
                        csa_regions = ["apac", "emea", "amer", ""]
                        for token_name, token_val in tokens_to_try:
                            if conversations:
                                break
                            auth_schemes2 = [
                                ("Bearer", {"Authorization": f"Bearer {token_val}", "Content-Type": "application/json"}),
                                ("skypetoken", {"Authorization": f"skypetoken={token_val}", "Content-Type": "application/json"}),
                            ]
                            for auth_name, csa_headers in auth_schemes2:
                                if conversations:
                                    break
                                for region in csa_regions:
                                    if region:
                                        csa_base = f"https://{region}.chatsvcagg.teams.microsoft.com"
                                    else:
                                        csa_base = "https://chatsvcagg.teams.microsoft.com"

                                    # chatsvcagg uses /v1/users/ME/conversations too
                                    csa_url = f"{csa_base}/v1/users/ME/conversations?view=msnp24Equivalent&pageSize=200"
                                    try:
                                        resp = _req.get(csa_url, headers=csa_headers, timeout=15)
                                        if resp.status_code == 200:
                                            data = resp.json()
                                            conversations = data.get("conversations", [])
                                            if conversations:
                                                chat_errors.append(f"✅ ChatSvcAgg ({token_name}/{auth_name}/{region or 'global'}): {len(conversations)} conversations")
                                                break
                                        elif resp.status_code in (401, 403):
                                            continue
                                    except Exception:
                                        continue

                    # Parse conversations → chats_result format
                    for conv in conversations:
                        thread_props = conv.get("threadProperties", {})
                        conv_id = conv.get("id", "")

                        # Filter: chỉ lấy group chats, 1:1, meetings
                        # Teams conversations có format: 19:xxx@thread.v2 (group) hoặc 19:xxx@unq.gbl.spaces
                        # hoặc 8:orgid:xxx (1:1)
                        is_group = "@thread" in conv_id
                        is_one_on_one = conv_id.startswith("8:")
                        if not is_group and not is_one_on_one:
                            continue

                        topic = thread_props.get("topic", "")
                        members_raw = thread_props.get("members", "")
                        member_list = []
                        if isinstance(members_raw, str) and members_raw:
                            # members is a comma-separated string of MRIs
                            member_list = [m.strip() for m in members_raw.split(",") if m.strip()]

                        # Determine chat type
                        if is_group:
                            chat_type = "group"
                        else:
                            chat_type = "oneOnOne"

                        # Display name
                        if topic:
                            display_name = topic
                        elif len(member_list) <= 5:
                            # Try to make a name from member MRIs
                            names = []
                            for m in member_list[:5]:
                                # MRI format: 8:orgid:user-oid
                                name = m.split(":")[-1][:12]
                                names.append(name)
                            display_name = ", ".join(names) if names else f"Chat ({conv_id[:15]}…)"
                        else:
                            display_name = f"Group ({len(member_list)} members)"

                        chats_result.append({
                            "chat_id":      conv_id,
                            "chat_type":    chat_type,
                            "display_name": display_name,
                            "member_count": len(member_list),
                            "members":      member_list[:10],
                            "_source":      "teams_internal",
                        })

                    if chats_result:
                        # Thử enrich display names bằng Graph /users nếu có
                        try:
                            for chat_item in chats_result[:50]:
                                if chat_item.get("_source") == "teams_internal":
                                    raw_members = chat_item.get("members", [])
                                    display_members = []
                                    for mri in raw_members[:6]:
                                        # Extract OID from 8:orgid:oid
                                        parts = mri.split(":")
                                        if len(parts) >= 3:
                                            oid = parts[-1]
                                            try:
                                                user_resp = _req.get(
                                                    f"https://graph.microsoft.com/v1.0/users/{oid}?$select=displayName",
                                                    headers=headers, timeout=5,
                                                )
                                                if user_resp.status_code == 200:
                                                    dn = user_resp.json().get("displayName", "")
                                                    if dn:
                                                        display_members.append(dn)
                                                        continue
                                            except Exception:
                                                pass
                                        display_members.append(mri.split(":")[-1][:15])

                                    if display_members and chat_item["display_name"].startswith(("Chat (", "Group (")):
                                        chat_item["display_name"] = ", ".join(display_members[:5])
                                        if len(display_members) > 5:
                                            chat_item["display_name"] += f" +{len(display_members)-5}"
                                    chat_item["members"] = display_members
                        except Exception:
                            pass

                        chat_errors.append(f"✅ Tìm thấy {len(chats_result)} chats qua Teams internal API!")
                except Exception as e:
                    chat_errors.append(f"Teams internal API error: {e}")
            else:
                chat_errors.append(
                    "💡 Token thiếu Chat.Read scope và chưa có Teams internal token. "
                    "Chạy lại Step 1 (Playwright Browser Login) để lấy Teams token."
                )

        # ── Fallback cuối: Load từ pw_chats.json (Playwright scrape) ─────────
        if not chats_result:
            pw_chats_file = BASE_DIR / "pw_chats.json"
            if pw_chats_file.exists():
                try:
                    pw_chats = json.loads(pw_chats_file.read_text(encoding="utf-8"))
                    if pw_chats:
                        chats_result = pw_chats
                        chat_errors.append(
                            f"✅ Loaded {len(pw_chats)} chats từ Playwright scrape (pw_chats.json)."
                        )
                except Exception as e:
                    chat_errors.append(f"⚠️ Load pw_chats.json error: {e}")
            if not chats_result:
                chat_errors.append(
                    '💡 Nhấn nút "🔍 Scrape Chats (Playwright)" bên dưới '
                    "để mở Teams và lấy danh sách chat trực tiếp."
                )

        # Nếu có chats từ Graph API, thử lấy members (best effort)
        # Skip cho chats từ Teams internal API hoặc Playwright scrape
        non_graph_sources = ("teams_internal", "pw_graph", "pw_intercept", "pw_dom")
        for chat_item in chats_result:
            if chat_item.get("_source") in non_graph_sources:
                continue  # Đã có members hoặc không query được
            try:
                mem_resp = _req.get(
                    f"https://graph.microsoft.com/v1.0/me/chats/{chat_item['chat_id']}/members",
                    headers=headers, timeout=10,
                )
                if mem_resp.status_code == 200:
                    members = mem_resp.json().get("value", [])
                    member_names = [
                        m.get("displayName", "?")
                        for m in members if m.get("displayName")
                    ]
                    chat_item["member_count"] = len(members)
                    chat_item["members"]      = member_names[:10]

                    # Cập nhật display_name nếu chưa có topic
                    if chat_item["display_name"].startswith("Chat (") and member_names:
                        chat_item["display_name"] = ", ".join(member_names[:5])
                        if len(member_names) > 5:
                            chat_item["display_name"] += f" +{len(member_names) - 5}"
            except Exception:
                pass  # Không lấy được members → giữ nguyên tên

        resp_data = {
            "status": "ok",
            "teams": result,
            "chats": chats_result,
            "token_audience": token_aud,
            "token_scopes": token_scp[:200],
        }
        if team_errors:
            resp_data["team_errors"] = team_errors
        if chat_errors and not chats_result:
            resp_data["chat_errors"] = chat_errors
        return jsonify(resp_data)
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/step2/save", methods=["POST"])
def step2_save():
    data     = request.get_json()
    selected = data.get("channels", [])
    chats    = data.get("chats", [])
    cfg      = load_config()
    cfg["channels"] = selected
    cfg["chats"]    = chats
    save_config(cfg)
    return jsonify({"status": "ok", "count": len(selected) + len(chats)})


# ── Step 3: Configure Export ──────────────────────────────────────────────────

@app.route("/step3", methods=["GET"])
def step3():
    cfg = load_config()
    return render_template(
        "step3_config.html",
        config=cfg,
        token_status=get_token_status(),
        current_step=3,
    )


@app.route("/step3", methods=["POST"])
def step3_save():
    cfg = load_config()
    cfg["browser"] = request.form.get("browser", "edge")

    exp = cfg.get("export", {})
    exp["date_from"]       = request.form.get("date_from", "")
    exp["date_to"]         = request.form.get("date_to", "")
    exp["output_dir"]      = request.form.get("output_dir", "output")
    exp["include_replies"] = request.form.get("include_replies") == "on"
    exp["mark_replies"]    = request.form.get("mark_replies") == "on"
    exp["timezone_offset"] = int(request.form.get("timezone_offset", "7"))
    exp["file_naming"]     = request.form.get("file_naming", "{team}_{channel}_{date}.xlsx")
    exp["fields"] = {
        "type":       request.form.get("field_type")       == "on",
        "datetime":   request.form.get("field_datetime")   == "on",
        "sender":     request.form.get("field_sender")     == "on",
        "content":    request.form.get("field_content")    == "on",
        "message_id": request.form.get("field_message_id") == "on",
        "raw_html":   request.form.get("field_raw_html")   == "on",
    }
    cfg["export"] = exp
    save_config(cfg)
    flash("✅ Settings saved!", "success")
    return redirect(url_for("step4"))


# ── Step 4: Export & Download ─────────────────────────────────────────────────

@app.route("/step4")
def step4():
    cfg          = load_config()
    export_state = load_export_state()
    output_files = sorted(OUTPUT_DIR.glob("*.xlsx")) if OUTPUT_DIR.exists() else []
    return render_template(
        "step4_export.html",
        config=cfg,
        export_state=export_state,
        output_files=[f.name for f in output_files],
        token_status=get_token_status(),
        current_step=4,
    )


@app.route("/step4/start", methods=["POST"])
def step4_start():
    job_id = str(uuid.uuid4())
    jobs[job_id] = {"status": "running", "progress": [], "files": []}

    cfg = load_config()

    def run_export():
        job = jobs[job_id]
        q   = job["progress"]

        def emit(obj: dict):
            q.append(json.dumps(obj))

        try:
            token   = get_valid_token(browser=cfg.get("browser", "edge"))
            headers = em.make_headers(token)

            channels  = cfg.get("channels", [])
            chats     = cfg.get("chats", [])
            exp_cfg   = cfg.get("export", {})
            date_from = exp_cfg.get("date_from", "")
            date_to   = exp_cfg.get("date_to", "")
            inc_reply = exp_cfg.get("include_replies", True)

            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            export_state = load_export_state()
            total        = len(channels) + len(chats)

            for idx, ch in enumerate(channels, 1):
                team_name    = ch.get("team_name", "unknown_team")
                channel_name = ch.get("channel_name", "unknown_channel")
                team_id      = ch["team_id"]
                channel_id   = ch["channel_id"]

                emit({
                    "type":    "progress",
                    "idx":     idx,
                    "total":   total,
                    "channel": f"{team_name} / #{channel_name}",
                    "status":  "fetching",
                })

                rows = None
                for attempt in range(3):   # 0=v1.0, 1=/beta fallback, 2=token-refresh+v1.0
                    try:
                        fetch_fn = _fetch_messages_beta if attempt == 1 else em.fetch_messages
                        rows = fetch_fn(
                            team_id, channel_id, headers,
                            date_from=date_from,
                            date_to=date_to,
                            include_replies=inc_reply,
                        )
                        if attempt == 1:
                            emit({"type": "info", "message": "ℹ️ Used /beta endpoint (v1.0 returned 403)"})
                        break
                    except PermissionError as e:
                        err = str(e)
                        if "TOKEN_EXPIRED" in err and attempt != 1:
                            emit({"type": "info", "message": "🔄 Token expired, refreshing via SSO..."})
                            new_token = get_token_sso(browser=cfg.get("browser", "edge"))
                            if new_token:
                                save_token(new_token)
                                headers = em.make_headers(new_token)
                                emit({"type": "info", "message": "✅ Token refreshed, resuming..."})
                                continue
                            else:
                                emit({"type": "error", "message": "❌ Could not refresh token"})
                                break
                        elif "ACCESS_DENIED" in err and attempt == 0:
                            emit({"type": "info", "message": "⚠️ /v1.0 returned 403 — retrying with /beta endpoint..."})
                            continue   # → attempt 1: beta fallback
                        else:
                            emit({
                                "type": "error",
                                "message": (
                                    f"⛔ Skipping — {err}\n"
                                    "Hint: In DevTools, open a channel in Teams to trigger a "
                                    "graph.microsoft.com/v1.0/teams/.../messages request, then "
                                    "copy that Authorization token to Step 1."
                                ),
                            })
                            break

                if rows is None:
                    continue

                if not rows:
                    emit({
                        "type":    "warning",
                        "channel": f"{team_name} / #{channel_name}",
                        "message": "No messages found (or filtered out by date range)",
                    })
                    continue

                safe      = _safe_filename(f"{team_name}_{channel_name}")
                today     = datetime.now().strftime("%Y-%m")
                out_path  = OUTPUT_DIR / f"{safe}_{today}.xlsx"

                em.write_excel(rows, out_path, sheet_name=_safe_sheet_name(channel_name))
                job["files"].append(out_path.name)

                state_key = f"{team_id}_{channel_id}"
                export_state[state_key] = {
                    "team_name":     team_name,
                    "channel_name":  channel_name,
                    "last_exported": datetime.now(timezone.utc).isoformat(),
                    "row_count":     len(rows),
                    "file":          out_path.name,
                }
                save_export_state(export_state)

                emit({
                    "type":    "done_channel",
                    "idx":     idx,
                    "total":   total,
                    "channel": f"{team_name} / #{channel_name}",
                    "rows":    len(rows),
                    "file":    out_path.name,
                })

            # ── Export Group Chats → DOCX ─────────────────────────────────────
            ch_count = len(channels)
            for cidx, chat in enumerate(chats, 1):
                idx = ch_count + cidx
                chat_id      = chat["chat_id"]
                display_name = chat.get("display_name", f"Chat {chat_id[:8]}…")

                emit({
                    "type":    "progress",
                    "idx":     idx,
                    "total":   total,
                    "channel": f"💬 {display_name}",
                    "status":  "fetching",
                })

                rows = None
                try:
                    rows = _fetch_chat_messages(
                        chat_id, headers,
                        date_from=date_from,
                        date_to=date_to,
                    )
                    if rows is not None:
                        emit({"type": "info", "message": f"✅ Fetched {len(rows)} messages"})
                except PermissionError as e:
                    err = str(e)
                    if "TOKEN_EXPIRED" in err:
                        emit({"type": "info", "message": "🔄 Token expired, refreshing…"})
                        new_token = get_token_sso(browser=cfg.get("browser", "edge"))
                        if new_token:
                            save_token(new_token)
                            headers = em.make_headers(new_token)
                            try:
                                rows = _fetch_chat_messages(
                                    chat_id, headers,
                                    date_from=date_from, date_to=date_to,
                                )
                            except Exception as e2:
                                emit({"type": "error", "message": f"⛔ Retry failed: {e2}"})
                    if rows is None:
                        emit({"type": "error", "message": f"⛔ Skipping chat — {err}"})
                except Exception as e:
                    emit({"type": "error", "message": f"⛔ Chat error: {e}"})

                if rows is None:
                    continue

                if not rows:
                    emit({
                        "type":    "warning",
                        "channel": f"💬 {display_name}",
                        "message": "No messages found (or filtered out by date range)",
                    })
                    continue

                # ── Write DOCX instead of Excel ──────────────────────────
                safe     = _safe_filename(f"Chat_{display_name}")
                today    = datetime.now().strftime("%Y-%m")
                out_path = OUTPUT_DIR / f"{safe}_{today}.docx"

                # Prepare auth headers for image downloading
                docx_auth = None
                try:
                    chatsvc_td = json.loads(
                        (BASE_DIR / "chatsvc_token.json").read_text(encoding="utf-8")
                    )
                    docx_auth = {"Authorization": f"Bearer {chatsvc_td.get('token', '')}"}
                except Exception:
                    pass

                def _emit_log(msg):
                    emit({"type": "info", "message": msg})

                msg_count = export_docx.write_docx(
                    messages=rows,
                    output_path=out_path,
                    chat_name=display_name,
                    auth_headers=docx_auth,
                    log_fn=_emit_log,
                )
                job["files"].append(out_path.name)

                state_key = f"chat_{chat_id}"
                export_state[state_key] = {
                    "chat_name":     display_name,
                    "last_exported": datetime.now(timezone.utc).isoformat(),
                    "row_count":     msg_count,
                    "file":          out_path.name,
                }
                save_export_state(export_state)

                emit({
                    "type":    "done_channel",
                    "idx":     idx,
                    "total":   total,
                    "channel": f"💬 {display_name}",
                    "rows":    msg_count,
                    "file":    out_path.name,
                })

            job["status"] = "complete"
            emit({"type": "complete", "files": job["files"]})

        except Exception as exc:
            jobs[job_id]["status"] = "error"
            q.append(json.dumps({"type": "error", "message": str(exc)}))

    threading.Thread(target=run_export, daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/step4/progress/<job_id>")
def step4_progress(job_id: str):
    if job_id not in jobs:
        return Response(
            'data: {"type":"error","message":"Job not found"}\n\n',
            content_type="text/event-stream",
        )

    def generate():
        job  = jobs[job_id]
        sent = 0
        while True:
            while sent < len(job["progress"]):
                yield f"data: {job['progress'][sent]}\n\n"
                sent += 1
            if job["status"] in ("complete", "error") and sent >= len(job["progress"]):
                break
            time.sleep(0.4)

    return Response(
        generate(),
        content_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/step4/download/<path:filename>")
def step4_download(filename: str):
    # Prevent path traversal: only allow bare filenames from output dir
    safe = Path(filename).name
    if not safe.endswith((".xlsx", ".docx")):
        return "Invalid file type", 400
    file_path = OUTPUT_DIR / safe
    if not file_path.exists():
        return "File not found", 404
    mimetype = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        if safe.endswith(".docx") else
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    return send_file(file_path, as_attachment=True, mimetype=mimetype)


@app.route("/step4/download_zip")
def step4_download_zip():
    if not OUTPUT_DIR.exists():
        return "No output directory", 404
    files = list(OUTPUT_DIR.glob("*.xlsx")) + list(OUTPUT_DIR.glob("*.docx"))
    if not files:
        return "No export files found", 404

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(f, f.name)
    buf.seek(0)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"teams_export_{timestamp}.zip",
        mimetype="application/zip",
    )


# ── API helpers ───────────────────────────────────────────────────────────────

@app.route("/api/token_status")
def api_token_status():
    return jsonify(get_token_status())


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("🚀  Teams Export Web — running at http://localhost:5015")
    app.run(debug=True, port=5015, threaded=True)
