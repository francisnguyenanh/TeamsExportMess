"""
get_token_helper.py
===================
Module dùng chung: load token từ file, tự làm mới nếu hết hạn.
Được import bởi 2_list_channels.py và 3_export.py.
"""

import json
from pathlib import Path


def get_valid_token(browser: str = None) -> str:
    """
    Trả về token hợp lệ:
    - Nếu token.json còn hạn → dùng luôn
    - Nếu hết hạn → tự động gọi SSO để lấy mới

    Tham số browser đọc từ config.json nếu không truyền vào.
    """
    # Đọc browser từ config.json nếu có
    if browser is None:
        try:
            config = json.loads(Path("config.json").read_text(encoding="utf-8"))
            browser = config.get("browser", "edge")
        except Exception:
            browser = "edge"

    # Import ở đây để tránh circular import
    from get_token import load_token, get_token_sso, save_token   # noqa: E402  (1_get_token.py)

    token = load_token()
    if not token:
        print(f"🔄 Token hết hạn/chưa có → tự động lấy mới qua SSO ({browser})...")
        token = get_token_sso(browser=browser)
        if not token:
            raise RuntimeError(
                "❌ Không lấy được token!\n"
                "   Thử chạy thủ công: python 1_get_token.py"
            )
        save_token(token)
    return token
