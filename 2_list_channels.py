"""
2_list_channels.py
==================
Liệt kê tất cả Team và Channel mà tài khoản bạn có quyền truy cập.
Lưu kết quả vào all_channels.json để tiện copy vào config.json.

Cách dùng:
    python 2_list_channels.py
"""

import json
import requests
from pathlib import Path
from get_token_helper import get_valid_token   # helper dùng chung

OUTPUT_FILE = "all_channels.json"


def make_headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def fetch_joined_teams(headers: dict) -> list:
    resp = requests.get(
        "https://graph.microsoft.com/v1.0/me/joinedTeams", headers=headers
    )
    resp.raise_for_status()
    return resp.json().get("value", [])


def fetch_channels(team_id: str, headers: dict) -> list:
    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/teams/{team_id}/channels",
        headers=headers,
    )
    if resp.status_code == 403:
        return []   # Không có quyền → bỏ qua
    resp.raise_for_status()
    return resp.json().get("value", [])


def main():
    token = get_valid_token()
    headers = make_headers(token)

    print("📋 Đang lấy danh sách Team & Channel...\n")
    teams = fetch_joined_teams(headers)

    all_channels = []
    for i, team in enumerate(teams, 1):
        team_name = team["displayName"]
        team_id   = team["id"]
        print(f"[{i}] {team_name}  (id: {team_id})")

        channels = fetch_channels(team_id, headers)
        for ch in channels:
            ch_name = ch["displayName"]
            ch_id   = ch["id"]
            print(f"    ├── {ch_name:<30} (id: {ch_id})")
            all_channels.append({
                "team_name":    team_name,
                "team_id":      team_id,
                "channel_name": ch_name,
                "channel_id":   ch_id,
            })
        print()

    Path(OUTPUT_FILE).write_text(
        json.dumps(all_channels, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"✅ Đã lưu {len(all_channels)} channel vào {OUTPUT_FILE}")
    print("   → Mở file đó, chọn channel muốn export, copy vào config.json")


if __name__ == "__main__":
    main()
