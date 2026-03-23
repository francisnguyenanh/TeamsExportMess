# Microsoft Teams Export Tool
> Xuất tin nhắn Teams ra Excel — không cần IT, không cần app registration

---

## Mục lục

1. [Cấu trúc project](#1-cấu-trúc-project)
2. [Cài đặt](#2-cài-đặt)
3. [Lấy Bearer Token](#3-lấy-bearer-token)
   - [Cách A: Thủ công qua DevTools](#cách-a-thủ-công-qua-devtools-nhanh-nhất)
   - [Cách B: Script Python — Email/Password](#cách-b-script-python--emailpassword)
   - [Cách C: Script Python — SSO Windows ⭐ Khuyến nghị](#cách-c-script-python--sso-windows--khuyến-nghị)
4. [Lấy danh sách Team/Channel ID](#4-lấy-danh-sách-teamchannel-id)
5. [Cấu hình config.json](#5-cấu-hình-configjson)
6. [Chạy export](#6-chạy-export)
7. [Xử lý token hết hạn](#7-xử-lý-token-hết-hạn)
8. [Kết quả Excel](#8-kết-quả-excel)
9. [Troubleshooting](#9-troubleshooting)

---

## 1. Cấu trúc project

```
teams_export/
│
├── README.md               ← tài liệu này
├── config.json             ← danh sách channel cần export + cấu hình
├── token.json              ← token được lưu tự động (đừng commit lên git)
│
├── 1_get_token.py          ← lấy token (SSO hoặc password)
├── 2_list_channels.py      ← liệt kê tất cả team/channel → điền vào config
├── 3_export.py             ← export tin nhắn ra Excel
│
└── output/                 ← file Excel xuất ra ở đây
    └── DuAnAlpha_QA_2024-03.xlsx
```

---

## 2. Cài đặt

```bash
# Thư viện Python
pip install requests openpyxl playwright python-dotenv

# Browser driver cho Playwright (chọn 1)
playwright install msedge      # nếu dùng Edge  ← khuyến nghị
playwright install chromium    # nếu dùng Chrome
```

> **Yêu cầu:** Python 3.8+, Windows (do SSO dựa trên Windows login)

---

## 3. Lấy Bearer Token

> Token là chuỗi JWT dài ~1000 ký tự, đại diện cho phiên đăng nhập của bạn.
> Token thường hết hạn sau **60–90 phút**.

### Cách A: Thủ công qua DevTools _(nhanh nhất)_

Dùng khi muốn test nhanh, không cần cài thêm gì.

**Bước 1:** Mở `https://teams.microsoft.com` → đăng nhập

**Bước 2:** Nhấn `F12` → tab **Network** → tick **Preserve log**

**Bước 3:** Trong ô Filter gõ: `graph.microsoft.com`

**Bước 4:** Click vào bất kỳ channel nào trong Teams để trigger request

**Bước 5:** Click vào 1 request bất kỳ → tab **Headers** → **Request Headers**

**Bước 6:** Tìm dòng:
```
Authorization: Bearer eyJ0eXAiOiJKV1Qi...
```
Copy toàn bộ phần sau `Bearer ` (chuỗi rất dài)

**Bước 7:** Lưu vào file `token.json`:
```json
{
  "token": "eyJ0eXAiOiJKV1Qi...",
  "fetched_at": "2024-03-15T09:00:00+00:00",
  "expires_at": null
}
```

---

### Cách B: Script Python — Email/Password

> Dùng khi tài khoản đăng nhập bằng email/password thông thường (không có MFA).
> ⚠️ **Không hoạt động nếu công ty bật MFA/Conditional Access.**

```python
# Chạy: python 1_get_token.py --method password
```

Script sẽ mở browser, tự điền email/password, bắt token.
Xem chi tiết trong file `1_get_token.py` — phần `get_token_password()`.

**Lưu ý bảo mật:** Password không được lưu vào file nào — chỉ truyền trực tiếp qua Playwright.

---

### Cách C: Script Python — SSO Windows ⭐ _Khuyến nghị_

> Dùng khi máy tính đã đăng nhập vào domain công ty (Windows login = Teams login).
> ✅ Hoạt động kể cả khi có MFA — vì Windows đã xác thực rồi.

```bash
python 1_get_token.py --method sso
# hoặc mặc định (không truyền tham số)
python 1_get_token.py
```

**Cách hoạt động:**
1. Playwright mở Edge/Chrome với **profile Windows hiện tại** (giữ nguyên session SSO)
2. Truy cập `teams.microsoft.com` → SSO tự xử lý, không cần nhập gì
3. Script bắt Bearer token từ network request đầu tiên tới `graph.microsoft.com`
4. Lưu token + thời gian hết hạn vào `token.json`

```
✅ Đã bắt được Bearer token!
⏱  Token hết hạn lúc: 10:35:22 (còn ~87 phút)
```

---

## 4. Lấy danh sách Team/Channel ID

Chạy script sau **một lần** để xem tất cả team/channel bạn có quyền truy cập:

```bash
python 2_list_channels.py
```

Output mẫu:
```
📋 Danh sách tất cả Team & Channel:

[1] Dự án Alpha  (id: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)
    ├── General          (id: 19:aaa@thread.tacv2)
    ├── QA Hỏi Đáp       (id: 19:bbb@thread.tacv2)
    └── Deploy & Release (id: 19:ccc@thread.tacv2)

[2] Dự án Beta  (id: yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy)
    ├── General          (id: 19:ddd@thread.tacv2)
    └── Hỏi đáp kỹ thuật (id: 19:eee@thread.tacv2)

✅ Đã lưu toàn bộ vào all_channels.json
   → Copy channel muốn export vào config.json
```

---

## 5. Cấu hình `config.json`

```json
{
  "browser": "edge",
  "channels": [
    {
      "team_name": "Dự án Alpha",
      "team_id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
      "channel_name": "QA Hỏi Đáp",
      "channel_id": "19:bbb@thread.tacv2"
    },
    {
      "team_name": "Dự án Beta",
      "team_id": "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy",
      "channel_name": "Hỏi đáp kỹ thuật",
      "channel_id": "19:eee@thread.tacv2"
    }
  ],
  "export": {
    "date_from": "2024-01-01",
    "date_to": "",
    "output_dir": "output",
    "include_replies": true
  }
}
```

| Trường | Mô tả |
|--------|-------|
| `browser` | `"edge"` hoặc `"chrome"` |
| `date_from` | Lọc từ ngày (YYYY-MM-DD), để trống = lấy tất cả |
| `date_to` | Lọc đến ngày, để trống = đến hiện tại |
| `include_replies` | `true` = lấy cả reply trong thread |

---

## 6. Chạy export

```bash
# Bước 1: Lấy token (tự động qua SSO)
python 1_get_token.py

# Bước 2: Xem danh sách channel (chỉ cần làm 1 lần)
python 2_list_channels.py

# Bước 3: Export tất cả channel trong config.json
python 3_export.py
```

---

## 7. Xử lý token hết hạn

Script `3_export.py` tự động xử lý token hết hạn:

```
Token còn hạn (~45 phút)  →  dùng luôn, không mở browser
Token hết hạn hoặc gần hết  →  tự động gọi lại SSO, lấy token mới
Token hết hạn giữa chừng  →  bắt lỗi 401, làm mới token, retry channel đó
```

**Khi có nhiều group/channel:**
- Mỗi channel export tuần tự
- Nếu token hết giữa channel X → làm mới token → tiếp tục từ channel X (không mất dữ liệu đã export)
- Mỗi channel lưu ra **file Excel riêng** → an toàn, không mất dữ liệu cũ

---

## 8. Kết quả Excel

Mỗi channel tạo 1 file: `output/{TeamName}_{ChannelName}_{YYYY-MM}.xlsx`

| Type | DateTime | Sender | Content | MessageID |
|------|----------|--------|---------|-----------|
| Message | 2024-03-01 09:15:00 | Nguyễn A | Bạn ơi lỗi này fix thế nào? | abc123 |
| Reply | 2024-03-01 09:20:00 | Trần B | Bạn thử restart service xem | abc123 |
| Reply | 2024-03-01 09:25:00 | Nguyễn A | Được rồi, cảm ơn bạn! | abc123 |
| Message | 2024-03-01 10:00:00 | Lê C | Deploy lên staging bị lỗi 500 | def456 |

> **MessageID** giống nhau = cùng 1 thread (1 câu hỏi + nhiều reply)

---

## 9. Troubleshooting

| Lỗi | Nguyên nhân | Giải pháp |
|-----|-------------|-----------|
| `401 Unauthorized` | Token hết hạn | Chạy lại `python 1_get_token.py` |
| `403 Forbidden` | Không có quyền vào channel | Kiểm tra lại bạn có phải member không |
| `404 Not Found` | Team/Channel ID sai | Chạy lại `2_list_channels.py` để lấy ID mới |
| Browser không mở được | Playwright chưa cài driver | Chạy `playwright install msedge` |
| SSO không tự đăng nhập | Profile path sai | Đổi `browser = "chrome"` trong config |
| Token bị `None` | Teams chưa load kịp | Tăng `timeout` trong `1_get_token.py` lên 60000 |

---

> 💡 **Tip:** Thêm `token.json` và `output/` vào `.gitignore` nếu dùng git.
