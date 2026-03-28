"""Quick test for export_docx module."""
from pathlib import Path
import export_docx

msgs = [
    {
        "sender": "Loc Nguyen",
        "datetime": "2026-03-28 10:30:00",
        "content_raw": '<p>Hello team! Check this <a href="https://sharepoint.com/doc">document</a></p>',
        "segments": [
            {"type": "text", "value": "Hello team! Check this "},
            {"type": "link", "url": "https://sharepoint.com/doc", "text": "document"},
        ],
        "images": [],
        "attachments": [{"name": "Report.xlsx", "url": "https://sharepoint.com/doc"}],
        "message_id": "123",
    },
    {
        "sender": "Tanaka",
        "datetime": "2026-03-28 10:31:00",
        "content_raw": "OK got it!",
        "segments": [{"type": "text", "value": "OK got it!"}],
        "images": [],
        "attachments": [],
        "message_id": "124",
    },
]

out = Path("output/test_export.docx")
cnt = export_docx.write_docx(msgs, out, "Test Chat", log_fn=print)
print(f"Done: {cnt} messages written to {out}")
