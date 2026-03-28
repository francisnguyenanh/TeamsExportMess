"""Debug: tìm và thử mở Cookies DB của Chrome."""
import os
import sqlite3
from pathlib import Path

profile = Path(os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data"))
print(f"Profile dir: {profile}")
print(f"  exists: {profile.exists()}")

# Tìm tất cả Cookies file
print("\n== All Cookies files ==")
found = list(profile.rglob("Cookies"))
for f in found:
    try:
        sz = f.stat().st_size
    except Exception:
        sz = "???"
    print(f"  {f}  ({sz} bytes)")

if not found:
    print("  NONE FOUND!")

# Thử mở từng cách
for label, p in [
    ("Default/Network/Cookies", profile / "Default" / "Network" / "Cookies"),
    ("Default/Cookies",         profile / "Default" / "Cookies"),
]:
    print(f"\n== Try: {label} ==")
    print(f"  path: {p}")
    print(f"  exists: {p.exists()}")
    if not p.exists():
        continue

    # Cách 1: immutable URI
    try:
        uri = f"file:{p}?mode=ro&immutable=1"
        print(f"  URI: {uri}")
        conn = sqlite3.connect(uri, uri=True, timeout=3)
        count = conn.execute("SELECT count(*) FROM cookies").fetchone()[0]
        conn.close()
        print(f"  ✅ OK (immutable) — {count} cookies")
    except Exception as e:
        print(f"  ❌ immutable FAIL: {e}")

    # Cách 2: immutable URI (forward slashes)
    try:
        p_fwd = str(p).replace("\\", "/")
        uri2 = f"file:///{p_fwd}?mode=ro&immutable=1"
        print(f"  URI2: {uri2}")
        conn = sqlite3.connect(uri2, uri=True, timeout=3)
        count = conn.execute("SELECT count(*) FROM cookies").fetchone()[0]
        conn.close()
        print(f"  ✅ OK (URI2) — {count} cookies")
    except Exception as e:
        print(f"  ❌ URI2 FAIL: {e}")

    # Cách 3: direct path, no URI
    try:
        conn = sqlite3.connect(str(p), timeout=3)
        count = conn.execute("SELECT count(*) FROM cookies").fetchone()[0]
        conn.close()
        print(f"  ✅ OK (direct) — {count} cookies")
    except Exception as e:
        print(f"  ❌ direct FAIL: {e}")
