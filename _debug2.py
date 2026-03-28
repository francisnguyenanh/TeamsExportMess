import os, sqlite3, json
from pathlib import Path

results = []
profile = Path(os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data"))
results.append(f"profile={profile}, exists={profile.exists()}")

# Find Cookies files
for f in profile.rglob("Cookies"):
    try:
        sz = f.stat().st_size
    except:
        sz = -1
    results.append(f"found: {f} ({sz} bytes)")

# Try to open
paths_to_try = [
    profile / "Default" / "Network" / "Cookies",
    profile / "Default" / "Cookies",
]
for p in paths_to_try:
    results.append(f"try: {p}, exists={p.exists()}")
    if not p.exists():
        continue
    
    # Method: forward-slash URI
    try:
        p_str = str(p).replace("\\", "/")
        uri = f"file:///{p_str}?mode=ro&immutable=1"
        conn = sqlite3.connect(uri, uri=True, timeout=3)
        cnt = conn.execute("SELECT count(*) FROM cookies").fetchone()[0]
        ms_cnt = conn.execute("SELECT count(*) FROM cookies WHERE host_key LIKE '%.microsoft.com'").fetchone()[0]
        conn.close()
        results.append(f"  OK: {cnt} total, {ms_cnt} microsoft cookies")
    except Exception as e:
        results.append(f"  FAIL: {e}")

# Write to file
out = Path(__file__).parent / "_debug_result.json"
out.write_text(json.dumps(results, indent=2, ensure_ascii=False))
