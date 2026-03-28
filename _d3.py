import os
from pathlib import Path
p = Path(os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data\Default\Network\Cookies"))
out = Path(r"c:\Users\LocNTP\Downloads\ku_dev\teams_app\TeamsExportMess\_result.txt")
lines = [f"path={p}", f"exists={p.exists()}"]
if p.exists():
    lines.append(f"size={p.stat().st_size}")
try:
    import sqlite3
    conn = sqlite3.connect(str(p), timeout=3)
    cnt = conn.execute("SELECT count(*) FROM cookies").fetchone()[0]
    conn.close()
    lines.append(f"direct_ok={cnt}")
except Exception as e:
    lines.append(f"direct_fail={e}")
try:
    import sqlite3
    uri = "file:///" + str(p).replace("\\","/") + "?mode=ro&immutable=1"
    conn = sqlite3.connect(uri, uri=True, timeout=3)
    cnt = conn.execute("SELECT count(*) FROM cookies").fetchone()[0]
    conn.close()
    lines.append(f"uri_ok={cnt}")
except Exception as e:
    lines.append(f"uri_fail={e}")
out.write_text("\n".join(lines))
