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
    return render_template(
        "step2_channels.html",
        teams=teams,
        config=cfg,
        export_state=export_state,
        selected_ids=selected_ids,
        token_status=get_token_status(),
        current_step=2,
    )


@app.route("/step2/refresh", methods=["POST"])
def step2_refresh():
    try:
        token   = get_valid_token()
        headers = lc.make_headers(token)
        teams_list = lc.fetch_joined_teams(headers)

        result       = []
        flat_channels = []
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

        ALL_CHANNELS_FILE.write_text(
            json.dumps(flat_channels, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return jsonify({"status": "ok", "teams": result})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/step2/save", methods=["POST"])
def step2_save():
    data     = request.get_json()
    selected = data.get("channels", [])
    cfg      = load_config()
    cfg["channels"] = selected
    save_config(cfg)
    return jsonify({"status": "ok", "count": len(selected)})


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
            exp_cfg   = cfg.get("export", {})
            date_from = exp_cfg.get("date_from", "")
            date_to   = exp_cfg.get("date_to", "")
            inc_reply = exp_cfg.get("include_replies", True)

            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            export_state = load_export_state()
            total        = len(channels)

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
                for attempt in range(2):
                    try:
                        rows = em.fetch_messages(
                            team_id, channel_id, headers,
                            date_from=date_from,
                            date_to=date_to,
                            include_replies=inc_reply,
                        )
                        break
                    except PermissionError as e:
                        err = str(e)
                        if "TOKEN_EXPIRED" in err and attempt == 0:
                            emit({"type": "info", "message": "🔄 Token expired, refreshing via SSO..."})
                            new_token = get_token_sso(browser=cfg.get("browser", "edge"))
                            if new_token:
                                save_token(new_token)
                                headers = em.make_headers(new_token)
                                emit({"type": "info", "message": "✅ Token refreshed, resuming..."})
                            else:
                                emit({"type": "error", "message": "❌ Could not refresh token"})
                                break
                        else:
                            emit({"type": "error", "message": f"⛔ Skipping — {err}"})
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

                safe      = re.sub(r'[\\/*?:"<>|]', "_", f"{team_name}_{channel_name}")
                today     = datetime.now().strftime("%Y-%m")
                out_path  = OUTPUT_DIR / f"{safe}_{today}.xlsx"

                em.write_excel(rows, out_path, sheet_name=channel_name)
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
    if not safe.endswith(".xlsx"):
        return "Invalid file type", 400
    file_path = OUTPUT_DIR / safe
    if not file_path.exists():
        return "File not found", 404
    return send_file(file_path, as_attachment=True)


@app.route("/step4/download_zip")
def step4_download_zip():
    if not OUTPUT_DIR.exists():
        return "No output directory", 404
    files = list(OUTPUT_DIR.glob("*.xlsx"))
    if not files:
        return "No Excel files found", 404

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
