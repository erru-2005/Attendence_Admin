from __future__ import annotations

import json
import os
from datetime import datetime, time, timedelta
from pathlib import Path
from typing import List, Dict, Any

from flask import Flask, jsonify, render_template, request, redirect, url_for, session


app = Flask(__name__)
# Simple secret key for session management; replace via ENV in production
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-me")


# Configure via environment variable or default path; can be overridden at runtime by querystring
ATTENDANCE_DIR = os.environ.get("ATTENDANCE_DIR", r"E:\\webface_gui\\Attendence_System\\database\\attendance")
# Where to persist enriched daily files with delay field
OUTPUT_DIR = os.environ.get("OUTPUT_DIR", r"E:\\attendence_report\\Attendence_Admin\\Credentials\\Attendence")


def get_attendance_dir() -> Path:
    # Allow overriding via query parameter for flexibility (read-only)
    override = request.args.get("dir")
    path_str = override if override else ATTENDANCE_DIR
    return Path(path_str)


def read_attendance_for_date(base_dir: Path, date_str: str) -> List[Dict[str, Any]]:
    """Read a JSON file named YYYY-MM-DD.json and return list of rows.

    This is read-only and should not interfere with other processes writing files.
    """
    # Normalize expected filename
    filename = f"{date_str}.json"
    file_path = base_dir / filename
    if not file_path.exists():
        return []
    try:
        with file_path.open("r", encoding="utf-8") as f:  # read-only
            data = json.load(f)
            # Ensure we return a list of dicts
            if isinstance(data, list):
                return data
            return []
    except Exception:
        # Fail closed on parse errors
        return []


def is_logged_in() -> bool:
    return bool(session.get("logged_in"))


# --- Delay computation helpers ---

# Explicit staff categories; any non-teaching not listed in admin is treated as supportive
TEACHING_IDS = {
    # Teaching Staff IDs (partial list based on provided data; include as given)
    "BBHCF001","BBHCF002","BBHCF003","BBHCF004","BBHCF005","BBHCF006","BBHCF007","BBHCF008",
    "BBHCF009","BBHCF010","BBHCF011","BBHCF012","BBHCF013","BBHCF014","BBHCF015","BBHCF016",
    "BBHCF017","BBHCF018","BBHCF019","BBHCF020","BBHCF021","BBHCF022","BBHCF023","BBHCF024",
    "BBHCF025","BBHCF026","BBHCF027","BBHCF028","BBHCF029","BBHCF030","BBHCF031","BBHCF032",
    "BBHCF033","BBHCF034","BBHCF035","BBHCF036","BBHCF037","BBHCF038","BBHCF039","BBHCF040",
    "BBHCF041","BBHCF042","BBHCF043","BBHCF044","BBHCF045","BBHCF046","BBHCF047","BBHCF048",
    "BBHCF049","BBHCF050","BBHCF051","BBHCF052","BBHCF053","BBHCF054","BBHCF055",
    # Some with N suffix also requested as teaching
    "BBHCFN014","BBHCFN015","BBHCFN016","BBHCFN017","BBHCFN012","BBHCFN011",
}

ADMIN_IDS = {
    # Non-Teaching Administrative Staff
    "BBHCFN002","BBHCFN003","BBHCFN004","BBHCFN005","BBHCFN006","BBHCFN007","BBHCFN008","BBHCFN009",
    "BBHCFN010","BBHCFN013",
}

# Supportive IDs provided; any other BBHCFN* not in ADMIN_IDS defaults to supportive
SUPPORTIVE_IDS = {
    "BBHCFN018","BBHCFN019","BBHCFN020","BBHCFN021","BBHCFN022","BBHCFN023","BBHCFN024","BBHCFN025",
    "BBHCFN026","BBHCFN027","BBHCFN028","BBHCFN029",
}


def classify_staff(student_id: str) -> str:
    sid = (student_id or "").upper().strip()
    if sid in TEACHING_IDS:
        return "teaching"
    if sid in ADMIN_IDS:
        return "admin"
    if sid in SUPPORTIVE_IDS or sid.startswith("BBHCFN"):
        return "supportive"
    # default unknowns (e.g., BBHCF125) to teaching unless it's clearly N-series
    return "teaching"


def get_required_window(category: str, day: datetime) -> tuple[datetime, datetime]:
    # Weekday: 0=Mon ... 5=Sat, 6=Sun (assume closed on Sunday -> zero requirement)
    weekday = day.weekday()
    date_only = day.date()
    if weekday == 6:
        start = datetime.combine(date_only, time(0, 0, 0))
        end = start  # no requirement on Sunday
        return start, end

    if category == "teaching":
        if weekday == 5:  # Saturday
            start_t, end_t = time(9, 25), time(13, 0)
        else:
            start_t, end_t = time(9, 25), time(16, 30)
    elif category == "admin":
        if weekday == 5:
            start_t, end_t = time(9, 15), time(13, 30)
        else:
            start_t, end_t = time(9, 15), time(17, 15)
    else:  # supportive
        if weekday == 5:
            start_t, end_t = time(8, 30), time(13, 30)
        else:
            start_t, end_t = time(8, 30), time(17, 15)
    return datetime.combine(date_only, start_t), datetime.combine(date_only, end_t)


def parse_iso(ts: str) -> datetime | None:
    if not ts:
        return None
    try:
        return datetime.fromisoformat(ts)
    except Exception:
        return None


def merge_intervals(intervals: List[tuple[datetime, datetime]]) -> List[tuple[datetime, datetime]]:
    if not intervals:
        return []
    intervals = [(a, b) for a, b in intervals if a and b and b > a]
    if not intervals:
        return []
    intervals.sort(key=lambda x: x[0])
    merged: List[tuple[datetime, datetime]] = []
    cur_start, cur_end = intervals[0]
    for s, e in intervals[1:]:
        if s <= cur_end:
            if e > cur_end:
                cur_end = e
        else:
            merged.append((cur_start, cur_end))
            cur_start, cur_end = s, e
    merged.append((cur_start, cur_end))
    return merged


def compute_delay_for_person(day_dt: datetime, category: str, intervals: List[tuple[datetime, datetime]]) -> int:
    req_start, req_end = get_required_window(category, day_dt)
    if req_end <= req_start:
        return 0
    required_seconds = int((req_end - req_start).total_seconds())
    if required_seconds <= 0:
        return 0
    # Clamp presence intervals to the required window
    clamped: List[tuple[datetime, datetime]] = []
    for s, e in intervals:
        s2 = max(s, req_start)
        e2 = min(e, req_end)
        if e2 > s2:
            clamped.append((s2, e2))
    merged = merge_intervals(clamped)
    present_seconds = sum(int((e - s).total_seconds()) for s, e in merged)
    delay_seconds = max(required_seconds - present_seconds, 0)
    return delay_seconds


def format_hms(total_seconds: int) -> str:
    h = total_seconds // 3600
    m = (total_seconds % 3600) // 60
    s = total_seconds % 60
    return f"{h:02d}h:{m:02d}m:{s:02d}s"


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        if username == "bbhcadmin" and password == "123456":
            session["logged_in"] = True
            return redirect(url_for("index"))
        return render_template("login.html", error="Invalid username or password")
    if is_logged_in():
        return redirect(url_for("index"))
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
def index():
    if not is_logged_in():
        return redirect(url_for("login"))
    return render_template("index.html")


@app.route("/api/attendance")
def api_attendance():
    if not is_logged_in():
        return redirect(url_for("login"))
    # Use current date by default
    date_str = request.args.get("date")
    if not date_str:
        date_str = datetime.now().strftime("%Y-%m-%d")

    base_dir = get_attendance_dir()
    rows = read_attendance_for_date(base_dir, date_str)

    # Aggregate intervals per person for the day
    day_dt = datetime.strptime(date_str, "%Y-%m-%d")
    per_id: Dict[str, Dict[str, Any]] = {}
    for r in rows:
        sid = (r.get("student_id") or "").strip()
        name = r.get("name") or ""
        ci = parse_iso(r.get("checkin") or "")
        co = parse_iso(r.get("checkout") or "")
        if sid not in per_id:
            per_id[sid] = {"student_id": sid, "name": name, "intervals": []}
        # Only count intervals with valid checkout
        if ci and co and co > ci:
            per_id[sid]["intervals"].append((ci, co))
        # Track earliest/latest for display
        if ci:
            prev_min = per_id[sid].get("min_ci")
            per_id[sid]["min_ci"] = ci if prev_min is None or ci < prev_min else prev_min
        if co:
            prev_max = per_id[sid].get("max_co")
            per_id[sid]["max_co"] = co if prev_max is None or co > prev_max else prev_max

    enriched: List[Dict[str, Any]] = []
    for sid, info in per_id.items():
        category = classify_staff(sid)
        intervals = info.get("intervals", [])
        delay_sec = compute_delay_for_person(day_dt, category, intervals)
        enriched.append({
            "student_id": sid,
            "name": info.get("name", ""),
            "checkin": (info.get("min_ci").isoformat() if info.get("min_ci") else ""),
            "checkout": (info.get("max_co").isoformat() if info.get("max_co") else ""),
            "delay": format_hms(delay_sec),
        })

    # Persist enriched JSON for the date
    try:
        out_dir = Path(OUTPUT_DIR)
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"{date_str}.json"
        with out_path.open("w", encoding="utf-8") as f:
            json.dump(enriched, f, ensure_ascii=False, indent=2)
    except Exception:
        # Non-fatal if persistence fails
        pass

    # Return enriched rows to UI
    return jsonify(enriched)


if __name__ == "__main__":
    app.run(debug=True)


