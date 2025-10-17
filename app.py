from __future__ import annotations

import json
import os
import shutil
import threading
import time as time_module
from datetime import datetime, time as dt_time
from pathlib import Path
from typing import List, Dict, Any, Tuple

from flask import Flask, jsonify, render_template, request, redirect, url_for, session


app = Flask(__name__)
# Simple secret key for session management; replace via ENV in production
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-me")


# Configure via environment variable or default path; can be overridden at runtime by querystring
ATTENDANCE_DIR = os.environ.get("ATTENDANCE_DIR", r"E:\webface_gui\Attendence_System\database\attendance")

# JSON storage directory for synced attendance data - DYNAMIC and NON-CHANGEABLE
# Always relative to current working directory
JSON_ATTENDANCE_DIR = Path.cwd() / "JSON" / "Attendence"
JSON_ATTENDANCE_DIR.mkdir(parents=True, exist_ok=True)

# File tracking for change detection - DYNAMIC and NON-CHANGEABLE
TRACKING_FILE = Path.cwd() / "JSON" / "file_tracking.json"

# Background monitoring
monitoring_active = False
monitor_thread = None


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


def load_file_tracking() -> Dict[str, str]:
    """Load file modification tracking data."""
    if not TRACKING_FILE.exists():
        return {}
    try:
        with TRACKING_FILE.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_file_tracking(tracking_data: Dict[str, str]) -> None:
    """Save file modification tracking data."""
    try:
        with TRACKING_FILE.open("w", encoding="utf-8") as f:
            json.dump(tracking_data, f, indent=2)
    except Exception:
        pass


def get_file_modification_time(file_path: Path) -> str:
    """Get file modification time as formatted string DD-MM-YYYY HH:MM:SS."""
    try:
        mtime = file_path.stat().st_mtime
        dt = datetime.fromtimestamp(mtime)
        return dt.strftime("%d-%m-%Y %H:%M:%S")
    except Exception:
        return "01-01-1970 00:00:00"


def sync_attendance_files() -> Dict[str, Any]:
    """Sync attendance files from source to JSON directory."""
    source_dir = Path(ATTENDANCE_DIR)
    if not source_dir.exists():
        return {"success": False, "error": f"Source directory not found: {ATTENDANCE_DIR}"}
    
    tracking_data = load_file_tracking()
    synced_files = []
    updated_files = []
    errors = []
    
    try:
        # Get all JSON files from source directory
        source_files = list(source_dir.glob("*.json"))
        
        for source_file in source_files:
            try:
                # Get file modification time
                current_mtime = get_file_modification_time(source_file)
                file_key = source_file.name
                
                # Check if file has been modified
                last_mtime = tracking_data.get(file_key, "01-01-1970 00:00:00")
                
                if current_mtime > last_mtime:
                    # File has been modified or is new
                    dest_file = JSON_ATTENDANCE_DIR / source_file.name
                    
                    # Copy file to JSON directory
                    shutil.copy2(source_file, dest_file)
                    
                    # Annotate delay for that date file after copy
                    try:
                        annotate_file_with_delay(dest_file)
                    except Exception:
                        # non-fatal
                        pass
                    
                    # Update tracking data
                    tracking_data[file_key] = current_mtime
                    
                    if last_mtime != "01-01-1970 00:00:00":
                        updated_files.append(file_key)
                    else:
                        synced_files.append(file_key)
                        
            except Exception as e:
                errors.append(f"Error processing {source_file.name}: {str(e)}")
        
        # Save updated tracking data
        save_file_tracking(tracking_data)
        
        return {
            "success": True,
            "synced_files": synced_files,
            "updated_files": updated_files,
            "total_processed": len(source_files),
            "errors": errors
        }
        
    except Exception as e:
        return {"success": False, "error": f"Sync failed: {str(e)}"}


def get_attendance_from_json(date_str: str) -> List[Dict[str, Any]]:
    """Read attendance data from JSON directory."""
    json_file = JSON_ATTENDANCE_DIR / f"{date_str}.json"
    if not json_file.exists():
        return []
    
    try:
        with json_file.open("r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, list):
                return data
            return []
    except Exception:
        return []


# ---------- Delay annotation logic ----------
FACULTY_DETAIL_PATH = Path.cwd() / "JSON" / "faculty_detail.json"

_faculty_cache: Dict[str, Dict[str, Any]] | None = None

def load_faculty_details() -> Dict[str, Dict[str, Any]]:
    global _faculty_cache
    if _faculty_cache is not None:
        return _faculty_cache
    try:
        with FACULTY_DETAIL_PATH.open("r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, dict):
                _faculty_cache = data
                return data
    except Exception:
        pass
    _faculty_cache = {}
    return _faculty_cache

def determine_staff_type(faculty_id: str) -> str:
    """Return 'teaching', 'admin', or 'support' using faculty_detail.json.
    - category == 'Teaching Faculty' => teaching
    - category == 'Non-Teaching Faculty' and department includes 'Administrative' => admin
    - department == 'Computer Applications' => support (per requirement)
    - department includes 'Support' => support
    Defaults to teaching if unknown.
    """
    details = load_faculty_details().get((faculty_id or '').strip().upper()) or {}
    category = (details.get('category') or '').strip().lower()
    department = (details.get('department') or '').strip().lower()
    if 'teaching' in category:
        return 'teaching'
    if 'non-teaching' in category:
        if 'computer applications' in department:
            return 'support'
        if 'support' in department:
            return 'support'
        if 'administrative' in department or 'administration' in department:
            return 'admin'
        # default non-teaching → admin
        return 'admin'
    # fallback
    return 'teaching'

def get_thresholds_for(date_obj: datetime, staff_type: str) -> Tuple[dt_time, dt_time]:
    weekday = date_obj.weekday()  # 0=Mon ... 5=Sat ... 6=Sun
    is_saturday = (weekday == 5)
    if staff_type == 'teaching':
        checkin_deadline = dt_time(9, 25)
        checkout_after = dt_time(13, 0) if is_saturday else dt_time(16, 30)
    elif staff_type == 'admin':
        checkin_deadline = dt_time(9, 15)
        checkout_after = dt_time(13, 30) if is_saturday else dt_time(17, 15)
    else:  # support
        checkin_deadline = dt_time(8, 30)
        checkout_after = dt_time(13, 30) if is_saturday else dt_time(17, 15)
    return checkin_deadline, checkout_after

def parse_ts(value: Any) -> datetime | None:
    if not value:
        return None
    if isinstance(value, (int, float)):
        try:
            return datetime.fromtimestamp(float(value) / (1000.0 if float(value) > 1e12 else 1.0))
        except Exception:
            return None
    if isinstance(value, str):
        for fmt in ("%Y-%m-%dT%H:%M:%S.%fZ", "%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
            try:
                return datetime.strptime(value, fmt)
            except Exception:
                continue
        try:
            return datetime.fromisoformat(value)
        except Exception:
            return None
    return None

def _seconds_to_hhmmss(total_seconds: int) -> str:
    if total_seconds < 0:
        total_seconds = 0
    hours = total_seconds // 3600
    minutes_only = (total_seconds % 3600) // 60
    seconds_only = total_seconds % 60
    return f"{hours:02d}:{minutes_only:02d}:{seconds_only:02d}"


def compute_daily_delay_for_records(records: List[Dict[str, Any]], date_obj: datetime, staff_type: str) -> str:
    """Compute total delay as HH:MM:SS for a day's records for a single faculty.
    Uses earliest check-in and latest check-out. If no checkout present in any record, returns 'N/A'.
    """
    checkins: List[datetime] = []
    checkouts: List[datetime] = []
    for r in records:
        ci = parse_ts(r.get('checkin'))
        co = parse_ts(r.get('checkout'))
        if ci:
            checkins.append(ci)
        if co:
            checkouts.append(co)
    if not checkins:
        return _seconds_to_hhmmss(0)  # no checkin → treat as 0 delay
    earliest_in = min(checkins)
    checkin_deadline, checkout_after = get_thresholds_for(date_obj, staff_type)
    late_seconds = 0
    if earliest_in:
        deadline_dt = datetime.combine(earliest_in.date(), checkin_deadline)
        if earliest_in > deadline_dt:
            late_seconds = int((earliest_in - deadline_dt).total_seconds())
    if not checkouts:
        return 'N/A'
    latest_out = max(checkouts)
    required_dt = datetime.combine(latest_out.date(), checkout_after)
    early_leave_seconds = 0
    if latest_out < required_dt:
        early_leave_seconds = int((required_dt - latest_out).total_seconds())
    return _seconds_to_hhmmss(late_seconds + early_leave_seconds)

def annotate_file_with_delay(json_file: Path) -> None:
    """Read a date file, compute delay per faculty for that date, assign same delay to all their records."""
    try:
        basename = json_file.stem  # YYYY-MM-DD
        date_obj = datetime.strptime(basename, "%Y-%m-%d")
    except Exception:
        return
    try:
        with json_file.open('r', encoding='utf-8') as f:
            data = json.load(f)
        if not isinstance(data, list):
            return
        # normalize possible 'timestamp' only schema → treat as checkin-only
        for row in data:
            if 'checkin' not in row and isinstance(row.get('timestamp'), str):
                row['checkin'] = row.get('timestamp')
            if 'checkout' not in row:
                row['checkout'] = row.get('checkout', '')
        # group by student_id
        by_id: Dict[str, List[Dict[str, Any]]] = {}
        for row in data:
            sid = (row.get('student_id') or '').strip().upper()
            by_id.setdefault(sid, []).append(row)
        # compute and annotate
        for sid, rows in by_id.items():
            staff_type = determine_staff_type(sid)
            delay_val = compute_daily_delay_for_records(rows, date_obj, staff_type)
            for r in rows:
                r['delay'] = delay_val
        with json_file.open('w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        # ignore annotation errors
        pass

def annotate_all_existing_files() -> Dict[str, Any]:
    """Annotate all JSON files currently present in JSON_ATTENDANCE_DIR."""
    results = {"annotated": [], "errors": []}
    try:
        for jf in JSON_ATTENDANCE_DIR.glob('*.json'):
            try:
                annotate_file_with_delay(jf)
                results["annotated"].append(jf.name)
            except Exception as e:
                results["errors"].append({"file": jf.name, "error": str(e)})
    except Exception as e:
        results["errors"].append({"file": "*", "error": str(e)})
    return results

def monitor_directory_changes():
    """Background thread to monitor source directory for changes."""
    global monitoring_active
    
    print("Starting directory monitoring...")
    source_dir = Path(ATTENDANCE_DIR)
    
    while monitoring_active:
        try:
            if source_dir.exists():
                # Check for changes and sync if needed
                result = sync_attendance_files()
                if result.get("success") and (result.get("synced_files") or result.get("updated_files")):
                    print(f"Auto-sync completed: {len(result.get('synced_files', []))} new, {len(result.get('updated_files', []))} updated")
            
            # Sleep for 30 seconds before next check
            time_module.sleep(30)
            
        except Exception as e:
            print(f"Error in directory monitoring: {e}")
            time_module.sleep(60)  # Wait longer on error
    
    print("Directory monitoring stopped.")


def start_monitoring():
    """Start background monitoring thread."""
    global monitoring_active, monitor_thread
    
    if not monitoring_active:
        monitoring_active = True
        monitor_thread = threading.Thread(target=monitor_directory_changes, daemon=True)
        monitor_thread.start()
        print("Background monitoring started.")


def stop_monitoring():
    """Stop background monitoring thread."""
    global monitoring_active, monitor_thread
    
    if monitoring_active:
        monitoring_active = False
        if monitor_thread:
            monitor_thread.join(timeout=5)
        print("Background monitoring stopped.")


def initial_sync():
    """Perform initial sync on server start."""
    print("Performing initial sync...")
    result = sync_attendance_files()
    if result.get("success"):
        print(f"Initial sync completed: {result.get('total_processed', 0)} files processed")
        if result.get("synced_files"):
            print(f"New files synced: {result['synced_files']}")
        if result.get("updated_files"):
            print(f"Updated files: {result['updated_files']}")
        # After syncing, ensure all JSON files have delay annotated
        ann = annotate_all_existing_files()
        print(f"Delay annotated for {len(ann.get('annotated', []))} files at startup")
    else:
        print(f"Initial sync failed: {result.get('error', 'Unknown error')}")
    
    return result


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

    # Try to get data from JSON directory first
    rows = get_attendance_from_json(date_str)
    
    # If not found in JSON directory, fallback to source directory
    if not rows:
        base_dir = get_attendance_dir()
        rows = read_attendance_for_date(base_dir, date_str)

    # Backend uses provided keys from files (student_id etc.),
    # but the UI labels them as Faculty as per requirement.
    return jsonify(rows)


@app.route("/api/sync")
def api_sync():
    """Sync attendance files from source to JSON directory."""
    if not is_logged_in():
        return redirect(url_for("login"))
    
    result = sync_attendance_files()
    # Ensure annotation for all existing files after each sync
    ann = annotate_all_existing_files()
    result["annotation"] = {"annotated": len(ann.get("annotated", [])), "errors": ann.get("errors", [])}
    return jsonify(result)


@app.route("/api/sync/status")
def api_sync_status():
    """Get sync status and file information."""
    if not is_logged_in():
        return redirect(url_for("login"))
    
    source_dir = Path(ATTENDANCE_DIR)
    json_dir = JSON_ATTENDANCE_DIR
    tracking_data = load_file_tracking()
    
    source_files = []
    json_files = []
    
    if source_dir.exists():
        source_files = [f.name for f in source_dir.glob("*.json")]
    
    if json_dir.exists():
        json_files = [f.name for f in json_dir.glob("*.json")]
    
    return jsonify({
        "source_directory": str(source_dir),
        "json_directory": str(json_dir),
        "source_files_count": len(source_files),
        "json_files_count": len(json_files),
        "tracked_files": len(tracking_data),
        "monitoring_active": monitoring_active,
        "source_files": source_files,
        "json_files": json_files
    })


@app.route("/api/sync/start")
def api_start_monitoring():
    """Start background monitoring."""
    if not is_logged_in():
        return redirect(url_for("login"))
    
    start_monitoring()
    return jsonify({"success": True, "message": "Monitoring started"})


@app.route("/api/sync/stop")
def api_stop_monitoring():
    """Stop background monitoring."""
    if not is_logged_in():
        return redirect(url_for("login"))
    
    stop_monitoring()
    return jsonify({"success": True, "message": "Monitoring stopped"})


if __name__ == "__main__":
    # Perform initial sync on server start
    initial_sync()
    
    # Start background monitoring
    start_monitoring()
    
    try:
        app.run(debug=True)
    finally:
        # Stop monitoring when server shuts down
        stop_monitoring()


