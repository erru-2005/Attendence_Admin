from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any

from flask import Flask, jsonify, render_template, request


app = Flask(__name__)


# Configure via environment variable or default path; can be overridden at runtime by querystring
ATTENDANCE_DIR = os.environ.get("ATTENDANCE_DIR", r"G:\\database\\attendance")


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


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/attendance")
def api_attendance():
    # Use current date by default
    date_str = request.args.get("date")
    if not date_str:
        date_str = datetime.now().strftime("%Y-%m-%d")

    base_dir = get_attendance_dir()

    rows = read_attendance_for_date(base_dir, date_str)

    # Backend uses provided keys from files (student_id etc.),
    # but the UI labels them as Faculty as per requirement.
    return jsonify(rows)


if __name__ == "__main__":
    app.run(debug=True)


