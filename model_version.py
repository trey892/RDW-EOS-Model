"""
Tracks a simple incrementing version number for RDW_EOS_Master_latest.xlsx,
stored in output/model_version.json alongside it. Introduced 2026-08-08 when
the model moved off numbered filenames (v5/v6/v7/v8) to a single canonical
"latest" file for the automated daily refresh -- version numbering restarts
at 1 here rather than continuing the old v8 sequence, since this is a new,
separate tracking mechanism (a counter of refresh EVENTS, not file variants).

Call bump() at the end of any script that actually changes
RDW_EOS_Master_latest.xlsx (refresh_model.py, refresh_fuel_history.py, etc).
Call read() from scripts that just need the current value (e.g.
build_dashboard_data.py) without incrementing it.
"""
import datetime
import json
from pathlib import Path

VERSION_JSON = Path(__file__).parent / "output" / "model_version.json"


def read():
    if not VERSION_JSON.exists():
        return {"version": 0, "lastRefreshed": None}
    return json.loads(VERSION_JSON.read_text(encoding="utf-8"))


def bump():
    current = read()
    new = {"version": current["version"] + 1, "lastRefreshed": datetime.date.today().isoformat()}
    VERSION_JSON.write_text(json.dumps(new), encoding="utf-8")
    return new
