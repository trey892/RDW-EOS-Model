"""
Tracks a day.refresh version number for RDW_EOS_Master_latest.xlsx, stored in
output/model_version.json alongside it. Introduced 2026-08-08 when the model
moved off numbered filenames (v5/v6/v7/v8) to a single canonical "latest"
file for the automated refresh -- version numbering restarts at 1 here
rather than continuing the old v8 sequence, since this is a new, separate
tracking mechanism (a counter of refresh EVENTS, not file variants).

Since 2026-08-19 the refresh runs hourly rather than once a day, so the
version is major.minor: major increments once on the first refresh of a new
calendar day (e.g. v7); minor increments on every subsequent refresh that
same day (v7.1, v7.2, ...) and resets to 0 (shown bare, no ".0" suffix) when
the day rolls over. Day boundary is judged off the *previous* bump's local
calendar date, not a fixed refresh count, so it stays correct regardless of
how many refreshes actually ran that day.

Call bump() at the end of any script that actually changes
RDW_EOS_Master_latest.xlsx (refresh_model.py, refresh_fuel_history.py, etc).
Call read() from scripts that just need the current value (e.g.
build_dashboard_data.py) without incrementing it.
"""
import datetime
import json
from pathlib import Path

VERSION_JSON = Path(__file__).parent / "output" / "model_version.json"


def _version_str(major, minor):
    return str(major) if minor == 0 else f"{major}.{minor}"


def read():
    if not VERSION_JSON.exists():
        return {"major": 0, "minor": 0, "version": "0", "lastRefreshed": None}
    return json.loads(VERSION_JSON.read_text(encoding="utf-8"))


def bump():
    current = read()
    now = datetime.datetime.now()
    today = now.date().isoformat()

    # last_date handles both the new "major"/"minor" schema and the old
    # schema (plain int "version", date-only "lastRefreshed").
    last_date = (current.get("lastRefreshed") or "")[:10]
    prior_major = current.get("major")
    if prior_major is None:
        prior_major = current.get("version") or 0
        if isinstance(prior_major, str):
            prior_major = int(float(prior_major))

    if last_date == today:
        major = prior_major
        minor = current.get("minor", 0) + 1
    else:
        major = prior_major + 1
        minor = 0

    new = {
        "major": major,
        "minor": minor,
        "version": _version_str(major, minor),
        "lastRefreshed": now.isoformat(timespec="seconds"),
    }
    VERSION_JSON.write_text(json.dumps(new), encoding="utf-8")
    return new
