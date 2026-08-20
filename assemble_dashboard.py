"""
Assemble the final dashboard HTML by embedding the JSON datasets into
dashboard_template.html in place of its placeholder markers. Run after
build_dashboard_data.py / build_asset_data.py / build_maintenance_data.py
(and extract_lease_data.py, if lease_data.json needs refreshing).

The "optional source" datasets (orientation/pm_compliance/unbilled_orders/service_incidents/revenue_analysis)
fall back to an empty-shaped default if their JSON doesn't exist yet (e.g. very
first run, or that day's Drive pull found nothing) -- the dashboard should
render fine with those panels showing "no data" rather than the whole assemble
step failing.

Output is CSP-safe for Artifact publishing: all data is embedded, no fetch calls.

The real, git-tracked template source is templates/dashboard_template.html --
output/ is entirely gitignored (real data + build artifacts only). This script
always copies the tracked source into output/ itself before assembling, so
there's no separate manual copy step to forget and no way for a stale/edited
copy to silently drift out of sync with the tracked source (this happened
once, 2026-08-19 -- an editing session edited output/dashboard_template.html
directly, which was never committed and would have been lost on the next
copy-over). Always edit templates/dashboard_template.html, never output/'s copy.
"""
import json
import shutil
from pathlib import Path

ROOT = Path(__file__).parent
OUT_DIR = ROOT / "output"
TEMPLATE_SRC = ROOT / "templates" / "dashboard_template.html"
TEMPLATE = OUT_DIR / "dashboard_template.html"
OUT = OUT_DIR / "RDW_Fleet_Dashboard.html"

REQUIRED_PLACEHOLDERS = {
    "/*__DATA__*/": OUT_DIR / "dashboard_data.json",
    "/*__ASSET_DATA__*/": OUT_DIR / "asset_data.json",
    "/*__MAINT_DATA__*/": OUT_DIR / "maintenance_data.json",
    "/*__LEASE_DATA__*/": OUT_DIR / "lease_data.json",
}

OPTIONAL_PLACEHOLDERS = {
    "/*__ORIENTATION_DATA__*/": (OUT_DIR / "orientation_data.json", {"classDate": None, "classDateLabel": None, "driverCount": 0, "drivers": [], "fleets": []}),
    "/*__PM_COMPLIANCE_DATA__*/": (OUT_DIR / "pm_compliance_data.json", {"dateRange": None, "reportDate": None, "vehicles": []}),
    "/*__UNBILLED_ORDER_DATA__*/": (OUT_DIR / "unbilled_orders_data.json", {"checkedAt": None, "orderCount": 0, "totalCharges": 0, "rows": []}),
    "/*__SERVICE_INCIDENTS_DATA__*/": (OUT_DIR / "service_incidents_data.json", {"checkedAt": None, "dateRangeLabel": None, "reportTotalIncidents": None, "parsedIncidents": 0, "unparsedIncidents": 0, "unmappedIncidents": 0, "rows": []}),
    "/*__SOURCE_DATA__*/": (OUT_DIR / "sources_data.json", {"checkedAt": None, "files": []}),
    "/*__PACCAR_LESSEE_DATA__*/": (OUT_DIR / "paccar_lessee_data.json", {"inspectionSource": None, "reportDate": None, "rtdReportDate": None, "portalUrl": None, "units": []}),
    "/*__REVENUE_ANALYSIS_DATA__*/": (OUT_DIR / "revenue_analysis_data.json", {"checkedAt": None, "periodStart": None, "periodEnd": None, "totals": None, "byOwnership": [], "byTerminal": [], "idleIron": {"thresholdDays": 5, "count": 0, "ghostCount": 0, "units": []}, "dataIntegrity": {}, "droppedSections": {}}),
    "/*__REVENUE_NARRATIVE_DATA__*/": (OUT_DIR / "revenue_narrative.json", {"writtenAt": None, "html": None}),
    "/*__SAMSARA_DRIVER_DATA__*/": (OUT_DIR / "samsara_driver_data.json", {"checkedAt": None, "matchedDriverCount": 0, "fuelReport": {"matched": 0, "unmatchedCount": 0}, "safetyReport": {"matched": 0, "unmatchedCount": 0}, "drivers": []}),
}


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(TEMPLATE_SRC, TEMPLATE)

    html = TEMPLATE.read_text(encoding="utf-8")

    for marker, data_file in REQUIRED_PLACEHOLDERS.items():
        if marker not in html:
            raise ValueError(f"Marker {marker} not found in {TEMPLATE}")
        payload = data_file.read_text(encoding="utf-8")
        html = html.replace(marker, payload, 1)

    for marker, (data_file, default) in OPTIONAL_PLACEHOLDERS.items():
        if marker not in html:
            raise ValueError(f"Marker {marker} not found in {TEMPLATE}")
        payload = data_file.read_text(encoding="utf-8") if data_file.exists() else json.dumps(default)
        html = html.replace(marker, payload, 1)

    OUT.write_text(html, encoding="utf-8")
    print(f"Wrote {OUT} ({OUT.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
