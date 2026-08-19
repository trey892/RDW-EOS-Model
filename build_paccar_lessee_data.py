"""
Merge the static PACCAR Lessee inspection/repair snapshot (data/raw/paccar_lessee_snapshot.json
-- real data, extracted from the prior fleet-site dashboard's own hardcoded fallback, see that
file's _note) with this dashboard's own live tractor identity fields (terminal, dispatch status,
assigned driver, current hub mileage), keyed by assetId.

Inspection/repair fields (operationalStatus, stage, dueLabel, repairs*, inspectionComplete, etc.)
have no live source in this pipeline yet and stay exactly as captured in the snapshot. Identity
fields are overlaid from output/dashboard_data.json each run, since that IS refreshed daily --
so a unit's current terminal/driver/mileage stays accurate even though its PACCAR status doesn't.

Units in the snapshot no longer present in the current tractor roster (already fully processed
off PACCAR's books) are dropped, not fabricated forward -- flagged in the printed summary.
"""
import json
from datetime import datetime
from pathlib import Path

SNAPSHOT_JSON = Path(__file__).parent / "data" / "raw" / "paccar_lessee_snapshot.json"
DASHBOARD_JSON = Path(__file__).parent / "output" / "dashboard_data.json"
OUT = Path(__file__).parent / "output" / "paccar_lessee_data.json"


def main():
    snapshot = json.loads(SNAPSHOT_JSON.read_text(encoding="utf-8"))
    dashboard = json.loads(DASHBOARD_JSON.read_text(encoding="utf-8"))
    tractors_by_id = {str(t["assetId"]).strip().upper(): t for t in dashboard["tractors"]}

    units = []
    missing = []
    for unit in snapshot["units"]:
        asset_id = str(unit["assetId"]).strip().upper()
        tractor = tractors_by_id.get(asset_id)
        if tractor is None:
            missing.append(asset_id)
            continue
        merged = dict(unit)
        merged["fleetCode"] = tractor.get("terminal")
        merged["rtdStatus"] = tractor.get("dispatchStatus")
        merged["assignedDriver"] = tractor.get("assignedDriverRaw")
        merged["currentHub"] = tractor.get("currentHubLive") or tractor.get("hubMileage")
        units.append(merged)

    raw_refreshed = dashboard["meta"].get("lastRefreshed")
    try:
        _d = datetime.strptime(raw_refreshed, "%Y-%m-%d") if raw_refreshed else None
        rtd_report_date = f"{_d.strftime('%b')} {_d.day}, {_d.year}" if _d else None  # avoid %-d, not portable on Windows
    except ValueError:
        rtd_report_date = raw_refreshed  # unexpected format -- show as-is rather than hide it

    out = {
        "inspectionSource": snapshot["inspectionSource"],
        "inspectionSourceUrl": snapshot["inspectionSourceUrl"],
        "reportDate": snapshot["reportDate"],
        "rtdReportDate": rtd_report_date,
        "portalUrl": snapshot["portalUrl"],
        "units": units,
    }

    print(f"=== PACCAR Lessee: {len(units)} units merged with live identity ===")
    for u in units:
        print(f"  {u['assetId']}: {u['stage']} @ {u.get('fleetCode') or 'terminal unknown'} -- {u['dueLabel']}")
    if missing:
        print(f"\nDropped (not in current tractor roster, already off PACCAR's books): {missing}")

    OUT.write_text(json.dumps(out, default=str), encoding="utf-8")
    print(f"\nWrote {OUT}")


if __name__ == "__main__":
    main()
