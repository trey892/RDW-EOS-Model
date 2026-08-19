"""
Build the Source Worksheets tab's status snapshot from what the pipeline
actually knows about each source -- no live Drive polling (this is a static
Artifact build), but every timestamp here is real, pulled from the parser
outputs and local file mtimes, not fabricated or copied from the old
dashboard's sample data.

Bucket meanings (must match sourceLabels/feedNames in the template JS):
  synced  -- Shared Sheets Sync (Power Automate -> Drive)
  rta     -- RTA Reports Inbox (Gmail -> Apps Script -> Drive)
  mcleod  -- McLeod Reports Inbox (Gmail -> Apps Script -> Drive)
  samsara -- no automated ingestion exists yet; intentionally empty, not faked
  drops   -- not used yet
"""
import datetime
import json
from pathlib import Path

OUT_DIR = Path(__file__).parent / "output"
RAW_DIR = Path(__file__).parent / "data" / "raw"
OUT = OUT_DIR / "sources_data.json"


def _iso_mtime(path):
    if not path.exists():
        return None
    return datetime.datetime.fromtimestamp(path.stat().st_mtime, tz=datetime.timezone.utc).isoformat()


def _load(path, default=None):
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def main():
    files = []

    # -- Shared Sheets Sync (SharePoint Sync folder) --
    rtd_path = RAW_DIR / "RTD_latest.xlsx"
    files.append({
        "bucket": "synced", "name": "RTD Tractor Listing", "kind": "Dispatch status / hub mileage",
        "url": "https://drive.google.com/drive/folders/1uD98KZD7BaTsqCfclRKtBI7lwTgoH9a3",
        "modifiedTime": _iso_mtime(rtd_path), "dashboardFeed": True, "dashboardCurrent": rtd_path.exists(),
    })
    orientation_path = RAW_DIR / "Orientation_latest.xlsm"
    orientation = _load(OUT_DIR / "orientation_data.json")
    files.append({
        "bucket": "synced", "name": "Orientation.xlsm", "kind": "Driver orientation roster",
        "url": "https://drive.google.com/drive/folders/1uD98KZD7BaTsqCfclRKtBI7lwTgoH9a3",
        "modifiedTime": _iso_mtime(orientation_path),
        "reportTimestamp": orientation.get("checkedAt") if orientation else None,
        "dashboardFeed": True, "dashboardCurrent": orientation_path.exists(),
    })

    # -- RTA Reports Inbox --
    pm_due_path = RAW_DIR / "Asset_and_Equipment_PMs_Due_latest.csv"
    files.append({
        "bucket": "rta", "name": "Asset and Equipment PM Due", "kind": "PM due list, all facilities",
        "url": "https://drive.google.com/drive/folders/1YlR0efDJstVmmpInbyG6KCIuxO-I1IDF",
        "modifiedTime": _iso_mtime(pm_due_path),
        "reportTimestamp": None,  # RTA's own report date -- see SKILL.md staleness check; not carried into this JSON yet
        "dashboardFeed": True, "dashboardCurrent": pm_due_path.exists(),
    })
    pm_compliance_path = RAW_DIR / "PM_Compliance_latest.csv"
    pm_compliance = _load(OUT_DIR / "pm_compliance_data.json")
    files.append({
        "bucket": "rta", "name": "PM Compliance", "kind": "On-time/early/late by vehicle",
        "url": "https://drive.google.com/drive/folders/1YlR0efDJstVmmpInbyG6KCIuxO-I1IDF",
        "modifiedTime": _iso_mtime(pm_compliance_path),
        "reportTimestamp": pm_compliance.get("reportDate") if pm_compliance else None,
        "dashboardFeed": True, "dashboardCurrent": pm_compliance_path.exists(),
    })

    # -- McLeod Reports Inbox --
    unbilled_path = RAW_DIR / "Unbilled_Orders_latest.txt"
    unbilled = _load(OUT_DIR / "unbilled_orders_data.json")
    files.append({
        "bucket": "mcleod", "name": "Unbilled Orders Report", "kind": "Open orders by terminal",
        "url": "https://drive.google.com/drive/folders/1oGIyA_0U6XLYjmv3mVadsH3xIYbkVh5k",
        "modifiedTime": _iso_mtime(unbilled_path),
        "reportTimestamp": unbilled.get("checkedAt") if unbilled else None,
        "dashboardFeed": True, "dashboardCurrent": unbilled_path.exists(),
    })

    # -- Samsara: intentionally no rows. No automated parser exists yet (see
    # Metric Source Map note in the template) -- an empty bucket here is honest,
    # not a bug to "fix" by adding a fake row.

    checked_at = datetime.datetime.now(tz=datetime.timezone.utc).isoformat()
    payload = {"checkedAt": checked_at, "files": files}
    OUT.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")
    print(f"Wrote {len(files)} source file entries -> {OUT}")
    for f in files:
        print(f"  [{f['bucket']}] {f['name']}: modified={f['modifiedTime']}, current={f['dashboardCurrent']}")


if __name__ == "__main__":
    main()
