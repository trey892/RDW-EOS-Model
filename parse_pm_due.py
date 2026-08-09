"""Parse RTA's banded 'Asset and Equipment PM Due' CSV export into a flat table."""
import csv
import json
import re
from pathlib import Path

SRC = Path(__file__).parent / "data" / "raw" / "Asset_and_Equipment_PMs_Due_latest.csv"
OUT = Path(__file__).parent / "output" / "pm_due_data.json"

DUE_STATUSES = {"Past Due", "Due Soon", "Due Now", "Not Due"}
SCHEDULED_VALS = {"Scheduled", "Unscheduled"}


def parse():
    with open(SRC, encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))

    records = []
    facility = None
    group = None
    asset_id = asset_desc = license_ = operator = location = None

    i = 0
    n = len(rows)
    while i < n:
        row = rows[i]
        c0 = row[0].strip() if row and row[0] else ""

        if c0.startswith("Facility:"):
            facility = c0.replace("Facility:", "").strip()
            i += 1
            continue

        m = re.match(r"^\((\S+)\)\s+(.+)$", c0)
        if m and (len(row) < 4 or not row[3]):
            group = c0
            i += 1
            continue

        if c0.startswith("Asset:"):
            asset_id = c0.replace("Asset:", "").strip()
            asset_desc = row[2].replace("Description:", "").strip() if len(row) > 2 else ""
            license_ = row[5].replace("License:", "").strip() if len(row) > 5 else ""
            operator = row[7].replace("Operator:", "").strip() if len(row) > 7 else ""
            location = row[9].replace("Location:", "").strip() if len(row) > 9 else ""
            i += 1
            continue

        # PM row: blank col0, PM code in col1, due status in col9
        if (len(row) > 10 and c0 == "" and row[1].strip()
                and row[9].strip() in DUE_STATUSES and row[10].strip() in SCHEDULED_VALS):
            pm_code = row[1].strip()
            pm_desc = row[2].strip()
            due_status = row[9].strip()
            scheduled = row[10].strip()

            cycle_type = interval = previous_done = current = due_at = due_in = None
            if i + 1 < n:
                detail = rows[i + 1]
                if len(detail) > 8 and detail[0] == "" and detail[1] == "" and detail[2] == "":
                    cycle_type = detail[3].strip() or None
                    interval = detail[4].strip() or None
                    previous_done = detail[5].strip() or None
                    current = detail[6].strip() or None
                    due_at = detail[7].strip() or None
                    due_in = detail[8].strip() or None
                    i += 1  # consume detail row

            records.append({
                "facility": facility, "group": group,
                "assetId": asset_id, "assetDescription": asset_desc,
                "license": license_, "operator": operator, "location": location,
                "pmCode": pm_code, "pmDescription": pm_desc,
                "cycleType": cycle_type, "interval": interval,
                "previousDone": previous_done, "current": current,
                "dueAt": due_at, "dueIn": due_in,
                "dueStatus": due_status, "scheduled": scheduled,
            })
            i += 1
            continue

        i += 1

    return records


def main():
    records = parse()
    print(f"Parsed {len(records)} PM-due records")

    from collections import Counter
    print("Due status:", Counter(r["dueStatus"] for r in records))
    print("Distinct assets:", len({r["assetId"] for r in records}))
    print("Sample:", records[0])
    print("Sample past-due:", next(r for r in records if r["dueStatus"] == "Past Due"))

    OUT.write_text(json.dumps(records), encoding="utf-8")
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
