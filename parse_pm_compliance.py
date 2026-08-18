"""
Parse RTA's banded "PM Compliance" CSV export into a per-vehicle on-time/
early/late summary. Same general banded-report shape as
Asset_and_Equipment_PMs_Due (see parse_pm_due.py) but grouped by vehicle
with a "<vehicle> Summary" line followed by a "Total PMs" stats row, rather
than by asset with inline due-status rows.

Report covers a trailing date range embedded in its own header (e.g. "Date
Range: 07/01/2026 to 07/31/2026") -- that's RTA's reporting period, not
today's date; carry it through to the output so callers know what window
the compliance % actually reflects.
"""
import csv
import json
import re
from pathlib import Path

SRC = Path(__file__).parent / "data" / "raw" / "PM_Compliance_latest.csv"
OUT = Path(__file__).parent / "output" / "pm_compliance_data.json"


def _pct(s):
    if not s:
        return None
    try:
        return float(s.strip().rstrip("%"))
    except ValueError:
        return None


def parse():
    with open(SRC, encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))

    date_range = None
    report_date = None
    facility = customer = vehicle = None
    detail_rows = []
    vehicles = []

    def col(row, i):
        return row[i].strip() if len(row) > i and row[i] else ""

    for row in rows:
        c0 = col(row, 0)

        if c0.startswith("Date Range:"):
            date_range = c0.replace("Date Range:", "").strip()
            continue
        if col(row, 8) == "Date:":
            report_date = col(row, 9)
            continue
        if c0.startswith("Facility:") and len(c0) > len("Facility:"):
            facility = c0.replace("Facility:", "").strip()
            continue
        if c0.startswith("Customer:"):
            customer = c0.replace("Customer:", "").strip()
            continue
        if c0 == "Vehicle":  # column header row
            continue
        if c0 == "Total PMs":
            total = col(row, 1)
            on_time = col(row, 3)
            early = col(row, 5)
            late = col(row, 7)
            if vehicle:
                vehicles.append({
                    "vehicle": vehicle,
                    "facility": facility,
                    "customer": customer,
                    "totalPMs": int(total) if total.isdigit() else None,
                    "onTimePct": _pct(on_time),
                    "earlyPct": _pct(early),
                    "latePct": _pct(late),
                    "records": detail_rows,
                })
            vehicle = None
            detail_rows = []
            continue
        if c0.endswith(" Summary"):
            continue  # label-only row, stats are on the following "Total PMs" row

        # PM detail row: blank col0, PM code in col1
        if c0 == "" and col(row, 1) and not col(row, 1).startswith("("):
            detail_rows.append({
                "pmCode": col(row, 1),
                "primaryMeter": col(row, 2),
                "pmDue": col(row, 3),
                "pmCompleted": col(row, 4),
                "meterType": col(row, 5),
                "tolerance": col(row, 6),
                "notes": col(row, 7),
            })
            continue

        # New vehicle-id row: non-blank col0, every other col blank
        if c0 and not any(col(row, i) for i in range(1, len(row))):
            vehicle = c0
            detail_rows = []
            continue

    return {"dateRange": date_range, "reportDate": report_date, "vehicles": vehicles}


def main():
    result = parse()
    vehicles = result["vehicles"]
    print(f"=== PM Compliance ({result['dateRange']}, report dated {result['reportDate']}) ===")
    print(f"Parsed {len(vehicles)} vehicle(s)")
    if vehicles:
        avg_on_time = sum(v["onTimePct"] or 0 for v in vehicles) / len(vehicles)
        print(f"Average On Time %: {avg_on_time:.1f}%")
        print("Sample:", vehicles[0])

    OUT.write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(f"\nWrote -> {OUT}")


if __name__ == "__main__":
    main()
