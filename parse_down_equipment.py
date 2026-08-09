"""
Parse the "Down Equipment" tab of the Down Report workbook (SharePoint, pulled
by hand -- not on an automated feed) into per-tractor down/repair detail:
Down Days, Status, Location, Shop Number, Issue, ETA of Completion.

This is a richer, more direct signal for "is this truck currently down" than
the RTD Assigned Driver field's PENDING code: a truck can be down for repair
while still carrying its normal driver's code (the driver just doesn't have a
working unit at the moment), so this report catches down trucks the RTD
PENDING code alone would miss.

Filtering: a row counts as an active down record only if it has a real Status
value, that status isn't "Complete" (resolved), and the Tractor/Trailer field
isn't "Trailer" (this report also carries a few trailer rows, out of scope for
the tractor dashboard). Rows with no Status at all are roster/contact entries,
not down events, and are skipped.
"""
import json
from pathlib import Path

import openpyxl

SRC = Path(__file__).parent / "data" / "raw" / "Down_Report_latest.xlsm"
OUT = Path(__file__).parent / "output" / "down_equipment_data.json"


def main():
    wb = openpyxl.load_workbook(SRC, data_only=True, keep_vba=True)
    ws = wb["Down Equipment"]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[1]
    idx = {h: i for i, h in enumerate(header) if h}

    records = {}
    skipped_resolved = skipped_no_status = skipped_trailer = skipped_bad_data = 0

    for r in rows[2:]:
        unit = r[idx["Unit #"]]
        if unit is None:
            continue
        unit = str(unit).strip().upper()
        status = r[idx["Status"]]
        tractor_or_trailer = r[idx["Tractor/Trailer"]]

        if tractor_or_trailer == "Trailer":
            skipped_trailer += 1
            continue
        if status is None:
            skipped_no_status += 1
            continue
        if status == "Complete":
            skipped_resolved += 1
            continue
        if r[idx["Tractor#"]] == "BAD DATA":
            skipped_bad_data += 1
            continue

        records[unit] = {
            "downDays": r[idx["Down Days"]],
            "status": status,
            "location": r[idx["Location"]],
            "shopNumber": r[idx["Shop Number"]],
            "issue": r[idx["Issue"]],
            "etaCompletion": r[idx["ETA of Completion"]],
            "dateIn": r[idx["Date In"]],
        }

    print(f"=== Down Equipment: {len(records)} active down tractors ===")
    for unit, d in records.items():
        print(f"  {unit}: {d['status']} ({d['downDays']} days) @ {d['location']} -- {d['issue']}")
    print(f"\nSkipped: {skipped_resolved} resolved, {skipped_no_status} no-status (roster only), "
          f"{skipped_trailer} trailers, {skipped_bad_data} bad-data rows")

    OUT.write_text(json.dumps(records, default=str), encoding="utf-8")
    print(f"\nWrote {OUT}")


if __name__ == "__main__":
    main()
