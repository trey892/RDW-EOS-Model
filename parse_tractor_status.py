"""
Derive each tractor's dispatch status, style code, and current hub mileage from
the RTD "TRACTOR LISTING" export's Assigned Driver field ONLY -- per Trey's
direction, the sheet's own coarse "Status" column (blank/In Shop/Open/Yard
Truck) is not authoritative and is ignored here; it doesn't reliably agree
with Assigned Driver (some tractors marked "Open" already have a real driver
code, meaning Status wasn't updated at time of movement per the Tractor
Allocation SOP).

Per that SOP, the Assigned Driver field must never be blank -- when a tractor
has no actual driver, dispatch puts the status code itself into that field.
The 9 official codes: SPARE, ASSIGNED, AVAILABL, PENDING, PREP, ORIENTAT,
TOBELP, TOSELL, WRECK. "Open" = AVAILABL (cleaned, inspected, ready for
assignment).

As of 2026-08-19: this reads the same daily-synced RTD workbook (canonical
name RTD_latest.xlsx, refreshed each run from the "SharePoint Sync" Drive
folder -- same source as the PM/mileage data) instead of a separately
hand-supplied export, so this is now part of the automated daily refresh.
`style` is the sheet's raw "Type" column code (CS, MP, SN, etc.) -- these are
RTA-internal type codes, not yet mapped to friendly labels (Chemical/Waste/
Straight Truck style names seen elsewhere); shown as-is until a real mapping
is confirmed rather than guessed.

`rawSheetStatus` is the sheet's own "Status" column (blank / "Open" / "Yard
Truck" / "In Shop"), captured alongside dispatchStatus per Trey's 2026-08
request to show both side by side on the dashboard -- it is NOT used to
derive dispatchStatus or any other computed field, since it's the same
column called out above as unreliable against Assigned Driver.
"""
import json
from pathlib import Path

import openpyxl

SRC = Path(__file__).parent / "data" / "raw" / "RTD_latest.xlsx"
OUT = Path(__file__).parent / "output" / "tractor_status_data.json"

SOP_CODES = {"SPARE", "ASSIGNED", "AVAILABL", "PENDING", "PREP", "ORIENTAT", "TOBELP", "TOSELL", "WRECK"}


def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb["TRACTOR LISTING"]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    idx = {h: i for i, h in enumerate(header)}

    records = {}
    for r in rows[1:]:
        tractor = r[idx["Tractor Number"]]
        if not tractor:
            continue
        tractor = str(tractor).strip().upper()
        assigned_driver_raw = (r[idx["Assigned Driver"]] or "").strip()
        code = assigned_driver_raw.upper()

        if code == "":
            dispatch_status = "BLANK"  # SOP violation -- field must never be blank
        elif code in SOP_CODES:
            dispatch_status = code
        else:
            dispatch_status = "ASSIGNED"  # a real driver code is in the field

        raw_sheet_status = (r[idx["Status"]] or "").strip()

        records[tractor] = {
            "assignedDriverRaw": assigned_driver_raw,
            "dispatchStatus": dispatch_status,
            "rawSheetStatus": raw_sheet_status or None,  # sheet's own Status column, shown for reference only -- see module docstring
            "dateAssigned": r[idx["Date Assigned"]],
            "dispatcher": r[idx["Dispatcher"]],
            "fleet": r[idx["Fleet"]],
            "style": (r[idx["Type"]] or "").strip(),
            "currentHub": r[idx["Current Hub"]],
        }

    from collections import Counter
    counts = Counter(v["dispatchStatus"] for v in records.values())
    print("=== Dispatch status (Assigned Driver field only) ===")
    for status, n in counts.most_common():
        print(f"  {status}: {n}")
    print(f"\nWrote {len(records)} tractors -> {OUT}")

    OUT.write_text(json.dumps(records, default=str), encoding="utf-8")


if __name__ == "__main__":
    main()
