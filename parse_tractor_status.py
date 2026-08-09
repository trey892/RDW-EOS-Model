"""
Derive each tractor's dispatch status from the RTD "TRACTOR LISTING" export's
Assigned Driver field ONLY -- per Trey's direction, the sheet's own coarse
"Status" column (blank/In Shop/Open/Yard Truck) is not authoritative and is
ignored here; it doesn't reliably agree with Assigned Driver (some tractors
marked "Open" already have a real driver code, meaning Status wasn't updated
at time of movement per the Tractor Allocation SOP).

Per that SOP, the Assigned Driver field must never be blank -- when a tractor
has no actual driver, dispatch puts the status code itself into that field.
The 9 official codes: SPARE, ASSIGNED, AVAILABL, PENDING, PREP, ORIENTAT,
TOBELP, TOSELL, WRECK. "Open" = AVAILABL (cleaned, inspected, ready for
assignment).

Not part of the daily automated refresh -- dispatch pulls this export by hand,
re-run this script when Trey supplies a newer one.
"""
import json
from pathlib import Path

import openpyxl

SRC = Path(__file__).parent / "data" / "raw" / "RTD_Tractor_Listing_latest.xlsx"
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

        records[tractor] = {
            "assignedDriverRaw": assigned_driver_raw,
            "dispatchStatus": dispatch_status,
            "dateAssigned": r[idx["Date Assigned"]],
            "dispatcher": r[idx["Dispatcher"]],
            "fleet": r[idx["Fleet"]],
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
