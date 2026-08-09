"""
Aggregate the invoiced fuel-card transaction history (Fuel_History_Cleaned_and_Summarized.xlsx,
"Transactions" tab -- 36,101 rows, one per fuel-card swipe) into per-tractor totals.

Two separate figures come out of this, deliberately kept apart:
  - fuelCostNet: Net Total (post-discount, all fuel/reefer/misc charges) -- the RDW-cost basis,
    matches the "12 months invoiced net of discounts" language already used for Fact_Asset_Economics'
    Fuel Cost measure.
  - propulsionGallons: Diesel Gal + Other Fuel Gal only -- excludes Reefer Gal, since reefer-unit
    diesel doesn't power the tractor and would distort an MPG (propulsion efficiency) calculation.
"""
import json
from collections import defaultdict
from pathlib import Path

import openpyxl

SRC = Path(__file__).parent / "data" / "raw" / "Fuel_History_Cleaned_and_Summarized.xlsx"
OUT = Path(__file__).parent / "output" / "fuel_history_data.json"


def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb["Transactions"]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    idx = {h: i for i, h in enumerate(header)}

    agg = defaultdict(lambda: {
        "dieselGallons": 0.0, "otherFuelGallons": 0.0, "reeferGallons": 0.0,
        "fuelCostNet": 0.0, "transactionCount": 0,
    })
    total_diesel_gal = total_diesel_amt = total_net = 0.0
    unmatched_unit_rows = 0

    for r in rows[1:]:
        unit = r[idx["Unit #"]]
        total_diesel_gal += r[idx["Diesel Gal"]] or 0
        total_diesel_amt += r[idx["Diesel Amt"]] or 0
        total_net += r[idx["Net Total"]] or 0
        if unit is None or str(unit).strip() == "":
            unmatched_unit_rows += 1
            continue
        unit = str(unit).strip().upper()
        a = agg[unit]
        a["dieselGallons"] += r[idx["Diesel Gal"]] or 0
        a["otherFuelGallons"] += r[idx["Other Fuel Gal"]] or 0
        a["reeferGallons"] += r[idx["Reefer Gal"]] or 0
        a["fuelCostNet"] += r[idx["Net Total"]] or 0
        a["transactionCount"] += 1

    for unit, a in agg.items():
        a["propulsionGallons"] = round(a["dieselGallons"] + a["otherFuelGallons"], 2)
        a["dieselGallons"] = round(a["dieselGallons"], 2)
        a["otherFuelGallons"] = round(a["otherFuelGallons"], 2)
        a["reeferGallons"] = round(a["reeferGallons"], 2)
        a["fuelCostNet"] = round(a["fuelCostNet"], 2)

    print("=== RECONCILIATION vs file's own Summary tab ===")
    print(f"Total Diesel Gallons: {total_diesel_gal:,.2f}  (Summary tab says 3,605,630.48)")
    print(f"Total Diesel Amount:  ${total_diesel_amt:,.2f}  (Summary tab says $14,996,405.50)")
    print(f"Grand Net Total:      ${total_net:,.2f}  (Summary tab says $14,495,627.93)")
    print(f"Rows with no Unit #:  {unmatched_unit_rows}")
    print(f"Distinct units:       {len(agg)}")

    OUT.write_text(json.dumps(agg, default=str), encoding="utf-8")
    print(f"\nWrote {len(agg)} units -> {OUT}")


if __name__ == "__main__":
    main()
