"""Extract the full asset register (tractors + trailers + RBOX) from RDW_EOS_Master_latest.xlsx."""
import datetime
import json
from pathlib import Path

import openpyxl
import pandas as pd

from dashboard_filters import is_excluded_terminal

SRC = Path(__file__).parent / "output" / "RDW_EOS_Master_latest.xlsx"
OUT = Path(__file__).parent / "output" / "asset_data.json"

# Manual out-of-service override for Trailer Master -- ported as-is from the prior
# dashboard build (rdw-fleet-site), not derived from any live source. This is a
# point-in-time hand-curated list, not automated data -- review/update periodically
# rather than trusting it stays accurate indefinitely.
TRAILER_MASTER_OOS_IDS = {"1130", "1587", "1637", "462", "AES132", "AES137", "AES139"}


def sheet_df(wb, name):
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    return pd.DataFrame(rows[1:], columns=rows[0])


def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    dim_equipment = sheet_df(wb, "Dim_Equipment")

    econ = sheet_df(wb, "Fact_Asset_Economics")
    econ_wide = econ.pivot_table(index="AssetKey", columns="MeasureName", values="Value", aggfunc="first")

    maint = sheet_df(wb, "Fact_Maintenance_WO")
    maint_cost = maint.groupby("AssetKey")["TotalCost"].sum()

    today_year = 2026  # session "today" per environment context, avoid datetime.now() drift

    records = []
    excluded_count = 0
    for row in dim_equipment.itertuples():
        ak = row.AssetKey
        e = econ_wide.loc[ak] if ak in econ_wide.index else pd.Series(dtype=float)

        def g(key):
            v = e.get(key) if not e.empty else None
            return None if v is None or pd.isna(v) else round(float(v), 2)

        age = None
        if row.InServiceDate:
            try:
                yr = int(str(row.InServiceDate)[:4])
                age = today_year - yr
            except (ValueError, TypeError):
                age = None

        if is_excluded_terminal(row.Terminal, row.FleetCode):
            excluded_count += 1
            continue

        asset_id_norm = str(row.AssetID or "").strip().upper()
        raw_type = str(row.Model or "").strip().upper()
        trailer_equipment_type = "R/O" if raw_type == "R-O" else raw_type

        records.append({
            "assetKey": ak,
            "assetId": row.AssetID,
            "assetType": row.AssetType,
            "ownership": row.OwnershipClass,
            "division": row.Division,
            "terminal": row.Terminal,
            "fleetCode": row.FleetCode,
            "styleKey": row.StyleKey,
            "status": row.ServiceStatus,
            "year": None if pd.isna(row.Year) else int(row.Year),
            "make": row.Make,
            "model": row.Model,
            "inServiceDate": row.InServiceDate,
            "age": age,
            "rdwOwned": row.RDWOwned,
            "financed": row.Financed,
            "payoffDate": row.NotePayoffDate,
            "attributeStandard": row.AttributeStandard,
            "ambiguousId": row.AmbiguousAssetID == "TRUE" if row.AmbiguousAssetID else False,
            "netBookValue": g("Net Book Value"),
            "depreciation": g("Depreciation"),
            "maintenanceCostTotal": None if ak not in maint_cost.index else round(float(maint_cost.loc[ak]), 2),
            "trailerEquipmentType": trailer_equipment_type or None,
            "trailerMasterStatus": "Out of Service" if asset_id_norm in TRAILER_MASTER_OOS_IDS else (row.ServiceStatus or "Active"),
        })

    by_type = pd.DataFrame(records)["assetType"].value_counts().to_dict()

    trailers = [r for r in records if r["assetType"] == "TRAILER"]
    meta = {
        "builtFrom": "RDW_EOS_Master_latest.xlsx",
        "assetCount": len(records),
        "byType": by_type,
        "excludedTerminalCount": excluded_count,
        "trailerMasterRawCount": len(trailers),
        "trailerMasterOutOfServiceCount": sum(1 for r in trailers if r["trailerMasterStatus"] == "Out of Service"),
        "trailerMasterMissingTypeCount": sum(1 for r in trailers if not r["trailerEquipmentType"]),
    }
    payload = {"meta": meta, "assets": records}
    OUT.write_text(json.dumps(payload, default=str), encoding="utf-8")
    print(f"Wrote {len(records)} assets, {OUT.stat().st_size:,} bytes -> {OUT} ({excluded_count} excluded via dashboard_filters.EXCLUDED_TERMINALS)")
    print(by_type)


if __name__ == "__main__":
    main()
