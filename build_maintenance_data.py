"""Extract work-order history from RDW_EOS_Master_v8.xlsx for the Maintenance view."""
import json
from pathlib import Path

import openpyxl
import pandas as pd

SRC = Path(__file__).parent / "output" / "RDW_EOS_Master_v8.xlsx"
OUT = Path(__file__).parent / "output" / "maintenance_data.json"


def sheet_df(wb, name):
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    return pd.DataFrame(rows[1:], columns=rows[0])


def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    wo = sheet_df(wb, "Fact_Maintenance_WO")
    wo = wo[wo["AssetType"] == "TRACTOR"].copy()

    records = []
    for row in wo.itertuples():
        records.append({
            "workOrderKey": row.WorkOrderKey,
            "assetKey": row.AssetKey,
            "assetId": row.AssetID,
            "openedDate": row.OpenedDate,
            "closedDate": row.ClosedDate,
            "parts": None if pd.isna(row.Parts) else round(float(row.Parts), 2),
            "tires": None if pd.isna(row.Tires) else round(float(row.Tires), 2),
            "labor": None if pd.isna(row.Labor) else round(float(row.Labor), 2),
            "outside": None if pd.isna(row.Outside) else round(float(row.Outside), 2),
            "misc": None if pd.isna(row.Misc) else round(float(row.Misc), 2),
            "totalCost": None if pd.isna(row.TotalCost) else round(float(row.TotalCost), 2),
            "isOpen": row.IsOpen,
            "dateIntegrityFlag": row.DateIntegrityFlag,
        })

    pm_due = sheet_df(wb, "Fact_PM_Due")
    pm_due_tractors = pm_due[pm_due["AssetKey"].astype(str).str.startswith("TRC-", na=False)].copy()
    pm_records = []
    for row in pm_due_tractors.itertuples():
        pm_records.append({
            "assetKey": row.AssetKey,
            "assetId": row.AssetID,
            "ownership": row.OwnershipClass,
            "division": row.Division,
            "terminal": row.Terminal,
            "pmCode": row.PMCode,
            "pmDescription": row.PMDescription,
            "cycleType": row.CycleType,
            "interval": None if pd.isna(row.Interval) else int(row.Interval),
            "previousDone": row.PreviousDone,
            "dueAt": row.DueAt,
            "dueIn": None if pd.isna(row.DueIn) else int(row.DueIn),
            "dueUnit": row.DueUnit,
            "dueStatus": row.DueStatus,
            "scheduled": row.Scheduled,
        })

    n_past_due = sum(1 for r in pm_records if r["dueStatus"] == "Past Due")

    meta = {
        "builtFrom": "RDW_EOS_Master_v8.xlsx / Fact_Maintenance_WO",
        "workOrderCount": len(records),
        "tractorsWithWO": wo["AssetKey"].nunique(),
        "dateRange": [str(wo["OpenedDate"].min()), str(wo["OpenedDate"].max())],
        "pmDueCount": len(pm_records),
        "pmDuePastDueCount": n_past_due,
        "pmDueAssetCount": pm_due_tractors["AssetID"].nunique(),
        "pmDueSource": "RTA Asset and Equipment PM Due export, all facilities -- automated: RTA schedule -> email -> Gmail filter -> Apps Script -> Drive (first automated run 2026-08-08)",
    }
    payload = {"meta": meta, "workOrders": records, "pmDue": pm_records}
    OUT.write_text(json.dumps(payload, default=str), encoding="utf-8")
    print(f"Wrote {len(records)} work orders + {len(pm_records)} PM-due records, {OUT.stat().st_size:,} bytes -> {OUT}")
    print(meta)


if __name__ == "__main__":
    main()
