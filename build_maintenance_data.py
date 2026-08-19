"""Extract work-order history from RDW_EOS_Master_latest.xlsx for the Maintenance view."""
import datetime
import json
from pathlib import Path

import openpyxl
import pandas as pd

SRC = Path(__file__).parent / "output" / "RDW_EOS_Master_latest.xlsx"
OUT = Path(__file__).parent / "output" / "maintenance_data.json"
DOWN_EQUIPMENT_JSON = Path(__file__).parent / "output" / "down_equipment_data.json"
PM_COMPLIANCE_JSON = Path(__file__).parent / "output" / "pm_compliance_data.json"


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

    dim_equipment = sheet_df(wb, "Dim_Equipment")
    terminal_by_asset = {
        str(r.AssetID).strip().upper(): r.Terminal
        for r in dim_equipment.itertuples() if r.AssetType == "TRACTOR" and r.AssetID
    }

    econ = sheet_df(wb, "Fact_Asset_Economics")
    econ_wide = econ.pivot_table(index="AssetKey", columns="MeasureName", values="Value", aggfunc="first")

    # -- Highest-cost tractors --
    cost_by_asset = wo.groupby("AssetID")["TotalCost"].sum().dropna()
    highest_cost_tractors = [
        {"assetId": aid, "terminal": terminal_by_asset.get(str(aid).strip().upper()), "totalCost": round(float(c), 2)}
        for aid, c in cost_by_asset.sort_values(ascending=False).head(20).items()
    ]

    # -- Units omitted from Maintenance CPM (has cost, but no Annualized Distance to divide by) --
    omitted_from_cpm = []
    for aid, cost in cost_by_asset.items():
        ak = f"TRC-{aid}"
        dist = econ_wide.loc[ak].get("Annualized Distance") if ak in econ_wide.index else None
        if dist is None or pd.isna(dist) or float(dist) <= 0:
            omitted_from_cpm.append({"assetId": aid, "terminal": terminal_by_asset.get(str(aid).strip().upper()), "totalCost": round(float(cost), 2)})

    # -- Down units & downtime, by terminal --
    down_equipment = {}
    if DOWN_EQUIPMENT_JSON.exists():
        down_equipment = json.loads(DOWN_EQUIPMENT_JSON.read_text(encoding="utf-8"))
    down_by_terminal = {}
    for aid, d in down_equipment.items():
        t = terminal_by_asset.get(str(aid).strip().upper(), "UNKNOWN")
        agg = down_by_terminal.setdefault(t, {"terminal": t, "downCount": 0, "totalDownDays": 0})
        agg["downCount"] += 1
        agg["totalDownDays"] += d.get("downDays") or 0
    down_units_by_terminal = sorted(down_by_terminal.values(), key=lambda x: -x["downCount"])

    # -- PM Compliance % by terminal (joins PM Compliance vehicle IDs to terminal via Dim_Equipment) --
    pm_compliance_vehicles = []
    if PM_COMPLIANCE_JSON.exists():
        pm_compliance_vehicles = json.loads(PM_COMPLIANCE_JSON.read_text(encoding="utf-8")).get("vehicles", [])
    compliance_by_terminal = {}
    for v in pm_compliance_vehicles:
        t = terminal_by_asset.get(str(v["vehicle"]).strip().upper())
        if not t or v.get("onTimePct") is None:
            continue
        agg = compliance_by_terminal.setdefault(t, {"terminal": t, "vehicleCount": 0, "onTimeSum": 0.0})
        agg["vehicleCount"] += 1
        agg["onTimeSum"] += v["onTimePct"]
    pm_compliance_by_terminal = [
        {"terminal": t, "vehicleCount": a["vehicleCount"], "avgOnTimePct": round(a["onTimeSum"] / a["vehicleCount"], 1)}
        for t, a in compliance_by_terminal.items()
    ]

    # -- PM Due & Compliance by Terminal (combined) --
    pm_due_by_terminal = {}
    for r in pm_records:
        t = r["terminal"] or "UNKNOWN"
        agg = pm_due_by_terminal.setdefault(t, {"terminal": t, "pmDueCount": 0, "pastDueCount": 0})
        agg["pmDueCount"] += 1
        if r["dueStatus"] == "Past Due":
            agg["pastDueCount"] += 1
    compliance_lookup = {c["terminal"]: c["avgOnTimePct"] for c in pm_compliance_by_terminal}
    pm_due_compliance_by_terminal = [
        {**agg, "avgOnTimePct": compliance_lookup.get(t)}
        for t, agg in pm_due_by_terminal.items()
    ]

    # -- PM code filter options --
    pm_codes = sorted({r["pmCode"] for r in pm_records if r["pmCode"]})

    # Road Calls & Repeat Shop Visits: no data source exists for this anywhere in the
    # pipeline (the prior dashboard build hardcoded "N/A" for the same reason) --
    # explicitly marked unavailable rather than fabricated.
    road_calls_available = False

    meta = {
        "builtFrom": "RDW_EOS_Master_latest.xlsx / Fact_Maintenance_WO",
        "workOrderCount": len(records),
        "tractorsWithWO": wo["AssetKey"].nunique(),
        "dateRange": [str(wo["OpenedDate"].min()), str(wo["OpenedDate"].max())],
        "pmDueCount": len(pm_records),
        "pmDuePastDueCount": n_past_due,
        "pmDueAssetCount": pm_due_tractors["AssetID"].nunique(),
        "pmDueSource": (
            "RTA Asset and Equipment PM Due export, all facilities -- automated: "
            f"RTA schedule -> email -> Gmail filter -> Apps Script -> Drive (last refreshed {datetime.date.today().isoformat()})"
        ),
        "pmComplianceVehicleCount": len(pm_compliance_vehicles),
        "downUnitCount": len(down_equipment),
        "roadCallsAvailable": road_calls_available,
        "pmCodes": pm_codes,
    }
    payload = {
        "meta": meta,
        "workOrders": records,
        "pmDue": pm_records,
        "highestCostTractors": highest_cost_tractors,
        "omittedFromCpm": omitted_from_cpm,
        "downUnitsByTerminal": down_units_by_terminal,
        "pmComplianceByTerminal": pm_compliance_by_terminal,
        "pmDueComplianceByTerminal": pm_due_compliance_by_terminal,
    }
    OUT.write_text(json.dumps(payload, default=str), encoding="utf-8")
    print(f"Wrote {len(records)} work orders + {len(pm_records)} PM-due records, {OUT.stat().st_size:,} bytes -> {OUT}")
    print(meta)
    print(f"Highest-cost tractors: {len(highest_cost_tractors)}, omitted from CPM: {len(omitted_from_cpm)}")
    print(f"Down units by terminal: {down_units_by_terminal}")
    print(f"PM Due + Compliance by terminal: {pm_due_compliance_by_terminal}")


if __name__ == "__main__":
    main()
