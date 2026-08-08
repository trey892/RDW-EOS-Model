"""
Build RDW_EOS_Master_v8.xlsx from RDW_EOS_Master_v7.xlsx (base) plus:
  Fact_PM_Due REFRESHED from the first fully-automated pull: RTA scheduled report ->
  email -> Gmail label filter -> Apps Script -> Drive ("RTA Reports Inbox" folder).
  Unlike other tables, Fact_PM_Due is a full replace on each refresh (not additive),
  matching the "Replace CSV and refresh" method already in Data_Dictionary.
"""
import json
from pathlib import Path

import openpyxl
import pandas as pd

BASE = Path(__file__).parent / "output" / "RDW_EOS_Master_v7.xlsx"
PM_DUE_JSON = Path(__file__).parent / "output" / "pm_due_data.json"
OUT = Path(__file__).parent / "output" / "RDW_EOS_Master_v8.xlsx"


def sheet_df(wb, name):
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    return pd.DataFrame(rows[1:], columns=rows[0])


def parse_int(v):
    if not v:
        return None
    try:
        return int(str(v).replace(",", ""))
    except ValueError:
        return None


def main():
    base = openpyxl.load_workbook(BASE, data_only=True)

    dim_equipment = sheet_df(base, "Dim_Equipment")
    dim_program = sheet_df(base, "Dim_Program")
    dim_terminal = sheet_df(base, "Dim_Terminal")
    dim_account = sheet_df(base, "Dim_Account")
    dim_date = sheet_df(base, "Dim_Date")
    fact_movements = sheet_df(base, "Fact_Movements")
    fact_loads = sheet_df(base, "Fact_Loads")
    fact_vehicle_performance = sheet_df(base, "Fact_Vehicle_Performance")
    fact_idle_events = sheet_df(base, "Fact_Idle_Events")
    fact_maintenance_wo = sheet_df(base, "Fact_Maintenance_WO")
    fact_asset_economics = sheet_df(base, "Fact_Asset_Economics")
    qa_source_log = sheet_df(base, "QA_Source_Log")
    qa_exceptions = sheet_df(base, "QA_Exceptions")
    data_dictionary = sheet_df(base, "Data_Dictionary")

    def norm(v):
        return str(v).strip().upper() if v else None

    tractor_key = {norm(row.AssetID): row.AssetKey for row in dim_equipment.itertuples()}
    tractor_ownership = {norm(row.AssetID): row.OwnershipClass for row in dim_equipment.itertuples()}
    tractor_division = {norm(row.AssetID): row.Division for row in dim_equipment.itertuples()}
    tractor_terminal = {norm(row.AssetID): row.Terminal for row in dim_equipment.itertuples()}

    pm_records = json.loads(PM_DUE_JSON.read_text(encoding="utf-8"))
    pm_rows = []
    n_unmatched = 0
    for r in pm_records:
        ak = tractor_key.get(norm(r["assetId"]))
        if ak is None:
            n_unmatched += 1
        pm_rows.append({
            "AssetKey": ak,
            "AssetID": r["assetId"],
            "OwnershipClass": tractor_ownership.get(norm(r["assetId"])),
            "Division": tractor_division.get(norm(r["assetId"])),
            "Terminal": tractor_terminal.get(norm(r["assetId"])),
            "Facility": r["facility"],
            "Group": r["group"],
            "AssetDescription": r["assetDescription"],
            "License": r["license"] or None,
            "Operator": r["operator"] or None,
            "Location": r["location"] or None,
            "PMCode": r["pmCode"],
            "PMDescription": r["pmDescription"],
            "CycleType": r["cycleType"],
            "Interval": parse_int(r["interval"]),
            "PreviousDone": r["previousDone"],
            "Current": r["current"],
            "DueAt": r["dueAt"],
            "DueIn": parse_int(r["dueIn"]),
            "DueUnit": "days" if r["cycleType"] == "Date" else ("miles" if r["cycleType"] == "Miles" else None),
            "DueStatus": r["dueStatus"],
            "Scheduled": r["scheduled"],
            "IsOverdue": r["dueStatus"] == "Past Due",
        })
    fact_pm_due = pd.DataFrame(pm_rows)

    n_total = len(fact_pm_due)
    n_past_due = int((fact_pm_due["DueStatus"] == "Past Due").sum())
    n_assets = fact_pm_due["AssetID"].nunique()

    # ---- supersede the v7 manual-pull QA_Source_Log row rather than duplicate it ----
    manual_mask = qa_source_log["Source"] == "RTA Asset and Equipment PM Due (all facilities)"
    qa_source_log.loc[manual_mask, "Status"] = "SUPERSEDED"
    qa_source_log.loc[manual_mask, "Note"] = (
        qa_source_log.loc[manual_mask, "Note"] + " -- SUPERSEDED: pipeline is now automated, see PASS row below."
    )

    new_source_row = pd.DataFrame([[
        "RTA Asset and Equipment PM Due (automated)",
        "PM due-dates, intervals, overdue status per asset",
        "2026-08-08", n_total, "PASS",
        (f"First fully-automated pull: RTA scheduled report -> email (no-reply@rtafleet.com) -> "
         "Gmail filter (label 'RTA Reports') -> Apps Script -> Drive ('RTA Reports Inbox' folder). "
         f"{n_total - n_unmatched} of {n_total} matched to Dim_Equipment ({n_unmatched} unmatched). "
         f"{n_past_due} of {n_total} ({round(n_past_due/n_total*100)}%) Past Due -- same-day pull as "
         "the prior manual version, so figures are unchanged this round; will move day-over-day once "
         "the schedule has run more than once."),
    ]], columns=list(qa_source_log.columns))
    qa_source_log = pd.concat([qa_source_log, new_source_row], ignore_index=True)

    # ================= README =================
    readme_lines = [
        ["RDW EOS -- MASTER DATA MODEL (v8)"],
        ["Built 2026-08-08 from RDW_EOS_Master_v7.xlsx. Fact_PM_Due refreshed (full replace) from the first automated pull."],
        [""],
        ["WHAT CHANGED FROM v7"],
        ["1. PM Due pipeline is now automated end to end: RTA scheduled report -> email -> Gmail filter -> Apps Script -> Drive folder. No manual export step anymore."],
        [f"2. Fact_PM_Due refreshed: {n_total} records, {n_assets} assets, {n_past_due} Past Due ({round(n_past_due/n_total*100)}%). Same-day pull as v7's manual version, so the numbers match -- next scheduled run will show real day-over-day movement."],
        ["3. All other tables carried forward from v7 UNCHANGED."],
        [""],
        ["STILL OPEN"],
        ["- McLeod (Movements/Loads/Equipment List) is not yet automated -- still a manual export. Being worked next."],
        ["- Idle Events is still a one-day sample (2026-01-01 only)."],
        ["- Trailer 1831's conflicting title records: still unresolved."],
        ["- 49 lettered tractors have Division but not a specific Style (PTO/Straight Truck/Van) sub-type."],
    ]
    readme_df = pd.DataFrame(readme_lines)

    # ================= Write workbook =================
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
        readme_df.to_excel(writer, sheet_name="README", index=False, header=False)
        dim_equipment.to_excel(writer, sheet_name="Dim_Equipment", index=False)
        dim_program.to_excel(writer, sheet_name="Dim_Program", index=False)
        dim_terminal.to_excel(writer, sheet_name="Dim_Terminal", index=False)
        dim_account.to_excel(writer, sheet_name="Dim_Account", index=False)
        dim_date.to_excel(writer, sheet_name="Dim_Date", index=False)
        fact_movements.to_excel(writer, sheet_name="Fact_Movements", index=False)
        fact_loads.to_excel(writer, sheet_name="Fact_Loads", index=False)
        fact_vehicle_performance.to_excel(writer, sheet_name="Fact_Vehicle_Performance", index=False)
        fact_idle_events.to_excel(writer, sheet_name="Fact_Idle_Events", index=False)
        fact_maintenance_wo.to_excel(writer, sheet_name="Fact_Maintenance_WO", index=False)
        fact_pm_due.to_excel(writer, sheet_name="Fact_PM_Due", index=False)
        fact_asset_economics.to_excel(writer, sheet_name="Fact_Asset_Economics", index=False)
        qa_source_log.to_excel(writer, sheet_name="QA_Source_Log", index=False)
        qa_exceptions.to_excel(writer, sheet_name="QA_Exceptions", index=False)
        data_dictionary.to_excel(writer, sheet_name="Data_Dictionary", index=False)

    print("=== BUILD SUMMARY (v8) ===")
    print(f"Fact_PM_Due: {n_total} rows ({n_unmatched} unmatched), {n_past_due} Past Due ({round(n_past_due/n_total*100)}%)")
    print(f"QA_Source_Log rows: {len(qa_source_log)}")
    print(f"\nWrote {OUT}")


if __name__ == "__main__":
    main()
