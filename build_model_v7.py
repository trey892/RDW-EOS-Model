"""
Build RDW_EOS_Master_v7.xlsx from RDW_EOS_Master_v6.xlsx (base, unchanged) plus:
  Fact_PM_Due -- the PM due-date tracking that's been flagged as missing since v5.
  Parsed from RTA's "Asset and Equipment PM Due" export (data/raw/Asset_and_Equipment_PMs_Due.csv,
  via parse_pm_due.py -> output/pm_due_data.json), all facilities, pulled 2026-08-08.
"""
import json
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd

RAW = Path(__file__).parent / "data" / "raw"
BASE = Path(__file__).parent / "output" / "RDW_EOS_Master_v6.xlsx"
PM_DUE_JSON = Path(__file__).parent / "output" / "pm_due_data.json"
OUT = Path(__file__).parent / "output" / "RDW_EOS_Master_v7.xlsx"


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

    new_qa_source_rows = []
    new_qa_exception_rows = []

    # ================= Fact_PM_Due =================
    pm_records = json.loads(PM_DUE_JSON.read_text(encoding="utf-8"))

    def norm(v):
        return str(v).strip().upper() if v else None

    tractor_key = {
        norm(row.AssetID): row.AssetKey
        for row in dim_equipment.itertuples()
    }
    tractor_ownership = {norm(row.AssetID): row.OwnershipClass for row in dim_equipment.itertuples()}
    tractor_division = {norm(row.AssetID): row.Division for row in dim_equipment.itertuples()}
    tractor_terminal = {norm(row.AssetID): row.Terminal for row in dim_equipment.itertuples()}

    pm_rows = []
    n_unmatched = 0
    for r in pm_records:
        ak = tractor_key.get(norm(r["assetId"]))
        if ak is None:
            n_unmatched += 1
        due_in = parse_int(r["dueIn"])
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
            "DueIn": due_in,
            "DueUnit": "days" if r["cycleType"] == "Date" else ("miles" if r["cycleType"] == "Miles" else None),
            "DueStatus": r["dueStatus"],
            "Scheduled": r["scheduled"],
            "IsOverdue": r["dueStatus"] == "Past Due",
        })

    fact_pm_due = pd.DataFrame(pm_rows)

    n_total = len(fact_pm_due)
    n_matched = n_total - n_unmatched
    n_past_due = int((fact_pm_due["DueStatus"] == "Past Due").sum())
    n_assets = fact_pm_due["AssetID"].nunique()
    pct_past_due = round(n_past_due / n_total * 100)

    new_qa_source_rows.append([
        "RTA Asset and Equipment PM Due (all facilities)", "PM due-dates, intervals, overdue status per asset",
        "2026-08-08", n_total, "REVIEW",
        (f"Closes the PM due-date gap flagged since v5. {n_matched} of {n_total} PM records matched to "
         f"Dim_Equipment ({n_unmatched} unmatched, see QA_Exceptions). Pulled manually this round -- "
         "not yet on an automated daily schedule (RTA's report screen was failing when a broader/no facility "
         f"filter was applied; being worked). {n_past_due} of {n_total} records ({pct_past_due}%) "
         f"are Past Due across {n_assets} distinct assets -- a large backlog, not a data artifact; verify against "
         "RTA directly before acting on it at scale."),
    ])
    new_qa_exception_rows.append([
        "MEDIUM", "Fact_PM_Due", "PM records with no Dim_Equipment match", n_unmatched, None,
        "Asset ID on the PM Due export not found in Dim_Equipment -- likely units outside the current McLeod snapshot (same pattern seen with Samsara/RTA WO grid mismatches).",
        "Asset Manager",
    ])
    new_qa_exception_rows.append([
        "HIGH", "Fact_PM_Due", "PM records marked Past Due", int(n_past_due), None,
        f"{round(n_past_due/n_total*100)}% of all PM records across {n_assets} assets are overdue as of the 2026-08-08 pull. Confirm with maintenance whether this reflects real backlog or a data/process gap (e.g. completed PMs not being logged back into RTA) before treating it as a to-do list at face value.",
        "Maintenance Director",
    ])

    # ================= Assemble QA tables =================
    qa_source_log_cols = list(qa_source_log.columns)
    qa_source_log = pd.concat([qa_source_log, pd.DataFrame(new_qa_source_rows, columns=qa_source_log_cols)], ignore_index=True)
    qa_exceptions_cols = list(qa_exceptions.columns)
    qa_exceptions = pd.concat([qa_exceptions, pd.DataFrame(new_qa_exception_rows, columns=qa_exceptions_cols)], ignore_index=True)

    dd_new = pd.DataFrame([
        ["Fact_PM_Due", "Fact", "One PM (preventive maintenance) item due on one asset", "AssetID+PMCode",
         "RTA Asset and Equipment PM Due export", "Manual pull for now -- replace CSV and re-run parse_pm_due.py + build_model_v7.py. Daily automated schedule not yet live."],
    ], columns=list(data_dictionary.columns))
    data_dictionary = pd.concat([data_dictionary, dd_new], ignore_index=True)

    # ================= README =================
    readme_lines = [
        ["RDW EOS -- MASTER DATA MODEL (v7)"],
        ["Built 2026-08-08 from RDW_EOS_Master_v6.xlsx + RTA Asset and Equipment PM Due export (all facilities)."],
        [""],
        ["WHAT CHANGED FROM v6"],
        [f"1. NEW Fact_PM_Due table: {n_total} PM-due records across {n_assets} assets, all facilities. Closes the PM due-date tracking gap flagged since v5 -- has PMCode, Description, Interval, Previous Done, Due At, Due In (days or miles), Due Status (Past Due/Due Soon/Due Now/Not Due), and Scheduled flag per asset."],
        [f"2. {n_matched} of {n_total} records joined to Dim_Equipment for AssetKey/OwnershipClass/Division/Terminal; {n_unmatched} unmatched (see QA_Exceptions)."],
        [f"3. HIGH-severity QA flag: {n_past_due} of {n_total} PM records ({round(n_past_due/n_total*100)}%) are Past Due -- large enough that it needs verification with the maintenance team before being treated as ground truth (could be real backlog, or completed work not being closed out in RTA)."],
        ["4. All other tables carried forward from v6 UNCHANGED."],
        [""],
        ["STILL OPEN"],
        ["- This PM Due data was a manual pull, not yet an automated daily feed -- RTA's report screen was failing to load when the facility filter was cleared/widened. Once that's resolved (or a scheduled report is set up under a new schedule name, per standing instruction not to touch existing schedules), this becomes a live daily refresh."],
        ["- Idle Events is still a one-day sample (2026-01-01 only)."],
        ["- Trailer 1831's conflicting title records: still unresolved."],
        ["- 49 lettered tractors have Division but not a specific Style (PTO/Straight Truck/Van) sub-type."],
        ["- See QA_Source_Log and QA_Exceptions for full detail and confidence levels on everything added this build."],
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

    print("=== BUILD SUMMARY (v7) ===")
    print(f"Fact_PM_Due: {n_total} rows, {n_matched} matched ({n_unmatched} unmatched)")
    print(f"Past Due: {n_past_due} ({round(n_past_due/n_total*100)}%) across {n_assets} assets")
    print(f"QA_Source_Log rows: {len(qa_source_log)}  QA_Exceptions rows: {len(qa_exceptions)}")
    print(f"\nWrote {OUT}")


if __name__ == "__main__":
    main()
