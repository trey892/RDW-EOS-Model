"""
Refresh RDW_EOS_Master_latest.xlsx in place from a fresh RTA "Asset and Equipment
PM Due" pull (data/raw/Asset_and_Equipment_PMs_Due_latest.csv -> pm_due_data.json
via parse_pm_due.py, run before this script).

Fact_PM_Due is a full replace on each refresh (not additive) -- matches the
"Replace CSV and refresh" method in Data_Dictionary. Every other table is carried
forward unchanged: McLeod (Movements/Loads) and Samsara (fuel/idle) are not yet
automated, still manual exports.

Designed to run unattended (daily cloud routine): reads and writes the same
canonical filename, so no versioned filename needs to be tracked by hand --
bumps the simple incrementing counter in model_version.py instead.
"""
import datetime
import json
from pathlib import Path

import openpyxl
import pandas as pd

import model_version

MASTER = Path(__file__).parent / "output" / "RDW_EOS_Master_latest.xlsx"
PM_DUE_JSON = Path(__file__).parent / "output" / "pm_due_data.json"


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
    today = datetime.date.today().isoformat()

    base = openpyxl.load_workbook(MASTER, data_only=True)

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

    # ---- supersede any prior PM Due QA_Source_Log rows rather than duplicate ----
    prior_mask = qa_source_log["Source"].astype(str).str.startswith("RTA Asset and Equipment PM Due")
    still_current = qa_source_log["Status"] == "PASS"
    supersede_mask = prior_mask & still_current
    qa_source_log.loc[supersede_mask, "Status"] = "SUPERSEDED"
    qa_source_log.loc[supersede_mask, "Note"] = (
        qa_source_log.loc[supersede_mask, "Note"].astype(str)
        + f" -- SUPERSEDED by automated refresh on {today}."
    )

    new_source_row = pd.DataFrame([[
        "RTA Asset and Equipment PM Due (automated)",
        "PM due-dates, intervals, overdue status per asset",
        today, n_total, "PASS",
        (f"Automated daily pull: RTA scheduled report -> email (no-reply@rtafleet.com) -> "
         "Gmail filter (label 'RTA Reports') -> Apps Script -> Drive ('RTA Reports Inbox' folder). "
         f"{n_total - n_unmatched} of {n_total} matched to Dim_Equipment ({n_unmatched} unmatched). "
         f"{n_past_due} of {n_total} ({round(n_past_due/n_total*100)}%) Past Due."),
    ]], columns=list(qa_source_log.columns))
    qa_source_log = pd.concat([qa_source_log, new_source_row], ignore_index=True)

    # ================= README =================
    readme_lines = [
        ["RDW EOS -- MASTER DATA MODEL"],
        [f"Last automated refresh: {today}. Fact_PM_Due refreshed (full replace) from the latest RTA pull."],
        [""],
        ["AUTOMATED PIPELINE"],
        ["PM Due: RTA scheduled report -> email -> Gmail filter -> Apps Script -> Drive -> this refresh. Runs daily."],
        [f"Fact_PM_Due: {n_total} records, {n_assets} assets, {n_past_due} Past Due ({round(n_past_due/n_total*100)}%)."],
        ["All other tables carried forward unchanged from the prior refresh."],
        [""],
        ["STILL MANUAL / NOT YET AUTOMATED"],
        ["- McLeod (Movements/Loads/Equipment List) -- still a manual export."],
        ["- McLeod driver-revenue reports (4 PDFs) land in Drive automatically but are not yet parsed into the model."],
        ["- Samsara fuel/idle data -- still a manual export, not on a recurring pull."],
        ["- Trailer 1831's conflicting title records: still unresolved."],
        ["- 49 lettered tractors have Division but not a specific Style (PTO/Straight Truck/Van) sub-type."],
    ]
    readme_df = pd.DataFrame(readme_lines)

    # ================= Write workbook (temp file, then atomic replace) =================
    tmp = MASTER.with_suffix(".tmp.xlsx")
    with pd.ExcelWriter(tmp, engine="openpyxl") as writer:
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
    tmp.replace(MASTER)
    v = model_version.bump()

    print("=== REFRESH SUMMARY ===")
    print(f"Fact_PM_Due: {n_total} rows ({n_unmatched} unmatched), {n_past_due} Past Due ({round(n_past_due/n_total*100)}%)")
    print(f"QA_Source_Log rows: {len(qa_source_log)}")
    print(f"Model version: v{v['version']} (refreshed {v['lastRefreshed']})")
    print(f"\nRefreshed {MASTER}")


if __name__ == "__main__":
    main()
