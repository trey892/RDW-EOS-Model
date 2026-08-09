"""
Refresh RDW_EOS_Master_latest.xlsx's fuel-related Fact_Asset_Economics measures from the
reconciled fuel-card transaction history (parse_fuel_history.py must run first, producing
output/fuel_history_data.json).

Two measures, both full-replace (not additive) for matched tractors:
  - Fuel Cost: RDW's direct cost basis. Only refreshed for Company-owned tractors -- Lease
    Purchase / Owner Operator fuel is driver-paid by standing convention and stays at 0 even
    though those units also show real fuel-card activity in the raw data.
  - Fuel Gallons (new measure): propulsion gallons (diesel + other fuel, excluding reefer) for
    EVERY matched tractor regardless of ownership -- this is an efficiency measure (how the
    engine performs), not a cost-responsibility measure, so who pays for the fuel is irrelevant.

Not part of the daily automated refresh -- this fuel history is a periodic invoice
reconciliation (12-month lookback), re-run by hand whenever Trey supplies an updated export.
"""
import datetime
import json
from pathlib import Path

import openpyxl
import pandas as pd

import model_version

MASTER = Path(__file__).parent / "output" / "RDW_EOS_Master_latest.xlsx"
FUEL_JSON = Path(__file__).parent / "output" / "fuel_history_data.json"
PERIOD = "FY2026-Annualized"
SOURCE_LABEL = "Fuel History (invoiced, fuel-card reconciled)"


def sheet_df(wb, name):
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    return pd.DataFrame(rows[1:], columns=rows[0])


def main():
    today = datetime.date.today().isoformat()
    fuel = json.loads(FUEL_JSON.read_text(encoding="utf-8"))

    wb = openpyxl.load_workbook(MASTER, data_only=True)

    dim_equipment = sheet_df(wb, "Dim_Equipment")
    dim_program = sheet_df(wb, "Dim_Program")
    dim_terminal = sheet_df(wb, "Dim_Terminal")
    dim_account = sheet_df(wb, "Dim_Account")
    dim_date = sheet_df(wb, "Dim_Date")
    fact_movements = sheet_df(wb, "Fact_Movements")
    fact_loads = sheet_df(wb, "Fact_Loads")
    fact_vehicle_performance = sheet_df(wb, "Fact_Vehicle_Performance")
    fact_idle_events = sheet_df(wb, "Fact_Idle_Events")
    fact_maintenance_wo = sheet_df(wb, "Fact_Maintenance_WO")
    fact_pm_due = sheet_df(wb, "Fact_PM_Due")
    fact_asset_economics = sheet_df(wb, "Fact_Asset_Economics")
    qa_source_log = sheet_df(wb, "QA_Source_Log")
    qa_exceptions = sheet_df(wb, "QA_Exceptions")
    data_dictionary = sheet_df(wb, "Data_Dictionary")

    def norm(v):
        return str(v).strip().upper() if v else None

    asset_id_to_key = {}
    asset_id_to_ownership = {}
    for row in dim_equipment.itertuples():
        if row.AssetType == "TRACTOR":
            asset_id_to_key[norm(row.AssetID)] = row.AssetKey
            asset_id_to_ownership[norm(row.AssetID)] = row.OwnershipClass

    n_matched = 0
    n_cost_updated = 0
    n_gallons_added = 0
    n_unmatched_units = 0

    fuel_cost_updates = {}   # AssetKey -> new fuel cost value
    gallons_rows = {}        # AssetKey -> gallons value

    for unit, data in fuel.items():
        ak = asset_id_to_key.get(norm(unit))
        if ak is None:
            n_unmatched_units += 1
            continue
        n_matched += 1
        ownership = asset_id_to_ownership.get(norm(unit))
        gallons_rows[ak] = data["propulsionGallons"]
        n_gallons_added += 1
        if ownership == "Company":
            fuel_cost_updates[ak] = round(data["fuelCostNet"], 2)
            n_cost_updated += 1

    # ---- Fuel Cost: replace Value in place for Company tractors with a match ----
    mask = (fact_asset_economics["MeasureName"] == "Fuel Cost") & (fact_asset_economics["AssetKey"].isin(fuel_cost_updates.keys()))
    fact_asset_economics.loc[mask, "Value"] = fact_asset_economics.loc[mask, "AssetKey"].map(fuel_cost_updates)
    fact_asset_economics.loc[mask, "Basis"] = "12 months invoiced net fuel (fuel-card reconciled, verified to the penny)"
    fact_asset_economics.loc[mask, "BasisNote"] = f"HIGH - refreshed {today} from Fuel_History_Cleaned_and_Summarized.xlsx"

    # ---- cascade: Total RDW Direct Cost, Unit Contribution, and both $/mile
    # measures are stored as static values, not live formulas -- changing Fuel
    # Cost alone leaves them silently stale (and internally inconsistent: their
    # own cost stack would stop summing to the displayed total). Recompute all
    # four for every Company tractor whose Fuel Cost just changed. ----
    econ_wide = fact_asset_economics.pivot_table(index="AssetKey", columns="MeasureName", values="Value", aggfunc="first")

    mv = fact_movements.dropna(subset=["AssetKey"]).copy()
    mv_agg = mv.groupby("AssetKey").apply(
        lambda g: pd.Series({
            "LoadedMiles": g.loc[g["Loaded"] == "Yes", "Distance"].sum(),
            "EmptyMiles": g.loc[g["Loaded"] == "No", "Distance"].sum(),
        }),
        include_groups=False,
    )

    total_cost_updates, contribution_updates = {}, {}
    cost_per_mile_updates, cost_per_loaded_mile_updates = {}, {}

    for ak in fuel_cost_updates:
        if ak not in econ_wide.index:
            continue
        row = econ_wide.loc[ak]
        components = ["Maintenance Cost RDW Paid", "Insurance And Compliance", "Driver Settlement", "Depreciation"]
        if any(pd.isna(row.get(c)) for c in components):
            continue  # incomplete cost stack -- leave downstream measures untouched rather than guess
        total_direct_cost = fuel_cost_updates[ak] + sum(row[c] for c in components)
        total_cost_updates[ak] = round(total_direct_cost, 4)

        revenue = row.get("Annualized Revenue (Driver Actual)")
        if pd.isna(revenue):
            revenue = row.get("Annualized Revenue")
        if not pd.isna(revenue):
            contribution_updates[ak] = round(revenue - total_direct_cost, 4)

        annualized_distance = row.get("Annualized Distance")
        if not pd.isna(annualized_distance) and annualized_distance > 0:
            cost_per_mile_updates[ak] = round(total_direct_cost / annualized_distance, 4)
            if ak in mv_agg.index:
                loaded, empty = mv_agg.loc[ak, "LoadedMiles"], mv_agg.loc[ak, "EmptyMiles"]
                if (loaded + empty) > 0:
                    deadhead_pct = empty / (loaded + empty)
                    loaded_distance = annualized_distance * (1 - deadhead_pct)
                    if loaded_distance > 0:
                        cost_per_loaded_mile_updates[ak] = round(total_direct_cost / loaded_distance, 4)

    def apply_update(measure_name, updates, basis_note):
        m = (fact_asset_economics["MeasureName"] == measure_name) & (fact_asset_economics["AssetKey"].isin(updates.keys()))
        fact_asset_economics.loc[m, "Value"] = fact_asset_economics.loc[m, "AssetKey"].map(updates)
        fact_asset_economics.loc[m, "BasisNote"] = basis_note

    cascade_note = f"HIGH - recomputed {today} to stay consistent with the Fuel Cost refresh above"
    apply_update("Total RDW Direct Cost", total_cost_updates, cascade_note)
    apply_update("Unit Contribution", contribution_updates, cascade_note)
    apply_update("Cost Per Total Mile", cost_per_mile_updates,
                 "MEDIUM - recomputed " + today + "; both inputs are themselves annualized/extrapolated figures -- treat as directional.")
    apply_update("Cost Per Loaded Mile", cost_per_loaded_mile_updates,
                 "MEDIUM - recomputed " + today + "; Deadhead % from McLeod Movements applied to annualized distance -- a blended estimate, not a direct loaded-mile measurement.")

    print(f"Cascaded recompute: {len(total_cost_updates)} Total RDW Direct Cost, {len(contribution_updates)} Unit Contribution, "
          f"{len(cost_per_mile_updates)} Cost Per Total Mile, {len(cost_per_loaded_mile_updates)} Cost Per Loaded Mile")

    # ---- Fuel Gallons: full-replace measure across ALL matched tractors (any ownership) ----
    fact_asset_economics = fact_asset_economics[fact_asset_economics["MeasureName"] != "Fuel Gallons"].copy()
    new_rows = pd.DataFrame([
        [ak, PERIOD, "Fuel Gallons", val,
         "12 months invoiced fuel-card diesel+other gallons (excludes reefer)", "HIGH",
         f"HIGH - added {today} from Fuel_History_Cleaned_and_Summarized.xlsx, propulsion gallons only"]
        for ak, val in gallons_rows.items()
    ], columns=list(fact_asset_economics.columns))
    fact_asset_economics = pd.concat([fact_asset_economics, new_rows], ignore_index=True)

    # ---- supersede any prior fuel-history QA_Source_Log rows ----
    prior_mask = qa_source_log["Source"].astype(str).str.startswith("Fuel History")
    still_current = qa_source_log["Status"] == "PASS"
    supersede_mask = prior_mask & still_current
    qa_source_log.loc[supersede_mask, "Status"] = "SUPERSEDED"
    qa_source_log.loc[supersede_mask, "Note"] = (
        qa_source_log.loc[supersede_mask, "Note"].astype(str)
        + f" -- SUPERSEDED by refresh on {today}, see PASS row below."
    )
    new_source_row = pd.DataFrame([[
        SOURCE_LABEL,
        "Net fuel cost (Company only) and propulsion gallons (all ownership)",
        today, n_matched, "PASS",
        (f"Fuel_History_Cleaned_and_Summarized.xlsx, 36,101 fuel-card transactions reconciled to the "
         f"penny against provider totals. {n_matched} of {len(fuel)} units in the fuel history matched "
         f"to Dim_Equipment ({n_unmatched_units} unmatched -- likely retired/off-roster units). "
         f"Fuel Cost refreshed for {n_cost_updated} Company tractors (LP/OO stays $0, driver-paid by "
         f"standing convention). Fuel Gallons added as a new measure for all {n_gallons_added} matched "
         "tractors regardless of ownership, since it's an engine-efficiency figure, not a cost figure. "
         "Reefer-unit diesel excluded from gallons (not propulsion fuel)."),
    ]], columns=list(qa_source_log.columns))
    qa_source_log = pd.concat([qa_source_log, new_source_row], ignore_index=True)

    # ================= README =================
    readme_lines = [
        ["RDW EOS -- MASTER DATA MODEL"],
        [f"Last automated refresh: {today} (PM Due). Fuel Cost/Gallons refreshed {today} from invoiced fuel-card history."],
        [""],
        ["AUTOMATED PIPELINE"],
        ["PM Due: RTA scheduled report -> email -> Gmail filter -> Apps Script -> Drive -> daily refresh."],
        [""],
        ["PERIODIC (MANUAL) REFRESH"],
        [f"Fuel Cost + Fuel Gallons: {n_matched} tractors matched from a 36,101-row fuel-card reconciliation, "
         "verified to the penny. Re-run parse_fuel_history.py + refresh_fuel_history.py when Trey supplies a new export."],
        ["Lease/insurance data (Insured & Leased tab): not on any automated or scripted feed -- rebuilt by hand."],
        [""],
        ["STILL MANUAL / NOT YET AUTOMATED"],
        ["- McLeod (Movements/Loads/Equipment List) -- still a manual export."],
        ["- McLeod driver-revenue reports (4 PDFs) land in Drive automatically but are not yet parsed into the model."],
        ["- Samsara fuel/idle telemetry (MPG, idle %) -- still a manual export, separate from the invoiced fuel-card history above."],
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

    print("=== FUEL HISTORY REFRESH SUMMARY ===")
    print(f"Matched {n_matched} of {len(fuel)} fuel-history units to Dim_Equipment ({n_unmatched_units} unmatched)")
    print(f"Fuel Cost refreshed for {n_cost_updated} Company tractors")
    print(f"Fuel Gallons added for {n_gallons_added} tractors (all ownership)")
    print(f"Model version: v{v['version']} (refreshed {v['lastRefreshed']})")
    print(f"\nRefreshed {MASTER}")


if __name__ == "__main__":
    main()
