"""Extract a compact per-tractor JSON dataset from RDW_EOS_Master_latest.xlsx for the HTML dashboard."""
import json
from pathlib import Path

import openpyxl
import pandas as pd

SRC = Path(__file__).parent / "output" / "RDW_EOS_Master_latest.xlsx"
OUT = Path(__file__).parent / "output" / "dashboard_data.json"


def sheet_df(wb, name):
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    return pd.DataFrame(rows[1:], columns=rows[0])


def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)

    dim_equipment = sheet_df(wb, "Dim_Equipment")
    tractors = dim_equipment[dim_equipment["AssetType"] == "TRACTOR"].copy()

    econ = sheet_df(wb, "Fact_Asset_Economics")
    econ_wide = econ.pivot_table(index="AssetKey", columns="MeasureName", values="Value", aggfunc="first")

    movements = sheet_df(wb, "Fact_Movements")
    mv = movements.dropna(subset=["AssetKey"]).copy()
    mv_agg = mv.groupby("AssetKey").apply(
        lambda g: pd.Series({
            "MovementsCount": len(g),
            "LoadedMiles": g.loc[g["Loaded"] == "Yes", "Distance"].sum(),
            "EmptyMiles": g.loc[g["Loaded"] == "No", "Distance"].sum(),
        }),
        include_groups=False,
    )

    perf = sheet_df(wb, "Fact_Vehicle_Performance").set_index("AssetKey")

    records = []
    for row in tractors.itertuples():
        ak = row.AssetKey
        e = econ_wide.loc[ak] if ak in econ_wide.index else pd.Series(dtype=float)

        revenue_driver_actual = e.get("Annualized Revenue (Driver Actual)")
        revenue_one_week = e.get("Annualized Revenue")
        if pd.notna(revenue_driver_actual):
            best_revenue, revenue_confidence, revenue_basis = revenue_driver_actual, "MEDIUM", "Driver Actual (multi-month)"
        else:
            best_revenue, revenue_confidence, revenue_basis = revenue_one_week, "LOW", "One week x52"

        mv_row = mv_agg.loc[ak] if ak in mv_agg.index else None
        loaded = float(mv_row["LoadedMiles"]) if mv_row is not None else None
        empty = float(mv_row["EmptyMiles"]) if mv_row is not None else None
        deadhead_pct = (empty / (loaded + empty) * 100) if (loaded is not None and (loaded + empty) > 0) else None

        perf_row = perf.loc[ak] if ak in perf.index else None

        def g(series, key):
            v = series.get(key) if series is not None else None
            return None if v is None or pd.isna(v) else round(float(v), 2)

        records.append({
            "assetKey": ak,
            "assetId": row.AssetID,
            "ownership": row.OwnershipClass,
            "division": row.Division,
            "terminal": row.Terminal,
            "year": None if pd.isna(row.Year) else int(row.Year),
            "make": row.Make,
            "model": row.Model,
            "status": row.ServiceStatus,
            "revenue": g(e, None) if best_revenue is None else round(float(best_revenue), 2),
            "revenueConfidence": revenue_confidence,
            "revenueBasis": revenue_basis,
            "fuelCost": g(e, "Fuel Cost"),
            "maintenanceCost": g(e, "Maintenance Cost RDW Paid"),
            "insurance": g(e, "Insurance And Compliance"),
            "driverSettlement": g(e, "Driver Settlement"),
            "depreciation": g(e, "Depreciation"),
            "totalDirectCost": g(e, "Total RDW Direct Cost"),
            "contribution": g(e, "Unit Contribution"),
            "netBookValue": g(e, "Net Book Value"),
            "hubMileage": g(e, "Current Hub Mileage"),
            "annualizedDistance": g(e, "Annualized Distance"),
            "mpg": None if perf_row is None or pd.isna(perf_row.get("MPG")) else round(float(perf_row["MPG"]), 2),
            "idleTimePct": None if perf_row is None or pd.isna(perf_row.get("IdleTimePct")) else round(float(perf_row["IdleTimePct"]), 2),
            "movements": None if mv_row is None else int(mv_row["MovementsCount"]),
            "loadedMiles": None if loaded is None else round(loaded, 1),
            "emptyMiles": None if empty is None else round(empty, 1),
            "deadheadPct": None if deadhead_pct is None else round(deadhead_pct, 1),
            "fuelGallons": g(e, "Fuel Gallons"),
        })

    qa_log = sheet_df(wb, "QA_Source_Log")
    qa_status = qa_log["Status"].value_counts().to_dict()

    meta = {
        "builtFrom": "RDW_EOS_Master_latest.xlsx",
        "tractorCount": len(records),
        "fuelPeriod": "2026-05-06 to 2026-08-04",
        "fuelHistoryPeriod": "Aug 2025 to Jul 2026 (invoiced fuel-card, reconciled 2026-08-03)",
        "movementsPeriod": "McLeod export 2026-08-02",
        "qaStatus": qa_status,
    }

    payload = {"meta": meta, "tractors": records}
    OUT.write_text(json.dumps(payload, indent=None, default=str), encoding="utf-8")
    print(f"Wrote {len(records)} tractors, {OUT.stat().st_size:,} bytes -> {OUT}")

    # sanity prints
    df = pd.DataFrame(records)
    print(df.groupby("ownership")["contribution"].agg(["count", "sum", "mean"]))
    print(df.groupby("division")["contribution"].agg(["count", "sum", "mean"]))
    print("revenue confidence:", df["revenueConfidence"].value_counts().to_dict())
    print("mpg matched:", df["mpg"].notna().sum())
    print("fuelGallons matched:", df["fuelGallons"].notna().sum())


if __name__ == "__main__":
    main()
