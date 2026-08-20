"""
Computes the RDW Revenue tab's analytics from revenue_goal_data.json (this
week's parsed "Revenue Goal by Tractor" export) plus tractor_status_data.json
(RTD, already live in this pipeline -- used for fleet-wide composition and
the ghost-tractor / duplicate-serial checks below).

Replaces the old static 329-row snapshot (build_revenue_module.py, retired
2026-08-19) -- everything here is computed fresh from this week's real data,
not hand-embedded. What it does NOT reproduce from the old report:
  - Lane Economics (empty miles, short-haul/linehaul split) -- that needed
    the Inbound/Outbound Lane Analysis report, which has no recurring Drive
    source (the one full pull found was a one-time manual upload). Dropped
    rather than shown stale; add back if a recurring source shows up.
  - Samsara-sourced Asset Quality rows (Avg MPG, Avg 90-day safety score) --
    not parsed anywhere in this pipeline yet. Left as null/"--" rather than
    guessed.
  - The written narrative -- that's authored fresh each weekly refresh by
    whatever's running the scheduled task (see SKILL.md), reading this
    script's output and writing revenue_narrative.json alongside it. This
    script only computes numbers.
"""
import json
from datetime import datetime
from pathlib import Path
from statistics import median

SRC = Path(__file__).parent / "output" / "revenue_goal_data.json"
TRACTOR_STATUS = Path(__file__).parent / "output" / "tractor_status_data.json"
OUT = Path(__file__).parent / "output" / "revenue_analysis_data.json"

IDLE_DAYS_THRESHOLD = 5  # out of a 7-day report window


def quartiles(values):
    values = sorted(v for v in values if v is not None)
    if not values:
        return {"q1": None, "median": None, "q3": None, "min": None, "max": None, "n": 0}
    n = len(values)

    def pct(p):
        idx = p * (n - 1)
        lo, hi = int(idx), min(int(idx) + 1, n - 1)
        frac = idx - lo
        return round(values[lo] + (values[hi] - values[lo]) * frac, 2)

    return {"q1": pct(0.25), "median": pct(0.5), "q3": pct(0.75), "min": values[0], "max": values[-1], "n": n}


def main():
    if not SRC.exists():
        raise FileNotFoundError(f"{SRC} not found -- run parse_revenue_goal.py first")
    revenue = json.loads(SRC.read_text(encoding="utf-8"))
    tractors = revenue["tractors"]
    rtd = json.loads(TRACTOR_STATUS.read_text(encoding="utf-8")) if TRACTOR_STATUS.exists() else {}

    active = [t for t in tractors if t["tractor"] != "RDW"]  # "RDW" is a non-fleet placeholder row in the export

    totals = {
        "tractorCount": len(active),
        "revenueGoal": round(sum(t["revenueGoal"] for t in active), 2),
        "revenueActual": round(sum(t["revenueActual"] for t in active), 2),
        "distanceActual": sum(t["distanceActual"] for t in active),
    }
    totals["revenueDiff"] = round(totals["revenueActual"] - totals["revenueGoal"], 2)
    totals["revenuePct"] = round(totals["revenueActual"] / totals["revenueGoal"] * 100, 2) if totals["revenueGoal"] else None

    def ownership_bucket(items):
        rows = []
        by_own = {}
        for t in items:
            by_own.setdefault(t["ownership"], []).append(t)
        for own, ts in sorted(by_own.items(), key=lambda kv: -len(kv[1])):
            rev_actual = sum(t["revenueActual"] for t in ts)
            rev_goal = sum(t["revenueGoal"] for t in ts)
            rows.append({
                "ownership": own,
                "count": len(ts),
                "revenueActual": round(rev_actual, 2),
                "revenueGoal": round(rev_goal, 2),
                "revenuePct": round(rev_actual / rev_goal * 100, 2) if rev_goal else None,
                "revenuePerDay": quartiles([t["revenuePerDay"] for t in ts]),
                "revenuePerMile": quartiles([t["revenuePerMile"] for t in ts]),
            })
        return rows

    by_ownership = ownership_bucket(active)

    def terminal_bucket(items):
        rows = []
        by_term = {}
        for t in items:
            key = t["terminal"] or "UNAS"
            by_term.setdefault(key, []).append(t)
        for term, ts in sorted(by_term.items(), key=lambda kv: -sum(x["revenueActual"] for x in kv[1])):
            rev_actual = sum(t["revenueActual"] for t in ts)
            rev_goal = sum(t["revenueGoal"] for t in ts)
            rpd = [t["revenuePerDay"] for t in ts if t["revenuePerDay"] is not None]
            rpm = [t["revenuePerMile"] for t in ts if t["revenuePerMile"] is not None]
            rows.append({
                "terminal": term,
                "count": len(ts),
                "revenueActual": round(rev_actual, 2),
                "revenueGoal": round(rev_goal, 2),
                "revenuePct": round(rev_actual / rev_goal * 100, 2) if rev_goal else None,
                "medianRevenuePerDay": median(rpd) if rpd else None,
                "medianRevenuePerMile": round(median(rpm), 2) if rpm else None,
            })
        return rows

    by_terminal = terminal_bucket(active)

    # Idle iron: near-full-week tractors with zero revenue. A single week's zero is common
    # (maintenance, driver out) and NOT automatically a "ghost" -- only tractors that ALSO carry
    # a real Out Of Service date in RTD get called out as ghosts; the rest are just listed as-is.
    idle = [t for t in active if t["revenueActual"] == 0 and t["days"] >= IDLE_DAYS_THRESHOLD]
    idle_units = []
    ghost_count = 0
    for t in idle:
        rtd_rec = rtd.get(t["tractor"].upper(), {})
        oos_date = rtd_rec.get("outOfServiceDate")
        is_ghost = bool(oos_date)
        if is_ghost:
            ghost_count += 1
        idle_units.append({
            "tractor": t["tractor"], "driver": t["driver"], "terminal": t["terminal"],
            "ownership": t["ownership"], "days": t["days"], "revenueGoal": t["revenueGoal"],
            "outOfServiceDate": str(oos_date) if oos_date else None, "isGhost": is_ghost,
        })

    # Data integrity checks -- recomputed every run from current data, not carried over as fixed claims.
    unsuffixed = [t for t in active if t["tractor"] and not t["tractor"][-1].isalpha()]
    suffixed_company = [t for t in active if t["tractor"] and t["tractor"][-1].upper() in ("C", "P")]
    unsuffixed_rpd = [t["revenuePerDay"] for t in unsuffixed if t["revenuePerDay"] is not None]
    suffixed_rpd = [t["revenuePerDay"] for t in suffixed_company if t["revenuePerDay"] is not None]

    from collections import Counter
    serial_counts = Counter(rec.get("serialNumber") for rec in rtd.values() if rec.get("serialNumber"))
    duplicate_serials = [{"serialNumber": sn, "count": n} for sn, n in serial_counts.items() if n > 1]

    ownership_method_counts = dict(Counter(t["ownershipMethod"] for t in active))

    data_integrity = {
        "unsuffixedUnits": {
            "count": len(unsuffixed),
            "medianRevenuePerDay": median(unsuffixed_rpd) if unsuffixed_rpd else None,
        },
        "suffixedCompanyUnits": {
            "count": len(suffixed_company),
            "medianRevenuePerDay": median(suffixed_rpd) if suffixed_rpd else None,
        },
        "duplicateSerialNumbers": duplicate_serials,
        "ownershipClassification": {
            "rtdConfirmed": ownership_method_counts.get("rtd-owner", 0) + ownership_method_counts.get("rtd-confirmed-non-company", 0),
            "suffixFallback": ownership_method_counts.get("suffix-fallback", 0),
        },
    }

    # Fleet-wide composition from the live RTD snapshot (all active tractors, not just those with
    # revenue this week) -- more complete than counting only this week's revenue-report rows.
    fleet_composition = dict(Counter(
        "Company" if str(rec.get("owner") or "").strip().upper() == "RDW" else "Non-Company"
        for rec in rtd.values()
    ))

    out = {
        "checkedAt": datetime.now().isoformat(timespec="seconds"),
        "source": revenue["source"],
        "periodStart": revenue["startDate"],
        "periodEnd": revenue["endDate"],
        "totals": totals,
        "byOwnership": by_ownership,
        "byTerminal": by_terminal,
        "idleIron": {
            "thresholdDays": IDLE_DAYS_THRESHOLD,
            "count": len(idle_units),
            "ghostCount": ghost_count,
            "units": sorted(idle_units, key=lambda u: -u["revenueGoal"]),
        },
        "dataIntegrity": data_integrity,
        "fleetComposition": fleet_composition,
        "droppedSections": {
            "laneEconomics": "No recurring Drive source for the Inbound/Outbound Lane Analysis report -- "
                              "only a one-time manual upload was found (2026-08-11). Add back once a "
                              "recurring pull exists.",
            "samsaraAssetQuality": "Avg MPG and Avg 90-day safety score need Samsara data, not parsed "
                                    "anywhere in this pipeline yet.",
        },
    }

    print(f"=== Revenue Analysis: {revenue['startDate']} - {revenue['endDate']} ===")
    print(f"  {totals['tractorCount']} tractors, ${totals['revenueActual']:,.2f} actual vs "
          f"${totals['revenueGoal']:,.2f} goal ({totals['revenuePct']}%)")
    print(f"  Idle iron: {len(idle_units)} tractors with $0 revenue this week ({ghost_count} confirmed ghosts via RTD Out Of Service date)")
    print(f"  Unsuffixed units: {data_integrity['unsuffixedUnits']['count']}, "
          f"median rev/day ${data_integrity['unsuffixedUnits']['medianRevenuePerDay'] or 0:,.2f} "
          f"vs suffixed-Company ${data_integrity['suffixedCompanyUnits']['medianRevenuePerDay'] or 0:,.2f}")
    print(f"  Duplicate serials in RTD: {len(duplicate_serials)}")

    OUT.write_text(json.dumps(out, indent=2, default=str), encoding="utf-8")
    print(f"\nWrote -> {OUT}")


if __name__ == "__main__":
    main()
