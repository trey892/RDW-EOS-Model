"""
Parse McLeod's "Service Incident Report By Customer" (McLeod Reports Inbox,
the "...Service Incedent Trey..." PDF -- McLeod's own filename, misspelling
and all; NOT the "...Service Incidents BHM Fleet..." variant, which is a
narrower single-terminal subset, or the "...Service Failures ALLL..." xlsx,
which is a single-day export). Extracted to plain text via Drive's
read_file_content, same pattern as parse_unbilled_orders.py -- no local PDF
library involved.

The report has no Terminal column at all (it's grouped by Customer). Each
incident line does carry a "Driver Manager" code, and separately parse_tractor_status.py
already captures each tractor's Dispatcher + Fleet from the RTD workbook --
cross-checking the two shows Driver Manager codes in this report and Dispatcher
codes in the RTD tractor listing are the same person codes (e.g. "scotth" ->
BHM in both), so terminal here is DERIVED: Driver Manager -> majority Fleet
from that same-day RTD pull. This is a real, data-backed mapping, not an
invented one, but it IS an inference, not a field McLeod reports directly --
say so on the dashboard. Roughly 6% of incident rows either don't match the
row-1 line shape or have a Driver Manager code with no RTD dispatcher match
(e.g. "lmeadm", an LME system login, not a person tied to one terminal) --
those are counted and surfaced as "unmapped", never silently dropped or
folded into a guessed terminal.
"""
import json
import re
from datetime import datetime
from pathlib import Path

from dashboard_filters import is_excluded_terminal

SRC = Path(__file__).parent / "data" / "raw" / "Service_Incidents_latest.txt"
TRACTOR_STATUS = Path(__file__).parent / "output" / "tractor_status_data.json"
OUT = Path(__file__).parent / "output" / "service_incidents_data.json"

CANONICAL_TERMINALS = ["BHM", "MOB", "NMOB", "HOU", "ATL", "WIL", "RO"]

DATE_RANGE_RE = re.compile(r"Scheduled arrival date:\s*(\d{2}/\d{2}/\d{4})(?:\s*-\s*(\d{2}/\d{2}/\d{4}))?")
ROW1_RE = re.compile(
    r"^(\d{6,8}) (\S+) (\S+) (\S+) (\S+) (\d{6,8}) (Shipper|Pickup|Delivery|Consignee)$"
)
TOTAL_INCIDENTS_RE = re.compile(r"Total service incidents:\s*(\d+)")
TOTAL_ORDERS_RE = re.compile(r"Total orders:\s*(\d+)")
STOP_ONTIME_RE = re.compile(r"Stop on-time percent:\s*([\d.]+)%")
ORDER_ONTIME_RE = re.compile(r"Order on-time percent:\s*([\d.]+)%")
ORDERS_WITH_INCIDENTS_RE = re.compile(r"Orders with service incidents:\s*(\d+)")


def canonical_terminal(code):
    code = (code or "").strip().upper()
    if not code or is_excluded_terminal(code):
        return None
    for base in CANONICAL_TERMINALS:
        if code.startswith(base):
            return base
    return None


def build_dispatcher_terminal_map():
    """Driver Manager code -> majority Fleet/terminal, from the same-day RTD tractor listing."""
    if not TRACTOR_STATUS.exists():
        return {}
    records = json.loads(TRACTOR_STATUS.read_text(encoding="utf-8"))
    from collections import defaultdict, Counter
    tally = defaultdict(Counter)
    for rec in records.values():
        dispatcher = str(rec.get("dispatcher") or "").strip().lower()
        terminal = canonical_terminal(rec.get("fleet"))
        if dispatcher and terminal:
            tally[dispatcher][terminal] += 1
    return {d: counter.most_common(1)[0][0] for d, counter in tally.items()}


def parse():
    text = SRC.read_text(encoding="utf-8")

    date_match = DATE_RANGE_RE.search(text)
    start_date = date_match.group(1) if date_match else None
    end_date = date_match.group(2) if date_match and date_match.group(2) else start_date

    driver_managers = []
    for line in text.splitlines():
        m = ROW1_RE.match(line.strip())
        if m:
            driver_managers.append(m.group(3))

    # Report's own footer totals ("Report totals: ...") -- last occurrence in the
    # text is the grand total (earlier ones are per-customer subtotals).
    report_total_incidents = None
    stop_ontime_pct = None
    order_ontime_pct = None
    total_orders = None
    orders_with_incidents = None
    tail_idx = text.rfind("Report totals:")
    if tail_idx != -1:
        tail = text[tail_idx:]
        m = TOTAL_INCIDENTS_RE.search(tail)
        report_total_incidents = int(m.group(1)) if m else None
        m = TOTAL_ORDERS_RE.search(tail)
        total_orders = int(m.group(1)) if m else None
        m = STOP_ONTIME_RE.search(tail)
        stop_ontime_pct = float(m.group(1)) if m else None
        m = ORDER_ONTIME_RE.search(tail)
        order_ontime_pct = float(m.group(1)) if m else None
        m = ORDERS_WITH_INCIDENTS_RE.search(tail)
        orders_with_incidents = int(m.group(1)) if m else None

    return {
        "startDate": start_date,
        "endDate": end_date,
        "driverManagers": driver_managers,
        "reportTotalIncidents": report_total_incidents,
        "totalOrders": total_orders,
        "ordersWithIncidents": orders_with_incidents,
        "stopOnTimePct": stop_ontime_pct,
        "orderOnTimePct": order_ontime_pct,
    }


def main():
    parsed = parse()
    dispatcher_terminal = build_dispatcher_terminal_map()

    by_terminal = {}
    unmapped = 0
    for dmgr in parsed["driverManagers"]:
        terminal = dispatcher_terminal.get(dmgr.strip().lower())
        if terminal:
            by_terminal[terminal] = by_terminal.get(terminal, 0) + 1
        else:
            unmapped += 1

    rows = [{"terminal": t, "count": n} for t, n in by_terminal.items()]
    rows.sort(key=lambda r: -r["count"])

    parsed_rows = len(parsed["driverManagers"])
    report_total = parsed["reportTotalIncidents"]
    unparsed = (report_total - parsed_rows) if report_total is not None and report_total >= parsed_rows else 0

    if parsed["startDate"] == parsed["endDate"]:
        date_range_label = parsed["startDate"] or "Unknown period"
    else:
        date_range_label = f"{parsed['startDate']} - {parsed['endDate']}" if parsed["startDate"] else "Unknown period"

    out = {
        "checkedAt": datetime.now().isoformat(timespec="seconds"),
        "source": "McLeod Service Incident Report By Customer (McLeod Reports Inbox)",
        "methodologyNote": "Terminal is derived: Driver Manager code on each incident, mapped to that "
                            "code's majority Fleet in the same-day RTD tractor listing. McLeod does not "
                            "report a Terminal directly on this export.",
        "dateRangeLabel": date_range_label,
        "reportTotalIncidents": report_total,
        "totalOrders": parsed["totalOrders"],
        "ordersWithIncidents": parsed["ordersWithIncidents"],
        "stopOnTimePct": parsed["stopOnTimePct"],
        "orderOnTimePct": parsed["orderOnTimePct"],
        "parsedIncidents": parsed_rows,
        "unparsedIncidents": unparsed,  # rows the row-shape regex didn't match at all
        "unmappedIncidents": unmapped,  # rows parsed fine, but Driver Manager has no RTD terminal match
        "rows": rows,
    }

    print(f"=== Service Incidents By Terminal: {date_range_label} ===")
    print(f"  Report total: {report_total}, parsed: {parsed_rows}, unparsed: {unparsed}, unmapped: {unmapped}")
    for r in rows:
        print(f"  {r['terminal']}: {r['count']}")

    OUT.write_text(json.dumps(out, indent=2), encoding="utf-8")
    print(f"\nWrote -> {OUT}")


if __name__ == "__main__":
    main()
