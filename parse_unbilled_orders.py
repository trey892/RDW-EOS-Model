"""
Parse McLeod's "Unbilled Orders Report" -- extracted to plain text via Drive's
read_file_content (this is a PDF; no local PDF library is used, the text was
pulled agent-side and saved as-is). Only the general "Unbilled Orders" report
is needed (not the "BHM Waste Unbilled" / "Roll Off Unbilled Orders"
terminal-filtered variants) -- every order line already carries its own
terminal in the "Rev Code" column (RO, BHM, HOU, etc.), so the one combined
report covers every terminal.

Order line shape (one order per line in the extracted text):
  <order#> <del date> <on hold Y/N> <ready-to-bill Y/N> <transferred Y/N>
  <customer code> - <customer name> $<total charges> <rev code/terminal>
  <age since delivery, days> <status>

"Dispatcher ... totals:" / "Report totals:" lines are running subtotals in
the source and are NOT parsed as orders -- terminal-level totals are computed
here instead, from the actual order lines, so the numbers always reconcile
with what's listed.
"""
import json
import re
from datetime import datetime
from pathlib import Path

SRC = Path(__file__).parent / "data" / "raw" / "Unbilled_Orders_latest.txt"
OUT = Path(__file__).parent / "output" / "unbilled_orders_data.json"

ORDER_RE = re.compile(
    r"^(\d{6,8})\s+(\d{2}/\d{2}/\d{4})\s+(Yes|No)\s+(Yes|No)\s+(Yes|No)\s+"
    r"(\S+)\s*-\s*(.+?)\s+\$([\d,]+\.\d{2})\s+(\w+)\s+(\d+)\s+(\w+)\s*$"
)


def parse():
    text = SRC.read_text(encoding="utf-8")
    orders = []
    for line in text.splitlines():
        m = ORDER_RE.match(line.strip())
        if not m:
            continue
        (order_num, del_date, on_hold, ready, transferred,
         cust_code, cust_name, amount, terminal, age_days, status) = m.groups()
        orders.append({
            "orderNumber": order_num,
            "deliveryDate": del_date,
            "onHold": on_hold == "Yes",
            "readyToBill": ready == "Yes",
            "transferred": transferred == "Yes",
            "customerCode": cust_code,
            "customerName": cust_name.strip(),
            "totalCharges": float(amount.replace(",", "")),
            "terminal": terminal,
            "ageDays": int(age_days),
            "status": status,
        })
    return orders


def main():
    orders = parse()
    if not orders:
        raise RuntimeError(f"No order lines matched in {SRC} -- format may have changed, check ORDER_RE")

    by_terminal = {}
    for o in orders:
        t = by_terminal.setdefault(o["terminal"], {
            "terminal": o["terminal"], "orders": 0, "uniqueOrders": set(),
            "ready": 0, "notReady": 0, "totalCharges": 0.0, "oldestDays": 0,
        })
        t["orders"] += 1
        t["uniqueOrders"].add(o["orderNumber"])
        t["ready" if o["readyToBill"] else "notReady"] += 1
        t["totalCharges"] += o["totalCharges"]
        t["oldestDays"] = max(t["oldestDays"], o["ageDays"])

    rows = []
    for t in by_terminal.values():
        rows.append({
            "terminal": t["terminal"],
            "orders": t["orders"],
            "uniqueOrders": len(t["uniqueOrders"]),
            "ready": t["ready"],
            "notReady": t["notReady"],
            "totalCharges": round(t["totalCharges"], 2),
            "oldestDays": t["oldestDays"],
        })
    rows.sort(key=lambda r: -r["totalCharges"])

    out = {
        "checkedAt": datetime.now().isoformat(timespec="seconds"),
        "source": "McLeod Unbilled Orders Report (McLeod Reports Inbox)",
        "orderCount": len(orders),
        "totalCharges": round(sum(o["totalCharges"] for o in orders), 2),
        "rows": rows,
    }

    print(f"Parsed {len(orders)} unbilled order line(s), ${out['totalCharges']:,.2f} total")
    for r in rows:
        print(f"  {r['terminal']}: {r['orders']} orders (${r['totalCharges']:,.2f}), "
              f"{r['ready']} ready / {r['notReady']} not ready, oldest {r['oldestDays']}d")

    OUT.write_text(json.dumps(out, indent=2), encoding="utf-8")
    print(f"\nWrote -> {OUT}")


if __name__ == "__main__":
    main()
