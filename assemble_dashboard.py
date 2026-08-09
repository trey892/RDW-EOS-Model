"""
Assemble the final dashboard HTML by embedding the 4 JSON datasets into
dashboard_template.html in place of its placeholder markers. Run after
build_dashboard_data.py / build_asset_data.py / build_maintenance_data.py
(and extract_lease_data.py, if lease_data.json needs refreshing).

Output is CSP-safe for Artifact publishing: all data is embedded, no fetch calls.
"""
from pathlib import Path

OUT_DIR = Path(__file__).parent / "output"
TEMPLATE = OUT_DIR / "dashboard_template.html"
OUT = OUT_DIR / "RDW_Fleet_Dashboard.html"

PLACEHOLDERS = {
    "/*__DATA__*/": OUT_DIR / "dashboard_data.json",
    "/*__ASSET_DATA__*/": OUT_DIR / "asset_data.json",
    "/*__MAINT_DATA__*/": OUT_DIR / "maintenance_data.json",
    "/*__LEASE_DATA__*/": OUT_DIR / "lease_data.json",
}


def main():
    html = TEMPLATE.read_text(encoding="utf-8")

    for marker, data_file in PLACEHOLDERS.items():
        if marker not in html:
            raise ValueError(f"Marker {marker} not found in {TEMPLATE}")
        payload = data_file.read_text(encoding="utf-8")
        html = html.replace(marker, payload, 1)

    OUT.write_text(html, encoding="utf-8")
    print(f"Wrote {OUT} ({OUT.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
