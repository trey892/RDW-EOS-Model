"""
Runs the full daily refresh pipeline in order. Expects, before this runs:
  - output/RDW_EOS_Master_latest.xlsx already in place (the current canonical model)
  - data/raw/Asset_and_Equipment_PMs_Due_latest.csv already in place (freshest RTA pull)
  - output/lease_data.json already in place (lease data isn't automated -- carried forward as-is)
  - The following are all OPTIONAL -- if any is missing, that step is skipped and its
    prior output/*.json is carried forward rather than failing the run:
      data/raw/RTD_latest.xlsx             (dispatch status/style/current-hub mileage)
      data/raw/Orientation_latest.xlsm     (Orientation Inbound)
      data/raw/PM_Compliance_latest.csv    (PM Compliance % by vehicle)
      data/raw/Unbilled_Orders_latest.txt  (Unbilled Orders by terminal, PDF text
                                             pre-extracted via Drive's read_file_content)

Produces, in order:
  1. output/pm_due_data.json          (parse_pm_due)
  2. output/RDW_EOS_Master_latest.xlsx refreshed in place (refresh_model)
  3. output/tractor_status_data.json, orientation_data.json, pm_compliance_data.json,
     unbilled_orders_data.json (parse_tractor_status, parse_orientation,
     parse_pm_compliance, parse_unbilled_orders)
  4. output/dashboard_data.json, asset_data.json, maintenance_data.json (build_*_data)
  5. output/RDW_Fleet_Dashboard.html  (assemble_dashboard)

Steps 1-2 are skipped if no new PM Due CSV was found (see SKIP_MODEL_REFRESH below) --
the daily routine sets that after checking Drive for a new file, so a quiet night
still regenerates dashboard JSON/HTML from the unchanged model rather than erroring.

Does NOT talk to Google Drive or the Artifact tool itself -- the calling cloud
routine handles fetching inputs beforehand and pushing outputs (Drive upload,
Artifact publish) afterward, since only the agent (not this script) has those tools.
"""
import sys
from pathlib import Path

import parse_pm_due
import refresh_model
import parse_tractor_status
import parse_orientation
import parse_pm_compliance
import parse_unbilled_orders
import derive_lease_fields
import build_dashboard_data
import build_asset_data
import build_maintenance_data
import build_sources_data
import assemble_dashboard


def _run_optional(label, module):
    """Run module.main() if its SRC file exists; otherwise skip and keep prior output."""
    if module.SRC.exists():
        print(f"\n--- {label} ---")
        module.main()
    else:
        print(f"\n--- {label} skipped (no {module.SRC.name} found -- keeping prior output) ---")


def main():
    skip_model_refresh = "--skip-model-refresh" in sys.argv

    if not skip_model_refresh:
        print("\n--- [1/6] parse_pm_due ---")
        parse_pm_due.main()

        print("\n--- [2/6] refresh_model ---")
        refresh_model.main()
    else:
        print("\n--- [1-2/6] skipped (no new PM Due pull) ---")

    _run_optional("[3/6] parse_tractor_status", parse_tractor_status)
    _run_optional("[3b/6] parse_orientation", parse_orientation)
    _run_optional("[3c/6] parse_pm_compliance", parse_pm_compliance)
    _run_optional("[3d/6] parse_unbilled_orders", parse_unbilled_orders)

    print("\n--- derive_lease_fields ---")
    derive_lease_fields.main()

    print("\n--- [4/6] build_dashboard_data ---")
    build_dashboard_data.main()

    print("\n--- [5/6] build_asset_data ---")
    build_asset_data.main()

    print("\n--- [6/6] build_maintenance_data ---")
    build_maintenance_data.main()

    print("\n--- build_sources_data ---")
    build_sources_data.main()

    print("\n--- assemble_dashboard ---")
    assemble_dashboard.main()

    print("\n=== daily_refresh complete ===")


if __name__ == "__main__":
    main()
