"""
Runs the full daily refresh pipeline in order. Expects, before this runs:
  - output/RDW_EOS_Master_latest.xlsx already in place (the current canonical model)
  - data/raw/Asset_and_Equipment_PMs_Due_latest.csv already in place (freshest RTA pull)
  - data/raw/RTD_latest.xlsx already in place (freshest RTD Tractor Listing pull --
    optional; if missing, dispatch-status/style/current-hub data is carried forward
    from whatever tractor_status_data.json already exists rather than failing the run)
  - output/lease_data.json already in place (lease data isn't automated -- carried forward as-is)

Produces, in order:
  1. output/pm_due_data.json          (parse_pm_due)
  2. output/RDW_EOS_Master_latest.xlsx refreshed in place (refresh_model)
  3. output/tractor_status_data.json  (parse_tractor_status)
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
import build_dashboard_data
import build_asset_data
import build_maintenance_data
import assemble_dashboard


def main():
    skip_model_refresh = "--skip-model-refresh" in sys.argv

    if not skip_model_refresh:
        print("\n--- [1/6] parse_pm_due ---")
        parse_pm_due.main()

        print("\n--- [2/6] refresh_model ---")
        refresh_model.main()
    else:
        print("\n--- [1-2/6] skipped (no new PM Due pull) ---")

    if parse_tractor_status.SRC.exists():
        print("\n--- [3/6] parse_tractor_status ---")
        parse_tractor_status.main()
    else:
        print(f"\n--- [3/6] skipped (no {parse_tractor_status.SRC.name} found -- "
              "keeping prior tractor_status_data.json) ---")

    print("\n--- [4/6] build_dashboard_data ---")
    build_dashboard_data.main()

    print("\n--- [5/6] build_asset_data ---")
    build_asset_data.main()

    print("\n--- [6/6] build_maintenance_data ---")
    build_maintenance_data.main()

    print("\n--- assemble_dashboard ---")
    assemble_dashboard.main()

    print("\n=== daily_refresh complete ===")


if __name__ == "__main__":
    main()
