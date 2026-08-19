"""
Adds derived fields to output/lease_data.json that the Trailer Master and
Insurance tabs need but the underlying lease/insurance source doesn't carry
directly: trailerEquipmentType, trailerMasterStatus, plus the Trailer Master
ticker's summary meta counts.

Runs every refresh (not a one-time extract) because lease_data.json itself
gets refreshed from the canonical data repo copy each day (see SKILL.md step
3) -- these derived fields need to be recomputed against whatever units are
actually in that file each time, not baked in once.

TRAILER_MASTER_OOS_IDS is a manual out-of-service override list, ported as-is
from the prior dashboard build -- same list as build_asset_data.py's copy
(kept in sync manually; not derived from any live source). Review/update
periodically rather than trusting it stays accurate indefinitely.
"""
import json
from pathlib import Path

LEASE_JSON = Path(__file__).parent / "output" / "lease_data.json"

TRAILER_MASTER_OOS_IDS = {"1130", "1587", "1637", "462", "AES132", "AES137", "AES139"}


def main():
    data = json.loads(LEASE_JSON.read_text(encoding="utf-8"))
    units = data["units"]

    trailers = [u for u in units if u.get("assetType") == "TRAILER"]
    for u in trailers:
        raw_type = str(u.get("model") or "").strip().upper()
        u["trailerEquipmentType"] = ("R/O" if raw_type == "R-O" else raw_type) or None
        asset_id_norm = str(u.get("assetId") or "").strip().upper()
        u["trailerMasterStatus"] = "Out of Service" if asset_id_norm in TRAILER_MASTER_OOS_IDS else (u.get("status") or "Active")

    data["meta"]["trailerMasterRawCount"] = len(trailers)
    data["meta"]["trailerMasterOutOfServiceCount"] = sum(1 for u in trailers if u["trailerMasterStatus"] == "Out of Service")
    data["meta"]["trailerMasterMissingTypeCount"] = sum(1 for u in trailers if not u["trailerEquipmentType"])

    LEASE_JSON.write_text(json.dumps(data, indent=None, default=str), encoding="utf-8")
    print(f"Derived trailer fields for {len(trailers)} trailers -> {LEASE_JSON}")
    print(f"  Out of service: {data['meta']['trailerMasterOutOfServiceCount']}, missing type: {data['meta']['trailerMasterMissingTypeCount']}")


if __name__ == "__main__":
    main()
