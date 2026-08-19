"""Shared exclusion lists applied when building dashboard data, kept in one place
so every build_*.py script that groups or filters by terminal stays in sync.
"""

# Fleet/terminal sub-codes Trey asked to exclude from this dashboard entirely (2026-08-19).
# These are Dim_Equipment Terminal/FleetCode values, not physical terminals -- matched on
# either column since the two are populated inconsistently for the same asset.
EXCLUDED_TERMINALS = {
    "ATLO", "BHMO", "CPO", "HOUO", "MAU", "MAUO", "MOBO",
    "MTP", "MTPB", "MTPO", "ROO",
}


def is_excluded_terminal(*values):
    """True if any of the given Terminal/FleetCode values is on the exclusion list."""
    return any(str(v or "").strip().upper() in EXCLUDED_TERMINALS for v in values)
