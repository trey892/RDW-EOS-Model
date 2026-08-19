"""
Builds the RDW Revenue tab's embedded module: takes the standalone
revenue.html analytics page (a self-contained, no-backend module -- one
static 329-row data array, zero fetch/API calls, confirmed 2026-08-19), and
produces a JS string constant containing its full HTML with the one external
dependency (Chart.js, previously loaded from a CDN) inlined so it works
under the Artifact's strict CSP (no external network requests allowed).

This is a STATIC HISTORICAL SNAPSHOT, not a live-refreshed source -- the
329-row FLEET array is hand-embedded in revenue.html itself, not generated
from this pipeline's own data. Trey asked for a weekly (Tuesday 3pm) refresh
of this tab with fresh analysis -- that requires a real data pipeline behind
it (presumably the McLeod revenue PDFs) and is a separate, not-yet-built
piece of work; this script only re-packages what already exists.

Output is a JS string constant, not an HTML attribute value -- deliberately
avoids embedding via `<iframe srcdoc="...">` (which would require HTML-
attribute-escaping ~280KB of markup, expensive and error-prone) in favor of
`iframe.srcdoc = REVENUE_MODULE_HTML` set from JS, where json.dumps() safely
produces a valid JS string literal with no manual escaping.
"""
import json
import re
from pathlib import Path

REVENUE_SRC = Path(__file__).parent.parent / "rdw-fleet-site" / "worker" / "revenue.html"
CHARTJS_SRC = Path(__file__).parent / "vendor" / "chart.umd.min.js"
OUT = Path(__file__).parent / "output" / "revenue_module.json"

CDN_SCRIPT_TAG = '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.5.0/dist/chart.umd.js" integrity="sha384-iU8HYtnGQ8Cy4zl7gbNMOhsDTTKX02BTXptVP/vqAWIaTfM7isw76iyZCsjL2eVi" crossorigin="anonymous"></script>'

# The embedded HTML contains its own </script> tags (its module script, now also
# the inlined Chart.js). Once JSON-encoded into `const REVENUE_MODULE_HTML = "...";`
# in the OUTER page, those bytes are still literally "</script>" -- since JSON
# doesn't escape "/" -- which the browser's HTML tokenizer reads as closing the
# OUTER <script> block early, truncating the file mid-string (confirmed 2026-08-19:
# 3 "Invalid or unexpected token" console errors from the truncated JS). Standard
# fix: break up the byte sequence with an inert backslash (\/  -> JS unescapes
# this back to plain / at runtime) so "</script" never appears contiguously in
# the actual file text, only in the string's runtime VALUE.
SCRIPT_CLOSE_RE = re.compile(r"</script", re.IGNORECASE)


def main():
    html = REVENUE_SRC.read_text(encoding="utf-8")
    if CDN_SCRIPT_TAG not in html:
        raise ValueError("Chart.js CDN script tag not found in revenue.html -- source may have changed, update CDN_SCRIPT_TAG")

    chartjs = CHARTJS_SRC.read_text(encoding="utf-8")
    inlined_script = f"<script>\n{chartjs}\n</script>"
    html = html.replace(CDN_SCRIPT_TAG, inlined_script)

    encoded = json.dumps(html)
    encoded = SCRIPT_CLOSE_RE.sub(lambda m: m.group(0).replace("/", "\\/"), encoded)

    OUT.write_text(encoded, encoding="utf-8")
    print(f"Wrote {len(html):,} chars ({OUT.stat().st_size:,} bytes JSON-encoded) -> {OUT}")


if __name__ == "__main__":
    main()
