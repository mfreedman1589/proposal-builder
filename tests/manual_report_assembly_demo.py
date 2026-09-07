"""Manual walking-skeleton check for report_assembly.py -- not part of the
tests/run_all.py sweep (deliberately doesn't match test_*.py). Run
directly: python tests/manual_report_assembly_demo.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import db
import market_lookup
import attribution_import as ai
import report_assembly as ra

REPO = Path(__file__).resolve().parent.parent
OUT = REPO / "tests" / "_manual_output"
OUT.mkdir(exist_ok=True)

market_lookup.install()

local_fallback = REPO / "REPORT_MASTER_v0_3.pptx"
template_path, version_id, template_warning = db.report_master_deck(
    str(local_fallback) if local_fallback.exists() else None)
if template_warning:
    print("template warning:", template_warning)
print(f"template source: {'Supabase, version ' + str(version_id) if version_id else 'local fallback'} "
     f"({template_path})")

print("\n=== MW: attribution only, no delivery file, 4-market geography ===")
mw = ai.parse_attribution_export(str(REPO / "MW attribution excel.xlsx"))
out1, warnings1 = ra.build_report_deck(
    template_path, mw, None, str(OUT / "MW_report.pptx"),
    goals_bullets=["Drive online engagement for the Fourth of July mattress sale",
                  "Build awareness across the DC/NC/VA footprint"],
    whats_next_bullets=["Expand into the top-performing markets found here",
                        "Test a second creative against MW GEOVAST"],
)
print("wrote", out1)
for w in warnings1:
    print("  WARNING:", w)

print("\n=== Cardinal: attribution + delivery, real two-file pair ===")
cardinal_attr = ai.parse_attribution_export(str(REPO / "Premion Website Attribution Cardinal.xlsx"))
cardinal_delivery = ai.parse_delivery_export(str(REPO / "Premion OTT.xlsx"))
out2, warnings2 = ra.build_report_deck(
    template_path, cardinal_attr, cardinal_delivery,
    str(OUT / "Cardinal_report.pptx"),
    goals_bullets=["Drive plumbing/HVAC service calls across the WUSA DMA"],
    whats_next_bullets=["Renew into the next flight with the top-performing segment"],
)
print("wrote", out2)
for w in warnings2:
    print("  WARNING:", w)
print("\nBoth built without raising -- walking skeleton is end to end.")
