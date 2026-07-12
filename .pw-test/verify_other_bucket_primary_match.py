"""
Verification script: the "Colors to Produce" checklist's legacy "Other"
bucket (see Script.html's renderGroupedColorChecklist) now behaves like an
implicit non-primary Color Axis group for any process that has NOT been
formally configured with 2+ Color Axes.

Reproduces the live report: process "Packing Crysta 16 inch D/Gaddi Steel
Rim" grouped its checklist into "Fitted Frame 16 inch Crysta S/Rim" (rows
carrying the old cross-product-joined names, e.g. "Pink / BCP / Pink-White")
plus a leftover "Other" block (plain "Pink", "Purple", "Red-White", ...)
that never auto-selected anything, with no Primary radio anywhere since this
process isn't in axis mode. Checking "Pink / BCP / Pink-White" must now
auto-check AND auto-fill "Pink" (segment-match, not just hyphen-prefix; see
_colorNamesMatch), and that auto-filled "Pink" row must NOT double the lot's
saved quantity (see getCheckedColorQtys' countsTowardTotal / saveProduction).
An "Other" color with no match at all (e.g. "Blue") still counts as its own
real quantity if the operator checks it by hand.

Run: python .pw-test/verify_other_bucket_primary_match.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

COMPOSITE_COLORS = ["Pink / BCP / Pink-White", "Purple / BCP / Purple-White", "Red / BCP / Red-White"]
OTHER_COLORS = ["Blue", "Orange-White", "Pink", "Purple", "Red-White", "Sea Green-White"]
ALL_COLORS = COMPOSITE_COLORS + OTHER_COLORS

MOCK_PROCESS = {
    "processId": "PRC-1", "processName": "Packing Crysta 16 inch D/Gaddi Steel Rim", "sequence": 3,
    "lotPrefix": "PACK", "outputItemName": "16 inch Crysta D/Gaddi Steel Rim", "isFinalStage": True,
    "active": True, "processType": "Packing"
}
MOCK_COMPONENTS = [
    {"itemName": "Fitted Frame 16 inch Crysta S/Rim", "size": "16 inch", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
]
MOCK_POOL_ROWS = [
    {"outputItemName": "Fitted Frame 16 inch Crysta S/Rim", "color": c, "qty": 1000} for c in COMPOSITE_COLORS
]
MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": ALL_COLORS},
    "getProcessComponentsData": {"success": True, "data": MOCK_COMPONENTS},
    "getWarehousePoolData": {"success": True, "data": MOCK_POOL_ROWS},
    "getProcessWipData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
    # getProcessColorAxes intentionally NOT mocked - this process is not in
    # axis mode, so the client's try/catch must fall through to the legacy
    # signature-grouping branch, same as verify_pool_color_group_sum.py.
}


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_context(viewport={"width": 1500, "height": 1000}).new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.resolve().as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        page.evaluate(f"""
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
            App.State.globalItems = [];
            window.__mockResponses = {json.dumps(MOCK_API_RESPONSES)};
            window.google = {{
                script: {{
                    run: {{
                        withSuccessHandler(cb) {{
                            const runner = {{ withFailureHandler() {{ return runner; }} }};
                            Object.keys(window.__mockResponses).forEach(method => {{
                                runner[method] = (...args) => setTimeout(() => cb(window.__mockResponses[method]), 20);
                            }});
                            return runner;
                        }}
                    }}
                }}
            }};
        """)

        failures = []
        def check(cond, msg):
            print(("PASS: " if cond else "FAIL: ") + msg)
            if not cond:
                failures.append(msg)

        print("\n[Step 1] Open Create modal, select the process...")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.evaluate("document.getElementById('productionProcessId').value='PRC-1'; App.Production.handleProcessChange('PRC-1');")
        page.wait_for_timeout(800)

        group_headers = page.evaluate("""
            Array.from(document.querySelectorAll('#productionColorChecklist [data-group-master]'))
                .map(el => el.nextElementSibling ? el.nextElementSibling.textContent.trim() : '')
        """)
        check(any('Fitted Frame' in h for h in group_headers), f"pool-signature group header present (got {group_headers})")
        check(any(h == 'Other' for h in group_headers), f"'Other' group header present (got {group_headers})")
        check(not any('Primary' in h for h in group_headers), f"no axis-mode 'Primary' label anywhere - this process isn't in axis mode (got {group_headers})")

        rows_info = page.evaluate("""
            Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row')).map(r => ({
                color: r.dataset.color, group: r.dataset.group, primary: r.dataset.primary
            }))
        """)
        composite_rows = [r for r in rows_info if r['color'] in COMPOSITE_COLORS]
        other_rows = [r for r in rows_info if r['color'] in OTHER_COLORS]
        check(len(composite_rows) == 3 and all(r['primary'] == 'true' for r in composite_rows),
              f"all 3 pool-signature rows tagged data-primary=true (got {composite_rows})")
        check(len(other_rows) == 6 and all(r['primary'] == 'false' and r['group'] == 'other' for r in other_rows),
              f"all 6 'Other' rows tagged data-primary=false, group=other (got {other_rows})")

        print("\n[Step 2] Check 'Pink / BCP / Pink-White'=12 - matching 'Pink' (Other) must auto-check+auto-fill...")
        page.locator("#productionColorChecklist .production-color-row[data-color='Pink / BCP / Pink-White'] .production-color-check").click()
        page.wait_for_timeout(150)
        page.locator("#productionColorChecklist .production-color-row[data-color='Pink / BCP / Pink-White'] .production-color-qty").fill("12")
        page.wait_for_timeout(150)

        pink_other = page.locator("#productionColorChecklist .production-color-row[data-color='Pink']")
        check(pink_other.locator(".production-color-check").is_checked(), "'Pink' (Other) auto-CHECKED itself, no manual click")
        check(pink_other.locator(".production-color-qty").input_value() == "12", "'Pink' (Other) auto-filled to 12, matching the composite row's own qty")

        print("\n[Step 3] Check 'Red / BCP / Red-White'=8 - matching 'Red-White' (Other, a SUFFIX-position segment match) must auto-check+auto-fill...")
        page.locator("#productionColorChecklist .production-color-row[data-color='Red / BCP / Red-White'] .production-color-check").click()
        page.wait_for_timeout(150)
        page.locator("#productionColorChecklist .production-color-row[data-color='Red / BCP / Red-White'] .production-color-qty").fill("8")
        page.wait_for_timeout(150)

        redwhite_other = page.locator("#productionColorChecklist .production-color-row[data-color='Red-White']")
        check(redwhite_other.locator(".production-color-check").is_checked(), "'Red-White' (Other) auto-CHECKED itself via suffix-segment match")
        check(redwhite_other.locator(".production-color-qty").input_value() == "8", "'Red-White' (Other) auto-filled to 8")

        print("\n[Step 4] Manually check an UNMATCHED 'Other' color ('Blue')=5 - no counterpart to auto-fill from...")
        blue_other = page.locator("#productionColorChecklist .production-color-row[data-color='Blue']")
        blue_other.locator(".production-color-check").click()
        page.wait_for_timeout(150)
        blue_other.locator(".production-color-qty").fill("5")
        page.wait_for_timeout(150)

        print("\n[Step 5] Inspect getCheckedColorQtys() countsTowardTotal flags...")
        qtys = page.evaluate("App.Production.getCheckedColorQtys()")
        by_color = {q['color']: q for q in qtys}
        print(f"  {json.dumps(qtys, indent=2)}")
        check(by_color.get('Pink / BCP / Pink-White', {}).get('countsTowardTotal') is True, "composite 'Pink / BCP / Pink-White' counts toward total")
        check(by_color.get('Red / BCP / Red-White', {}).get('countsTowardTotal') is True, "composite 'Red / BCP / Red-White' counts toward total")
        check(by_color.get('Pink', {}).get('countsTowardTotal') is False, "matched 'Pink' (Other) does NOT count toward total (would double-count)")
        check(by_color.get('Red-White', {}).get('countsTowardTotal') is False, "matched 'Red-White' (Other) does NOT count toward total (would double-count)")
        check(by_color.get('Blue', {}).get('countsTowardTotal') is True, "unmatched 'Blue' (Other) DOES count toward total (a genuinely real quantity)")

        total = page.evaluate("App.Production._currentLotTotalQty()")
        check(total == 25, f"lot total is 12+8+5=25, not 12+8+12+8+5=45 (got {total})")

        if console_errors:
            print("\n⚠️  Console/page errors:")
            for e in console_errors:
                print(f"    {e}")
            failures.append("console errors present")

        print("\n" + ("ALL TESTS PASSED" if not failures else f"{len(failures)} TEST(S) FAILED"))
        browser.close()
        return len(failures) == 0


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
