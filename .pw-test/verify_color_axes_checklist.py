"""
Verification script: Production Lot form's "Colors to Produce" checklist in
Color Axis mode (2+ independent axes, e.g. Rim Color + Mudguard Color).

Reproduces the reported UX/bug fix end-to-end in the real client code
(renderGroupedColorChecklist / handleColorCheckToggle / _syncMatchingNonPrimaryRows
/ onColorQtyChanged / getCheckedColorQtys in Script.html): instead of one
flat/cross-multiplied list, the checklist must render one independent group
per axis; checking a primary axis row must auto-check AND auto-fill its
matching non-primary row (e.g. a "Red" mudguard following a "Red-White"
frame) with THAT specific color's own qty, not the grand total across every
checked primary color — still plain-editable afterward, and no longer
auto-tracked once the operator edits it by hand. A non-primary row with NO
matching primary color (e.g. a "Black" Fitting used regardless of frame
color) must instead keep tracking the running grand total live, the same as
a real Common Component would (_refreshAutoSyncedFallbackRows) — found via
screenshot review of the live app, where it was freezing at whatever total
existed the moment it was checked instead of following later edits. The raw
client-side colorBreakdown payload (what saveProduction sums server-side —
see test_color_axes.js for the server-side half of this fix) must carry
every row untouched, since the quantity correction itself happens
server-side.

Run: python .pw-test/verify_color_axes_checklist.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

MOCK_PROCESS = {
    "processId": "PRC-1", "processName": "Fitted Frame Assembly", "sequence": 1,
    "lotPrefix": "FFA", "outputItemName": "Fitted Frame", "isFinalStage": False,
    "active": True, "processType": "General"
}

MOCK_COMPONENTS = [
    {"itemName": "Assembly Screws", "size": "", "sourceType": "ITEM", "qtyPerUnit": 4, "colorGroup": "COMMON"},
    {"itemName": "Painted Rim - Blue-White", "size": "", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "Blue-White"},
    {"itemName": "Painted Rim - Red-White", "size": "", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "Red-White"},
    {"itemName": "Mudguard - Blue", "size": "", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "Blue"},
    {"itemName": "Mudguard - Red", "size": "", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "Red"},
    {"itemName": "Fitted Rim - Black", "size": "", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "Black"},
]
MOCK_ITEMS = [{"name": c["itemName"], "size": c["size"]} for c in MOCK_COMPONENTS]

MOCK_AXES = {
    "axes": [
        {"key": "tag:rim color", "label": "Rim Color", "colors": ["Blue-White", "Red-White"], "source": "tag"},
        {"key": "tag:mudguard color", "label": "Mudguard Color", "colors": ["Blue", "Red"], "source": "tag"},
        {"key": "tag:fitting", "label": "Fitting", "colors": ["Black"], "source": "tag"},
    ],
    "primaryColorAxis": "Rim Color",
    "primaryAxisKey": "tag:rim color",
}

MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": ["Black", "Blue", "Blue-White", "Red", "Red-White"]},
    "getProcessColorAxes": {"success": True, "data": MOCK_AXES},
    "getProcessComponentsData": {"success": True, "data": MOCK_COMPONENTS},
    "getWarehousePoolData": {"success": True, "data": []},
    "getProcessWipData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
}


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context()
        page = ctx.new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        page.evaluate(f"""
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
            App.State.globalItems = {json.dumps(MOCK_ITEMS)};
            window.__mockResponses = {json.dumps(MOCK_API_RESPONSES)};
            window.google = {{
                script: {{
                    run: {{
                        withSuccessHandler(cb) {{
                            const runner = {{ withFailureHandler() {{ return runner; }} }};
                            Object.keys(window.__mockResponses).forEach(method => {{
                                runner[method] = (...args) => setTimeout(() => cb(window.__mockResponses[method]), 30);
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

        print("[Step 1] Open Create modal, select the multi-axis process...")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.evaluate("""
            document.getElementById('productionProcessId').value = 'PRC-1';
            App.Production.handleProcessChange('PRC-1');
        """)
        page.wait_for_timeout(800)

        group_headers = page.evaluate("""
            Array.from(document.querySelectorAll('#productionColorChecklist [data-group-master]'))
                .map(el => el.nextElementSibling ? el.nextElementSibling.textContent.trim() : '')
        """)
        check(len(group_headers) == 3, f"exactly 3 group headers rendered (got {group_headers})")
        check(any('Rim Color' in h for h in group_headers), f"'Rim Color' group header present (got {group_headers})")
        check(any('Mudguard Color' in h for h in group_headers), f"'Mudguard Color' group header present (got {group_headers})")
        check(any('Fitting' in h for h in group_headers), f"'Fitting' group header present (got {group_headers})")

        rows_info = page.evaluate("""
            Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row')).map(r => ({
                color: r.dataset.color, primary: r.dataset.primary
            }))
        """)
        check(len(rows_info) == 5, f"exactly 5 checkbox rows total, no composite combo strings (got {rows_info})")
        check(not any('/' in (r['color'] or '') for r in rows_info), f"no row uses a composite 'X / Y' color string (got {rows_info})")
        rim_primary = all(r['primary'] == 'true' for r in rows_info if r['color'] in ('Blue-White', 'Red-White'))
        mud_nonprimary = all(r['primary'] == 'false' for r in rows_info if r['color'] in ('Blue', 'Red'))
        check(rim_primary, f"Rim Color rows tagged data-primary=true (got {rows_info})")
        check(mud_nonprimary, f"Mudguard Color rows tagged data-primary=false (got {rows_info})")

        print("\n[Step 2] Check TWO primary colors (Red-White=10, Blue-White=5) - matching mudguards must auto-check themselves...")
        # Reproduces the reported calculation mistake AND the follow-up ask
        # ("mudguard colors should match with frame colors" / "for any
        # primary process I choose, for any secondary items I choose"):
        # checking a primary row must both (a) auto-CHECK its matching
        # non-primary row with no separate manual click, and (b) fill that
        # row with ITS OWN corresponding qty (e.g. Red -> Red-White's 10),
        # never the grand total across every checked primary color (10+5=15).
        page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Red-White"] .production-color-check').click()
        """)
        page.wait_for_timeout(200)
        page.locator("#productionColorChecklist .production-color-row[data-color='Red-White'] .production-color-qty").fill("10")
        page.wait_for_timeout(200)
        page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Blue-White"] .production-color-check').click()
        """)
        page.wait_for_timeout(200)
        page.locator("#productionColorChecklist .production-color-row[data-color='Blue-White'] .production-color-qty").fill("5")
        page.wait_for_timeout(300)

        red_mudguard = page.evaluate("""
            (() => {
                const row = document.querySelector('#productionColorChecklist .production-color-row[data-color="Red"]');
                return { checked: row.querySelector('.production-color-check').checked, qty: row.querySelector('.production-color-qty').value };
            })()
        """)
        blue_mudguard = page.evaluate("""
            (() => {
                const row = document.querySelector('#productionColorChecklist .production-color-row[data-color="Blue"]');
                return { checked: row.querySelector('.production-color-check').checked, qty: row.querySelector('.production-color-qty').value };
            })()
        """)
        check(red_mudguard['checked'] is True, f"Mudguard 'Red' auto-CHECKED itself when Red-White was checked, no manual click (got {red_mudguard})")
        check(blue_mudguard['checked'] is True, f"Mudguard 'Blue' auto-CHECKED itself when Blue-White was checked, no manual click (got {blue_mudguard})")
        check(red_mudguard['qty'] == '10', f"Mudguard 'Red' auto-fills to its MATCHING Red-White qty (10), not the 15 grand total (got {red_mudguard})")
        check(blue_mudguard['qty'] == '5', f"Mudguard 'Blue' auto-fills to its MATCHING Blue-White qty (5), not the 15 grand total (got {blue_mudguard})")

        print("\n[Step 3] Auto-filled value stays editable (not locked/disabled)...")
        mudguard_disabled = page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Red"] .production-color-qty').disabled
        """)
        check(mudguard_disabled is False, "Mudguard qty field is enabled/editable, not locked")

        page.locator("#productionColorChecklist .production-color-row[data-color='Red'] .production-color-qty").fill("3")
        page.wait_for_timeout(200)
        edited_value = page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Red"] .production-color-qty').value
        """)
        check(edited_value == '3', f"operator can override the auto-filled value (got '{edited_value}')")

        print("\n[Step 4] getCheckedColorQtys() carries all four rows for the server-side primary-only sum to consume...")
        checked = page.evaluate("JSON.stringify(App.Production.getCheckedColorQtys())")
        checked_parsed = json.loads(checked)
        by_color = {c['color']: c['qty'] for c in checked_parsed}
        check(by_color.get('Red-White') == 10 and by_color.get('Blue-White') == 5 and by_color.get('Red') == 3 and by_color.get('Blue') == 5,
              f"all four rows present with their own quantities, unmerged (got {checked_parsed})")

        # Common Components suggested qty must scale off the PRIMARY total
        # across BOTH checked primary colors (10+5=15), not the non-primary
        # rows (3+5=8) - this is the client-side half of the same
        # double-count fix (see refreshCommonSuggestedQty / saveProduction).
        screws_qty = page.evaluate("""
            (() => {
                const row = Array.from(document.querySelectorAll('#productionComponentsBody tr'))
                    .find(r => (r.querySelector('.prod-comp-item-select')?.selectedOptions[0]?.textContent || '').includes('Assembly Screws'));
                return row ? row.querySelector('.prod-comp-qty')?.value : null;
            })()
        """)
        check(screws_qty == '60', f"Common Component 'Assembly Screws' (qtyPerUnit 4) suggests 60 = 4x15 (primary total 10+5), not scaled off non-primary rows (got '{screws_qty}')")

        print("\n[Step 5] Unchecking a primary row auto-unchecks its matching non-primary row...")
        page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Red-White"] .production-color-check').click()
        """)
        page.wait_for_timeout(200)
        red_mudguard_checked = page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Red"] .production-color-check').checked
        """)
        check(red_mudguard_checked is False, "unchecking Red-White auto-unchecks the matching Red mudguard")

        print("\n[Step 6] Checking a primary row auto-CHECKS the matching non-primary row (no manual click on it)...")
        page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Red-White"] .production-color-check').click()
        """)
        page.wait_for_timeout(200)
        page.locator("#productionColorChecklist .production-color-row[data-color='Red-White'] .production-color-qty").fill("7")
        page.wait_for_timeout(200)
        red_state = page.evaluate("""
            (() => {
                const row = document.querySelector('#productionColorChecklist .production-color-row[data-color="Red"]');
                return { checked: row.querySelector('.production-color-check').checked, qty: row.querySelector('.production-color-qty').value };
            })()
        """)
        check(red_state['checked'] is True, f"Red mudguard auto-CHECKED itself when Red-White was checked, no manual click needed (got {red_state})")
        check(red_state['qty'] == '7', f"Red mudguard's qty follows Red-White's own qty (7) (got {red_state})")

        print("\n[Step 7] Still-auto-tracked non-primary qty keeps following the primary row's edits...")
        page.locator("#productionColorChecklist .production-color-row[data-color='Red-White'] .production-color-qty").fill("9")
        page.wait_for_timeout(200)
        red_qty_after_edit = page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Red"] .production-color-qty').value
        """)
        check(red_qty_after_edit == '9', f"Red mudguard follows Red-White's qty change to 9 while still auto-tracked (got '{red_qty_after_edit}')")

        print("\n[Step 8] Manually editing the matched row's qty stops it from auto-tracking any further...")
        page.locator("#productionColorChecklist .production-color-row[data-color='Red'] .production-color-qty").fill("2")
        page.wait_for_timeout(200)
        page.locator("#productionColorChecklist .production-color-row[data-color='Red-White'] .production-color-qty").fill("15")
        page.wait_for_timeout(200)
        red_qty_final = page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Red"] .production-color-qty').value
        """)
        check(red_qty_final == '2', f"Red mudguard's manually-edited qty (2) is preserved even after Red-White changes again to 15 (got '{red_qty_final}')")

        print("\n[Step 9] A non-primary row with NO matching primary color (Black Fitting) defaults to the running grand total...")
        # State here: Red-White=15, Blue-White=5 -> primary total = 20.
        page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Black"] .production-color-check').click()
        """)
        page.wait_for_timeout(200)
        black_qty = page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Black"] .production-color-qty').value
        """)
        check(black_qty == '20', f"Black Fitting defaults to the grand total across every checked primary color (15+5=20) (got '{black_qty}')")

        print("\n[Step 10] Editing a primary row's qty keeps the still-tracked grand-total row (Black) in sync...")
        page.locator("#productionColorChecklist .production-color-row[data-color='Blue-White'] .production-color-qty").fill("8")
        page.wait_for_timeout(200)
        black_qty2 = page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Black"] .production-color-qty').value
        """)
        check(black_qty2 == '23', f"Black Fitting follows the new grand total (15+8=23), not frozen at the old 20 (got '{black_qty2}') - this is the exact live-app calculation mistake found via screenshot review")

        print("\n[Step 11] Unchecking a DIFFERENT primary color also refreshes the still-tracked Black row...")
        page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Red-White"] .production-color-check').click()
        """)
        page.wait_for_timeout(200)
        black_qty3 = page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Black"] .production-color-qty').value
        """)
        check(black_qty3 == '8', f"Black Fitting drops to just Blue-White's 8 once Red-White is unchecked (got '{black_qty3}')")

        print("\n[Step 12] Manually editing Black stops it from auto-tracking any further...")
        page.locator("#productionColorChecklist .production-color-row[data-color='Black'] .production-color-qty").fill("99")
        page.wait_for_timeout(200)
        page.locator("#productionColorChecklist .production-color-row[data-color='Blue-White'] .production-color-qty").fill("50")
        page.wait_for_timeout(200)
        black_qty_final = page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Black"] .production-color-qty').value
        """)
        check(black_qty_final == '99', f"Black's manually-edited value (99) is preserved even after Blue-White changes to 50 (got '{black_qty_final}')")

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
