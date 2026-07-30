"""
Verification: "Refresh from Items Master" re-renders an ALREADY-OPEN
Production Sheet dialog with the new narration/unit, instead of only taking
effect the next time the sheet is reopened.

Reported as "narration in production sheet isn't changing after refreshing".
The resolution itself was fine -- the Production Sheet resolves narration and
Base Unit live against App.State.globalItems at render time
(_resolveDisplayNarration / _resolveDisplayUnit), and both were correct once
that cache was reloaded. What was missing is that the refresh reloaded the
data but never repainted the dialog already on screen, so the operator kept
looking at the pre-refresh render.

Fixed by capturing the open sheet's stable rowIdx BEFORE the reload
(_openProductionSheetRowKey -- currentProductionSheet.idx is a position in
globalProduction, which loadData() replaces wholesale) and re-running
_populateProductionSheetData for it afterwards
(_rerenderOpenProductionSheet).

Covered:
  1. With the sheet open, a refresh updates the narration cell in place.
  2. ...and the unit cell too.
  3. The dialog stays open (the refresh must not close it out from under the
     operator).
  4. The re-render targets the SAME lot even when the reload reorders
     globalProduction, which a naive index-based lookup would get wrong.
  5. With no sheet open, the refresh completes normally and opens nothing.

Run: python .pw-test/verify_refresh_rerenders_open_production_sheet.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

PROCESSES = [{"processId": "PRC-1", "processName": "Packing", "sequence": 1, "lotPrefix": "PACK",
              "outputItemName": "Packed Cycle 16 inch", "isFinalStage": True, "active": True,
              "processType": "General"}]

ITEMS_V1 = [{"name": "Carton Box", "size": "", "narration": "OLD narration", "remarks": "", "baseUnit": "Pcs"}]
ITEMS_V2 = [{"name": "Carton Box", "size": "", "narration": "NEW narration", "remarks": "", "baseUnit": "Set"}]


def lot(row_idx, lot_no):
    return {
        "rowIdx": row_idx, "date": "01/07/2026", "dateRaw": "2026-07-01", "lotNumber": lot_no,
        "processId": "PRC-1", "outputItemName": "Packed Cycle 16 inch", "productId": "PT-1",
        "productName": "Model A", "qty": 10, "assignedBy": "A", "assignedTo": "B",
        "status": "Pending", "color": "", "colorBreakdown": [],
        "componentsConsumed": [
            {"itemName": "Carton Box", "size": "", "narration": "STORED old narration",
             "colorGroup": "COMMON", "sourceType": "ITEM", "qty": 10, "unit": ""}
        ],
        "customComponents": []
    }


# Two lots, so step 4 can reverse their order on reload.
LOTS_V1 = [lot(1, "PACK-1"), lot(2, "PACK-2")]
LOTS_V2 = [lot(2, "PACK-2"), lot(1, "PACK-1")]  # deliberately reordered


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_context().new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        page.evaluate(f"""
            window.__items = {json.dumps(ITEMS_V1)};
            window.__lots = {json.dumps(LOTS_V1)};
            window.__mock = {{
                "getProcessData": {{"success": true, "data": {json.dumps(PROCESSES)}}},
                "getColors": {{"success": true, "data": []}},
                "getStockData": {{"success": true, "data": []}},
                "getIssuedStockData": {{"success": true, "data": []}},
                "refreshProductionComponentsFromItemsMaster": {{"success": true, "message": "Refreshed 2 field(s).", "data": {{"fieldsUpdated": 2}}}}
            }};
            window.google = {{ script: {{ run: {{
                withSuccessHandler(cb) {{
                    const r = {{ withFailureHandler() {{ return r; }} }};
                    Object.keys(window.__mock).forEach(m => {{ r[m] = () => setTimeout(() => cb(window.__mock[m]), 15); }});
                    r.getItemsData = () => setTimeout(() => cb({{success: true, data: window.__items}}), 15);
                    r.getProductionData = () => setTimeout(() => cb({{success: true, data: window.__lots}}), 15);
                    return r;
                }}
            }} }} }};
            App.State.globalItems = [];
        """)

        failures = []

        def check(cond, msg):
            print(("PASS: " if cond else "FAIL: ") + msg)
            if not cond:
                failures.append(msg)

        def sheet_cells():
            return page.evaluate("""() => {
                const row = document.querySelector('#productionSheetCommonBody tr');
                if (!row) return {narration: '(no row)', unit: '(no row)'};
                return {
                    narration: row.querySelector('.prod-sheet-narration')?.value ?? '(none)',
                    unit: row.querySelector('.prod-sheet-unit')?.textContent.trim() ?? '(none)'
                };
            }""")

        print("[Step 1] Load Production and open PACK-1's Production Sheet...")
        page.evaluate("App.Navigation.showTab('productionTab')")
        page.wait_for_timeout(900)

        # Open the sheet for PACK-1 specifically (rowIdx 1, currently index 0).
        page.evaluate("""() => {
            const idx = App.State.globalProduction.findIndex(p => String(p.rowIdx) === '1');
            App.Production.viewProductionSheet(idx);
        }""")
        page.locator("#productionSheetModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(400)

        before = sheet_cells()
        print("  sheet before:", before)
        check(before["narration"] == "OLD narration",
              f"sheet initially shows the OLD Items Master narration (got '{before['narration']}')")
        check(before["unit"] == "Pcs", f"sheet initially shows the OLD unit (got '{before['unit']}')")

        print("\n[Step 2] Operator edits Items Master, then hits Refresh with the sheet still open...")
        page.evaluate("() => { window.__items = " + json.dumps(ITEMS_V2) + "; window.__lots = " + json.dumps(LOTS_V2) + "; }")
        page.evaluate("() => { App.Utils.confirmAction = (msg, cb) => cb(); }")
        page.evaluate("() => App.Production.refreshFromItemsMaster()")
        page.wait_for_timeout(1600)

        after = sheet_cells()
        print("  sheet after:", after)
        check(after["narration"] == "NEW narration",
              f"open sheet re-rendered with the NEW narration (got '{after['narration']}')")
        check(after["unit"] == "Set",
              f"open sheet re-rendered with the NEW unit (got '{after['unit']}')")

        still_open = page.evaluate("""() => {
            const m = document.getElementById('productionSheetModal');
            return !!m && m.classList.contains('show');
        }""")
        check(still_open, "the Production Sheet dialog stayed open through the refresh")

        print("\n[Step 3] The re-render followed the SAME lot despite the reload reordering the list...")
        shown_lot = page.evaluate("() => document.getElementById('prodSheetProductId')?.innerText || ''")
        current_idx = page.evaluate("() => App.State.currentProductionSheet?.idx")
        lot_at_idx = page.evaluate("() => (App.State.globalProduction[App.State.currentProductionSheet?.idx]||{}).lotNumber")
        print(f"  currentProductionSheet.idx={current_idx} -> lot {lot_at_idx}")
        check(lot_at_idx == "PACK-1",
              f"still pointing at PACK-1 after globalProduction was reordered (got '{lot_at_idx}')")

        print("\n[Step 4] With no sheet open, a refresh opens nothing...")
        page.evaluate("""() => {
            const m = document.getElementById('productionSheetModal');
            if (m && window.bootstrap) bootstrap.Modal.getOrCreateInstance(m).hide();
        }""")
        page.wait_for_timeout(500)
        page.evaluate("() => App.Production.refreshFromItemsMaster()")
        page.wait_for_timeout(1400)
        reopened = page.evaluate("""() => {
            const m = document.getElementById('productionSheetModal');
            return !!m && m.classList.contains('show');
        }""")
        check(not reopened, "refresh with the dialog closed does not pop it open")

        check(len(console_errors) == 0, f"No uncaught page errors (got {console_errors})")

        print("\n" + "=" * 60)
        if failures:
            print(f"FAILURES: {len(failures)}")
            for f in failures:
                print("  -", f)
        else:
            print("ALL CHECKS PASSED")
        print("=" * 60)

        browser.close()
        return len(failures) == 0


if __name__ == "__main__":
    sys.exit(0 if run() else 1)
