"""
Verifies two Edit/Create Production Lot form bugs reported 2026-07-22:

Bug 1 - Saving a Production lot (create or edit) never refreshed the
Warehouse Pool sub-tab (Script_Stock.html's App.Stock.loadWarehousePoolData),
so its color-combination quantities went stale until the operator manually
revisited the Stock tab. Same class of gap already fixed for Process saves
(see bug_warehouse_pool_search_and_colorrefresh_2026_07_15 in project
memory) - Production's save handler now does the same guarded refresh.

Bug 2 - Clicking "+ Add Component" (App.Production.addComponentRow() with no
argument) always left the new row's Qty blank, even with a lot Qty already
entered - the operator had to notice and fill it in by hand. Now it assumes
a qtyPerUnit of 1 (same as a qtyPerUnit=1 recipe row) so it pre-fills at the
current lot qty/color total immediately and keeps auto-tracking it via the
same _applyQtyPerUnit mechanism recipe-derived rows use.

Run: python .pw-test/verify_production_addcomponent_and_poolrefresh.py
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
    "processId": "PRC-1", "processName": "Frame Painting", "sequence": 1,
    "lotPrefix": "FP", "outputItemName": "Painted Frame", "isFinalStage": False,
    "active": True, "processType": "General"
}
MOCK_COMPONENTS = [
    {"itemName": "Brush", "size": "", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "COMMON"},
]

failures = 0
def check(cond, msg):
    global failures
    if cond:
        print(f"  PASS: {msg}")
    else:
        failures += 1
        print(f"  FAIL: {msg}")


def setup_mocks(page, extra_responses=None, save_response=None):
    responses = {
        "getProcessColorGroups": {"success": True, "data": []},
        "getProcessComponentsData": {"success": True, "data": MOCK_COMPONENTS},
        "getWarehousePoolData": {"success": True, "data": []},
        "getProcessWipData": {"success": True, "data": []},
        "getStockData": {"success": True, "data": []},
        "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
    }
    if extra_responses:
        responses.update(extra_responses)
    page.evaluate(f"""
        App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
        App.State.globalItems = [];
        App.State.globalColors = [];
        App.State.globalContractors = [{{name: 'Sanjay'}}];
        App.State.globalProduction = [];
        window.__mockResponses = {json.dumps(responses)};
        window.__saveProductionResponse = {json.dumps(save_response or {"success": True, "message": "Saved"})};
        window.google = {{
            script: {{ run: {{
                withSuccessHandler(cb) {{
                    const runner = {{ withFailureHandler() {{ return runner; }} }};
                    Object.keys(window.__mockResponses).forEach(method => {{
                        runner[method] = (...args) => setTimeout(() => cb(window.__mockResponses[method]), 20);
                    }});
                    runner.saveProduction = (...args) => setTimeout(() => cb(window.__saveProductionResponse), 20);
                    return runner;
                }}
            }} }}
        }};
    """)


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context()
        page = ctx.new_page()
        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(500)
        setup_mocks(page)

        print("[Bug 2a] Manual '+ Add Component' pre-fills qty from lot Qty (single-qty mode)")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(400)

        page.fill("#productionQty", "5")
        page.evaluate("App.Production.refreshSuggestedComponentQty()")
        page.click("#productionAddCommonBtn")
        page.wait_for_timeout(100)
        newRowQty = page.evaluate("""
            (() => {
                const rows = document.querySelectorAll('#productionComponentsBody tr');
                const last = rows[rows.length - 1];
                return last ? last.querySelector('.prod-comp-qty').value : null;
            })()
        """)
        check(newRowQty == "5", f"new manual row's qty pre-fills to lot Qty (5) x qtyPerUnit (1) -> got {newRowQty!r}")

        print("\n[Bug 2b] Manual row keeps tracking lot Qty afterward (data-qty-per-unit=1)")
        page.fill("#productionQty", "12")
        page.evaluate("App.Production.refreshSuggestedComponentQty()")
        page.wait_for_timeout(50)
        updatedQty = page.evaluate("""
            (() => {
                const rows = document.querySelectorAll('#productionComponentsBody tr');
                const last = rows[rows.length - 1];
                return last ? last.querySelector('.prod-comp-qty').value : null;
            })()
        """)
        check(updatedQty == "12", f"manual row's qty follows lot Qty changes (12) -> got {updatedQty!r}")

        print("\n[Bug 2c] Recipe-derived row is unaffected (still tracks independently)")
        recipeRowQty = page.evaluate("""
            (() => {
                const rows = document.querySelectorAll('#productionComponentsBody tr');
                const first = rows[0];
                return first ? first.querySelector('.prod-comp-qty').value : null;
            })()
        """)
        check(recipeRowQty == "12", f"recipe row (qtyPerUnit=1) also shows 12 -> got {recipeRowQty!r}")

        print("\n[Bug 2d] A saved lot's own recorded component (no qtyPerUnit) is NOT auto-recomputed")
        page.evaluate("""
            App.Production.addComponentRow({ itemName: 'Legacy Part', size: '', sourceType: 'ITEM', qty: 3 });
        """)
        page.fill("#productionQty", "99")
        page.evaluate("App.Production.refreshSuggestedComponentQty()")
        page.wait_for_timeout(50)
        legacyQty = page.evaluate("""
            (() => {
                const rows = document.querySelectorAll('#productionComponentsBody tr');
                const last = rows[rows.length - 1];
                return last ? last.querySelector('.prod-comp-qty').value : null;
            })()
        """)
        check(legacyQty == "3", f"a saved/legacy row with no qtyPerUnit keeps its own recorded qty (3), untouched by Qty changes -> got {legacyQty!r}")

        # Close this modal before opening a fresh one for the save-refresh checks
        page.evaluate("bootstrap.Modal.getOrCreateInstance(document.getElementById('editProductionModal')).hide()")
        page.wait_for_timeout(300)

        print("\n[Bug 1a] Saving a Production lot refreshes Warehouse Pool when its sub-tab IS visible")
        setup_mocks(page)
        page.evaluate("""
            const poolTab = document.createElement('div');
            poolTab.id = 'warehousePoolSubTab';
            poolTab.style.display = 'block';
            document.body.appendChild(poolTab);
            window.__poolRefreshCalls = 0;
            App.Stock = App.Stock || {};
            App.Stock.loadWarehousePoolData = () => { window.__poolRefreshCalls++; };
        """)
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(400)
        page.fill("#productionDate", "2026-07-22")
        page.fill("#productionQty", "5")
        page.evaluate("""
            document.getElementById('productionAssignedTo').innerHTML = '<option value="Sanjay">Sanjay</option>';
            document.getElementById('productionAssignedTo').value = 'Sanjay';
        """)
        page.evaluate("document.getElementById('productionForm').dispatchEvent(new Event('submit', {cancelable: true, bubbles: true}))")
        page.wait_for_timeout(300)
        poolRefreshCalls = page.evaluate("window.__poolRefreshCalls")
        check(poolRefreshCalls == 1, f"App.Stock.loadWarehousePoolData() called exactly once after save with pool sub-tab visible -> got {poolRefreshCalls}")

        print("\n[Bug 1b] Saving a Production lot does NOT touch Warehouse Pool when its sub-tab is hidden")
        page.evaluate("document.getElementById('warehousePoolSubTab').style.display = 'none'")
        page.evaluate("window.__poolRefreshCalls = 0")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(400)
        page.fill("#productionDate", "2026-07-22")
        page.fill("#productionQty", "5")
        page.evaluate("""
            document.getElementById('productionAssignedTo').innerHTML = '<option value="Sanjay">Sanjay</option>';
            document.getElementById('productionAssignedTo').value = 'Sanjay';
        """)
        page.evaluate("document.getElementById('productionForm').dispatchEvent(new Event('submit', {cancelable: true, bubbles: true}))")
        page.wait_for_timeout(300)
        poolRefreshCallsHidden = page.evaluate("window.__poolRefreshCalls")
        check(poolRefreshCallsHidden == 0, f"App.Stock.loadWarehousePoolData() NOT called when pool sub-tab is hidden -> got {poolRefreshCallsHidden}")

        if console_errors:
            print("\nConsole/page errors:")
            for e in console_errors:
                print(f"    {e}")
            failures_local = len(console_errors)
            global failures
            failures += failures_local

        browser.close()


if __name__ == "__main__":
    run()
    print(f"\n{'ALL PASS' if failures == 0 else str(failures) + ' FAILURE(S)'}")
    sys.exit(0 if failures == 0 else 1)
