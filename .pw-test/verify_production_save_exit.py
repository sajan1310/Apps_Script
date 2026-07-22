"""
Verifies the new Save/Exit behavior on the Production Lot Edit modal:
  1. Edit mode shows Exit (not Cancel), submit button reads "Save".
  2. Clicking Save (form submit) on success keeps the modal OPEN and
     re-populates it (via openEditModal) instead of closing.
  3. Exit with no unsaved changes closes immediately.
  4. Exit with unsaved changes shows the confirm dialog; confirming closes.
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(r"C:\Users\erkar\my-app-script-project\dist\index.html")
TIMEOUT = 8000

MOCK_PROCESS = {
    "processId": "PRC-1", "processName": "Packing", "sequence": 1,
    "lotPrefix": "PK", "outputItemName": "Packed Bicycle", "isFinalStage": True,
    "active": True, "processType": "General"
}
MOCK_LOT = {
    "rowIdx": 5, "lotNumber": "LOT-PK-0001", "processId": "PRC-1", "qty": 10,
    "date": "01/07/2026", "assignedBy": "Admin", "assignedTo": "Contractor A",
    "status": "Pending", "remarks": "orig remarks", "productId": "", "productName": "",
    "color": "", "colorBreakdown": [], "componentsConsumed": []
}
MOCK_LOT_2 = dict(MOCK_LOT, rowIdx=6, lotNumber="LOT-PK-0002")

MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": []},
    "getProcessComponentsData": {"success": True, "data": []},
    "getWarehousePoolData": {"success": True, "data": []},
    "getProcessWipData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
    "saveProduction": {"success": True, "data": {"lotNumber": "LOT-PK-0001"}, "message": "Production log updated successfully."},
    "getProductionData": {"success": True, "data": [MOCK_LOT, MOCK_LOT_2]},
}


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1400, "height": 900})
        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)
        page.evaluate(f"""
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
            App.State.globalProduction = [{json.dumps(MOCK_LOT)}, {json.dumps(MOCK_LOT_2)}];
            App.State.filteredProduction = App.State.globalProduction;
            App.State.globalItems = [];
            App.State.globalContractors = [{{"name": "Contractor A"}}];
            window.__mockResponses = {json.dumps(MOCK_API_RESPONSES)};
            window.google = {{
                script: {{ run: {{
                    withSuccessHandler(cb) {{
                        const runner = {{ withFailureHandler() {{ return runner; }} }};
                        Object.keys(window.__mockResponses).forEach(method => {{
                            runner[method] = (...args) => setTimeout(() => cb(window.__mockResponses[method]), 30);
                        }});
                        return runner;
                    }}
                }} }}
            }};
        """)

        print("[1] Open Edit modal for lot rowIdx=5...")
        page.evaluate("App.Production.openEditModal(0)")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(300)

        state = page.evaluate("""() => ({
            cancelVisible: document.getElementById('productionCancelBtn').style.display !== 'none',
            exitVisible: document.getElementById('productionExitBtn').style.display !== 'none',
            submitText: document.getElementById('productionSubmitBtn').innerText
        })""")
        print(f"  {state}")
        ok1 = (not state['cancelVisible']) and state['exitVisible'] and state['submitText'].strip() == 'Save'
        print("  " + ("✅ PASS" if ok1 else "❌ FAIL"))

        print("[2] Edit Remarks, click Save (submit) -> modal should STAY open, remarks field repopulated...")
        page.fill("#productionRemarks", "edited remarks")
        page.click("#productionSubmitBtn")
        page.wait_for_timeout(500)
        modal_visible = page.locator("#editProductionModal").is_visible()
        row_idx_val = page.evaluate("document.getElementById('productionRowIdx').value")
        print(f"  modal still visible: {modal_visible}, productionRowIdx: {row_idx_val}")
        ok2 = modal_visible and row_idx_val == '5'
        print("  " + ("✅ PASS" if ok2 else "❌ FAIL"))

        print("[3] Exit with NO unsaved changes -> closes immediately...")
        page.click("#productionExitBtn")
        page.wait_for_timeout(400)
        modal_visible_after_exit = page.locator("#editProductionModal").is_visible()
        print(f"  modal visible after Exit: {modal_visible_after_exit}")
        ok3 = not modal_visible_after_exit
        print("  " + ("✅ PASS" if ok3 else "❌ FAIL"))

        print("[4] Re-open, make an unsaved edit, click Exit -> confirm dialog appears...")
        page.evaluate("App.Production.openEditModal(0)")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(300)
        page.fill("#productionRemarks", "unsaved change, not saved")
        page.click("#productionExitBtn")
        page.wait_for_timeout(300)
        confirm_visible = page.locator("text=You have unsaved changes").count() > 0
        print(f"  confirm dialog visible: {confirm_visible}")
        modal_still_open = page.locator("#editProductionModal").is_visible()
        print(f"  edit modal still open (blocked by confirm): {modal_still_open}")
        ok4 = confirm_visible and modal_still_open

        print("  " + ("✅ PASS" if ok4 else "❌ FAIL"))

        if console_errors:
            print("\n⚠️ console errors:")
            for e in console_errors:
                print(" ", e)

        browser.close()
        return ok1 and ok2 and ok3 and ok4 and not console_errors


if __name__ == "__main__":
    ok = run()
    print("\n" + ("ALL PASS" if ok else "SOME FAILED"))
    sys.exit(0 if ok else 1)
