"""
Verifies Save/Exit on Item Master, BOM, and Process edit modals: Exit shown
(not Cancel), submit button reads "Save", and a successful Save keeps the
modal open on the SAME record instead of closing.
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
DIST_HTML = Path(r"C:\Users\erkar\my-app-script-project\dist\index.html")
TIMEOUT = 8000

MOCK_ITEM = {"name": "Widget", "size": "Small", "remarks": "orig", "narration": "", "specification": "",
             "baseUnit": "Pcs", "purchaseUnit": "Pcs", "weightPerBaseUnit": "", "vendors": []}
MOCK_BOM = {"productId": "PRD-1001", "productName": "Cruiser", "remarks": "orig",
            "components": [{"itemName": "Widget", "size": "Small", "narration": "", "rate": 10, "vendor": "V1",
                             "qtyPerProduct": 1, "lineCost": 10, "processId": "", "processGroup": "General"}],
            "totalCost": 10, "totalQty": 1, "additionalCosts": [], "totalAdditionalCost": 0, "grandTotal": 10}
MOCK_PROCESS = {"processId": "PRC-1", "processName": "Packing", "sequence": 1, "lotPrefix": "PK",
                 "outputItemName": "Packed Bicycle", "isFinalStage": True, "active": True, "processType": "General",
                 "remarks": "orig"}

MOCK_API_RESPONSES = {
    "saveItem": {"success": True, "data": {}, "message": "Item saved."},
    "getItemsData": {"success": True, "data": [MOCK_ITEM]},
    "saveBOM": {"success": True, "message": "BOM saved."},
    "getBOMData": {"success": True, "data": [MOCK_BOM]},
    "saveProcess": {"success": True, "message": "Process saved."},
    "getProcessData": {"success": True, "data": [MOCK_PROCESS]},
    "getProcessColorGroups": {"success": True, "data": []},
    "getWarehousePoolData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
    "getNextProductId": {"success": True, "data": "PRD-1002"},
}


def setup(page):
    page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
    page.wait_for_timeout(1000)
    # BOM is password-gated (App.BOM.getToken()/loadData) — seed a session
    # token so App.BOM.loadData() (called after a successful save) doesn't
    # bail into promptForAccess() and pop the access modal on top of
    # everything else.
    page.evaluate("sessionStorage.setItem('bomAccessToken', 'test-token')")
    page.evaluate(f"""
        App.State.globalItems = [{json.dumps(MOCK_ITEM)}];
        App.State.globalStock = [];
        App.State.globalBOMs = [{json.dumps(MOCK_BOM)}];
        App.State.filteredBOMs = App.State.globalBOMs;
        App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
        App.State.filteredProcesses = App.State.globalProcesses;
        App.State.globalProcessTypes = [{{"name": "General"}}];
        App.State.globalVendors = [{{"name": "V1"}}];
        App.State.globalContractors = [];
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


def check_modal(page, cancelId, exitId, submitId):
    state = page.evaluate(f"""() => ({{
        cancelVisible: document.getElementById('{cancelId}').style.display !== 'none',
        exitVisible: document.getElementById('{exitId}').style.display !== 'none',
        submitText: document.getElementById('{submitId}').innerText.trim()
    }})""")
    print(f"  {state}")
    ok = (not state['cancelVisible']) and state['exitVisible'] and state['submitText'] == 'Save'
    print("  " + ("✅ PASS button state" if ok else "❌ FAIL button state"))
    return ok


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1400, "height": 900})
        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))
        setup(page)
        results = []

        def force_close(modal_id):
            page.evaluate(f"""bootstrap.Modal.getInstance(document.getElementById('{modal_id}'))?.hide()""")
            page.wait_for_timeout(400)

        print("=== ITEM MASTER ===")
        page.evaluate("App.Item.openEditModal('Widget', 'Small')")
        page.locator("#itemModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(300)
        results.append(check_modal(page, "itemCancelBtn", "itemExitBtn", "itemSubmitBtn"))
        page.click("#itemSubmitBtn")
        page.wait_for_timeout(400)
        visible = page.locator("#itemModal").is_visible()
        print(f"  Item modal still open after Save: {visible}")
        results.append(visible)
        force_close("itemModal")

        print("=== BOM ===")
        page.evaluate("App.BOM.openEditModal(0)")
        page.locator("#editBomModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(300)
        results.append(check_modal(page, "bomCancelBtn", "bomExitBtn", "bomSubmitBtn"))
        page.click("#bomSubmitBtn")
        page.wait_for_timeout(400)
        visible = page.locator("#editBomModal").is_visible()
        print(f"  BOM modal still open after Save: {visible}")
        results.append(visible)
        force_close("editBomModal")

        print("=== PROCESS ===")
        page.evaluate("App.Process.openEditModal(0)")
        page.locator("#editProcessModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(300)
        results.append(check_modal(page, "processCancelBtn", "processExitBtn", "processSubmitBtn"))
        page.click("#processSubmitBtn")
        page.wait_for_timeout(400)
        visible = page.locator("#editProcessModal").is_visible()
        print(f"  Process modal still open after Save: {visible}")
        results.append(visible)
        force_close("editProcessModal")

        if console_errors:
            print("\n⚠️ console errors:")
            for e in console_errors:
                print(" ", e)

        browser.close()
        return all(results) and not console_errors


if __name__ == "__main__":
    ok = run()
    print("\n" + ("ALL PASS" if ok else "SOME FAILED"))
    sys.exit(0 if ok else 1)
