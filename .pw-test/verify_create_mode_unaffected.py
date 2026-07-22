"""
Confirms Create-mode buttons (Cancel + Save-labeled submit) are correctly
restored after having been in Edit mode (Cancel hidden, Exit shown, "Save")
on the same modal instance — the button-state toggle must be absolute, not
leak between mode switches.
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
DIST_HTML = Path(r"C:\Users\erkar\my-app-script-project\dist\index.html")
TIMEOUT = 8000

MOCK_PROCESS = {"processId": "PRC-1", "processName": "Packing", "sequence": 1, "lotPrefix": "PK",
                 "outputItemName": "Packed Bicycle", "isFinalStage": True, "active": True, "processType": "General"}
MOCK_LOT = {"rowIdx": 5, "lotNumber": "LOT-PK-0001", "processId": "PRC-1", "qty": 10, "date": "01/07/2026",
            "assignedBy": "Admin", "assignedTo": "Contractor A", "status": "Pending", "remarks": "orig",
            "productId": "", "productName": "", "color": "", "colorBreakdown": [], "componentsConsumed": []}

MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": []},
    "getProcessComponentsData": {"success": True, "data": []},
    "getWarehousePoolData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
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
            App.State.globalProduction = [{json.dumps(MOCK_LOT)}];
            App.State.filteredProduction = App.State.globalProduction;
            App.State.globalContractors = [{{"name": "Contractor A"}}];
            App.State.globalItems = [];
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

        print("[1] Open Edit modal (rowIdx=5)...")
        page.evaluate("App.Production.openEditModal(0)")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(300)
        edit_state = page.evaluate("""() => ({
            cancelVisible: document.getElementById('productionCancelBtn').style.display !== 'none',
            exitVisible: document.getElementById('productionExitBtn').style.display !== 'none',
            submitText: document.getElementById('productionSubmitBtn').innerText.trim()
        })""")
        print(f"  Edit mode: {edit_state}")

        print("[2] Close via Exit, then open Create modal...")
        page.evaluate("bootstrap.Modal.getInstance(document.getElementById('editProductionModal'))?.hide()")
        page.wait_for_timeout(400)
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(500)
        create_state = page.evaluate("""() => ({
            cancelVisible: document.getElementById('productionCancelBtn').style.display !== 'none',
            exitVisible: document.getElementById('productionExitBtn').style.display !== 'none',
            submitText: document.getElementById('productionSubmitBtn').innerText.trim()
        })""")
        print(f"  Create mode: {create_state}")

        ok = (edit_state['cancelVisible'] is False and edit_state['exitVisible'] is True and edit_state['submitText'] == 'Save'
              and create_state['cancelVisible'] is True and create_state['exitVisible'] is False and create_state['submitText'] == 'Record Production Run')
        print("  " + ("✅ PASS — no leakage between modes" if ok else "❌ FAIL — button state leaked between modes"))

        if console_errors:
            print("\n⚠️ console errors:")
            for e in console_errors:
                print(" ", e)

        browser.close()
        return ok and not console_errors


if __name__ == "__main__":
    ok = run()
    print("\n" + ("ALL PASS" if ok else "SOME FAILED"))
    sys.exit(0 if ok else 1)
