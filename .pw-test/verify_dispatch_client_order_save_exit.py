"""
Verifies Save/Exit on Dispatch, Client, and Client Order (PI/Estimate)
edit modals.
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
DIST_HTML = Path(r"C:\Users\erkar\my-app-script-project\dist\index.html")
TIMEOUT = 8000

MOCK_DISPATCH = {"rowIdx": 5, "dispatchNumber": "DSP-1", "dateRaw": "2026-07-01T00:00:00.000Z",
                  "transport": "", "invoiceNumber": "", "privateMark": "", "grNumber": "", "remarks": "orig",
                  "orderNumber": "", "productId": "PRD-1", "productName": "Cruiser", "qty": 2, "clientName": "Acme Client"}
MOCK_CLIENT = {"name": "Acme Client", "contact": "", "gstin": "", "address": "", "remarks": "orig"}
MOCK_ORDER = {"orderNumber": "PI-1", "orderDate": "01/07/2026", "dateRaw": "2026-07-01T00:00:00.000Z",
              "status": "Estimate", "orderRemarks": "orig", "clientName": "Acme Client", "lines": []}

MOCK_API_RESPONSES = {
    "getDispatchData": {"success": True, "data": [MOCK_DISPATCH]},
    "getReadyToDispatchData": {"success": True, "data": []},
    "getClientsData": {"success": True, "data": [MOCK_CLIENT]},
    "getClientOrdersData": {"success": True, "data": [MOCK_ORDER]},
    "saveDispatch": {"success": True, "message": "Dispatch updated successfully."},
    "saveClient": {"success": True, "message": "Client updated successfully."},
    "saveClientOrder": {"success": True, "message": "PI / Estimate updated successfully."},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
}


def setup(page):
    page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
    page.wait_for_timeout(1000)
    page.evaluate(f"""
        App.State.globalDispatch = [{json.dumps(MOCK_DISPATCH)}];
        App.State.filteredDispatch = App.State.globalDispatch;
        App.State.globalClients = [{json.dumps(MOCK_CLIENT)}];
        App.State.filteredClients = App.State.globalClients;
        App.State.globalOrders = [{json.dumps(MOCK_ORDER)}];
        App.State.filteredOrders = App.State.globalOrders;
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
            page.evaluate(f"bootstrap.Modal.getInstance(document.getElementById('{modal_id}'))?.hide()")
            page.wait_for_timeout(400)

        print("=== DISPATCH ===")
        page.evaluate("App.Dispatch.openEditDispatchModal(0)")
        page.locator("#dispatchModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(300)
        results.append(check_modal(page, "dispatchCancelBtn", "dispatchExitBtn", "dispatchSubmitBtn"))
        page.click("#dispatchSubmitBtn")
        page.wait_for_timeout(400)
        visible = page.locator("#dispatchModal").is_visible()
        print(f"  Dispatch modal still open after Save: {visible}")
        results.append(visible)
        force_close("dispatchModal")

        print("=== CLIENT ===")
        page.evaluate("App.Client.openEditClientModal('Acme Client')")
        page.locator("#clientModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(300)
        results.append(check_modal(page, "clientCancelBtn", "clientExitBtn", "clientSubmitBtn"))
        page.click("#clientSubmitBtn")
        page.wait_for_timeout(400)
        visible = page.locator("#clientModal").is_visible()
        print(f"  Client modal still open after Save: {visible}")
        results.append(visible)
        force_close("clientModal")

        print("=== CLIENT ORDER ===")
        page.evaluate("App.Client.openEditOrderModal(0)")
        page.locator("#clientOrderModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(300)
        results.append(check_modal(page, "clientOrderCancelBtn", "clientOrderExitBtn", "clientOrderSubmitBtn"))
        page.click("#clientOrderSubmitBtn")
        page.wait_for_timeout(400)
        visible = page.locator("#clientOrderModal").is_visible()
        print(f"  Client Order modal still open after Save: {visible}")
        results.append(visible)
        force_close("clientOrderModal")

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
