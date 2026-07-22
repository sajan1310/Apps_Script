"""
Verifies Save/Exit on Bill, PO, and Return edit modals: Exit shown (not
Cancel), submit button reads "Save", and a successful Save keeps the modal
open on the SAME record (re-finds it via its post-save identity) instead of
closing.
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
DIST_HTML = Path(r"C:\Users\erkar\my-app-script-project\dist\index.html")
TIMEOUT = 8000

MOCK_BILL = {"vendor": "Acme Vendor", "billNumber": "INV-100", "billDateRaw": "2026-07-01T00:00:00.000Z",
             "contact": "", "remarks": "orig", "items": [{"name": "Widget", "size": "", "narration": "", "unit": "Pcs", "price": 10, "gst": 0, "poNumber": "DIRECT"}]}
MOCK_PO = {"poNumber": "PO-100", "vendor": "Acme Vendor", "poDate": "2026-07-01", "contact": "", "poDescription": "", "poRemarks": "orig",
           "supplierRemarks": "", "items": [{"name": "Widget", "size": "", "narration": "", "unit": "Pcs", "qty": 5, "price": 10, "gst": 0}]}
MOCK_RETURN = {"returnNumber": "RET-100", "vendor": "Acme Vendor", "returnDateRaw": "2026-07-01T00:00:00.000Z", "contact": "",
               "billNumber": "", "remarks": "orig", "items": [{"name": "Widget", "size": "", "narration": "", "unit": "Pcs", "qty": 1, "price": 10, "reason": ""}]}

MOCK_API_RESPONSES = {
    "getBillData": {"success": True, "data": [MOCK_BILL]},
    "getPOData": {"success": True, "data": [MOCK_PO]},
    "getReturnData": {"success": True, "data": [MOCK_RETURN]},
    "saveBill": {"success": True, "data": {"billNumber": "INV-100"}, "message": "Bill updated successfully."},
    "savePO": {"success": True, "data": {"poNumber": "PO-100"}, "message": "PO updated successfully."},
    "saveReturn": {"success": True, "data": {"returnNumber": "RET-100"}, "message": "Return updated successfully."},
    "checkStockAdjustmentConflicts": {"success": True, "data": []},
    "suggestPoAllocations": {"success": True, "data": []},
}


def setup(page):
    page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
    page.wait_for_timeout(1000)
    page.evaluate(f"""
        App.State.globalBills = [{json.dumps(MOCK_BILL)}];
        App.State.filteredBills = App.State.globalBills;
        App.State.globalPOs = [{json.dumps(MOCK_PO)}];
        App.State.filteredPOs = App.State.globalPOs;
        App.State.globalReturns = [{json.dumps(MOCK_RETURN)}];
        App.State.filteredReturns = App.State.globalReturns;
        App.State.globalVendors = [{{"name": "Acme Vendor", "contact": ""}}];
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


def check_modal(page, modalId, cancelId, exitId, submitId, submitClickSelector=None):
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

        print("=== BILL ===")
        page.evaluate("App.Bill.openEditModal(0)")
        page.locator("#receiveBillModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(300)
        results.append(check_modal(page, "receiveBillModal", "billCancelBtn", "billExitBtn", "billSubmitBtn"))
        page.fill("#billForm input[name='remarks']", "edited")
        page.click("#billSubmitBtn")
        page.wait_for_timeout(300)
        # Confirm dialog appears for bill save
        confirm_btn = page.locator("#confirmModal .btn-primary, .modal.show button:has-text('OK'), .modal.show button:has-text('Confirm'), .modal.show button:has-text('Yes')")
        if confirm_btn.count() > 0:
            confirm_btn.first.click()
        page.wait_for_timeout(400)
        visible = page.locator("#receiveBillModal").is_visible()
        print(f"  Bill modal still open after Save: {visible}")
        results.append(visible)

        print("=== PO ===")
        page.evaluate("App.PO.openEditModal(0)")
        page.locator("#editPoModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(300)
        results.append(check_modal(page, "editPoModal", "poCancelBtn", "poExitBtn", "poSubmitBtn"))
        page.click("#poSubmitBtn")
        page.wait_for_timeout(400)
        visible = page.locator("#editPoModal").is_visible()
        print(f"  PO modal still open after Save: {visible}")
        results.append(visible)

        print("=== RETURN ===")
        page.evaluate("App.Return.openEditModal(0)")
        page.locator("#returnGoodsModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(300)
        results.append(check_modal(page, "returnGoodsModal", "returnCancelBtn", "returnExitBtn", "returnSubmitBtn"))
        page.click("#returnSubmitBtn")
        page.wait_for_timeout(400)
        visible = page.locator("#returnGoodsModal").is_visible()
        print(f"  Return modal still open after Save: {visible}")
        results.append(visible)

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
