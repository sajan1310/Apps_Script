"""
Confirms the select2DropdownParent fix (Script_Core.html) generalizes past
the one dialog it was already regression-tested against
(verify_drift_dropdown_scroll_tracking.py, Script_Items.html) — this probes
Script_Production.html's Process Select2 instead (initProcessSelect2), a
completely different file/call site, inside the Create Production Lot
modal, which is tall enough to scroll on its own with enough processes
seeded.
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(r"C:\Users\erkar\my-app-script-project\dist\index.html")
TIMEOUT = 8000

MOCK_PROCESSES = [
    {"processId": f"PRC-{i}", "processName": f"12in Frame Painting {i}", "sequence": i,
     "lotPrefix": f"FP{i}", "outputItemName": "12 inch Painted Frame", "isFinalStage": False,
     "active": True, "processType": "General"}
    for i in range(1, 12)
]
MOCK_MODELS = [{"name": "General", "remarks": ""}]
MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": []},
    "getProcessComponentsData": {"success": True, "data": []},
    "getWarehousePoolData": {"success": True, "data": []},
    "getProcessWipData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
}


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1400, "height": 700})
        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)
        page.evaluate(f"""
            App.State.globalProcesses = {json.dumps(MOCK_PROCESSES)};
            App.State.globalModels = {json.dumps(MOCK_MODELS)};
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

        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.evaluate("""
            document.getElementById('productionSize').value = '12 inch';
            App.Production.handleSizeChange('12 inch');
        """)
        page.wait_for_timeout(300)
        page.evaluate("""
            document.getElementById('productionModel').value = 'General';
            App.Production.handleModelChange('General');
        """)
        page.wait_for_timeout(300)

        modal_body = page.locator("#editProductionModal .modal-body")
        scroll_h = page.evaluate("document.querySelector('#editProductionModal .modal-body').scrollHeight")
        client_h = page.evaluate("document.querySelector('#editProductionModal .modal-body').clientHeight")
        print(f"modal-body scrollHeight={scroll_h} clientHeight={client_h} (scrollable: {scroll_h > client_h})")

        print("[1] Open Process dropdown, then scroll modal-body -> should close, not detach")
        trigger = page.locator("#productionProcessId + .select2 .select2-selection")
        trigger.click()
        page.wait_for_timeout(200)
        open_before = page.locator(".select2-dropdown").count()
        print(f"  dropdown open before scroll: {open_before}")
        page.evaluate("document.querySelector('#editProductionModal .modal-body').scrollTop += 200")
        page.wait_for_timeout(200)
        open_after = page.locator(".select2-dropdown").count()
        print(f"  dropdown open after scroll: {open_after}")
        ok1 = open_before == 1 and open_after == 0
        print("  " + ("✅ PASS" if ok1 else "❌ FAIL"))

        page.evaluate("document.querySelector('#editProductionModal .modal-body').scrollTop = 0")
        page.wait_for_timeout(200)

        print("[2] Scroll modal-body first, then open Process dropdown -> should position at trigger")
        page.evaluate("document.querySelector('#editProductionModal .modal-body').scrollTop = 150")
        page.wait_for_timeout(200)
        tbox = trigger.bounding_box()
        trigger.click()
        page.wait_for_timeout(250)
        dd = page.locator(".select2-dropdown")
        ok2 = dd.count() == 1
        if ok2:
            dbox = dd.bounding_box()
            gap = dbox["y"] - (tbox["y"] + tbox["height"])
            print(f"  gap={gap:.1f}px")
            ok2 = abs(gap) < 15
        print("  " + ("✅ PASS" if ok2 else "❌ FAIL"))

        if console_errors:
            print("\n⚠️ console errors:")
            for e in console_errors:
                print(" ", e)

        browser.close()
        return ok1 and ok2 and not console_errors


if __name__ == "__main__":
    ok = run()
    print("\n" + ("ALL PASS" if ok else "SOME FAILED"))
    sys.exit(0 if ok else 1)
