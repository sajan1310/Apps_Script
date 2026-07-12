"""
Verifies the three P0 fixes:

P0.1 - initContractorSelect2 (Production) and initLogisticsContractorSelect2
       (Dispatch) no longer leak a 'change' handler on repeated calls (each
       call must call $select.off('change') before re-attaching).

P0.2 - resetCreateForm() is a real tear-down+rebuild: it now awaits the full
       Size->Model->Process Type->Process->color-checklist cascade instead of
       firing it and returning early. Simulated with an artificially slow
       getProcessColorGroups response — before the fix, resetCreateForm()
       would resolve before the checklist finished rendering; after the fix,
       it does not resolve until the checklist rows are present. Also checks
       #productionFormFieldset exists so the submit handler can lock it.

Run: python .pw-test/verify_p0_fixes.py
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


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context()
        page = ctx.new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        url = DIST_HTML.as_uri()
        page.goto(url, wait_until="domcontentloaded")
        page.wait_for_timeout(500)

        page.evaluate(f"""
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
            App.State.globalItems = [];
            App.State.globalColors = [];
            App.State.globalContractors = [{{name: 'Sanjay'}}];
            window.__mockResponses = {{
                getProcessColorGroups: {{ success: true, data: ['Blue', 'Orange-White'] }},
                getProcessComponentsData: {{ success: true, data: {json.dumps(MOCK_COMPONENTS)} }},
                getWarehousePoolData: {{ success: true, data: [] }},
                getProcessWipData: {{ success: true, data: [] }},
                getStockData: {{ success: true, data: [] }},
                getContractorRateForProcess: {{ success: true, data: {{ ratePerUnit: 0 }} }}
            }};
            window.__delayMs = {{ getProcessColorGroups: 300 }};
            window.google = {{
                script: {{
                    run: {{
                        withSuccessHandler(cb) {{
                            const runner = {{ withFailureHandler() {{ return runner; }} }};
                            Object.keys(window.__mockResponses).forEach(method => {{
                                runner[method] = (...args) => {{
                                    const delay = (window.__delayMs && window.__delayMs[method]) || 20;
                                    setTimeout(() => cb(window.__mockResponses[method]), delay);
                                }};
                            }});
                            return runner;
                        }}
                    }}
                }}
            }};
        """)

        print("[P0.1] initContractorSelect2 change-handler leak (Production)")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(500)

        page.evaluate("""
            for (let i = 0; i < 5; i++) App.Production.initContractorSelect2('');
        """)
        handlerCount = page.evaluate("""
            (() => {
                const el = document.getElementById('productionAssignedTo');
                const events = window.jQuery._data(el, 'events');
                return (events && events.change) ? events.change.length : 0;
            })()
        """)
        check(handlerCount == 1, f"exactly 1 'change' handler on productionAssignedTo after 5 init calls (got {handlerCount})")

        print("\n[P0.1] initLogisticsContractorSelect2 change-handler leak (Dispatch)")
        page.evaluate("""
            App.State.globalReadyToDispatch = [];
            for (let i = 0; i < 5; i++) App.Dispatch.initLogisticsContractorSelect2('');
        """)
        dispatchHandlerCount = page.evaluate("""
            (() => {
                const el = document.getElementById('dispatchLogisticsContractor');
                if (!el) return -1;
                const events = window.jQuery._data(el, 'events');
                return (events && events.change) ? events.change.length : 0;
            })()
        """)
        check(dispatchHandlerCount == 1, f"exactly 1 'change' handler on dispatchLogisticsContractor after 5 init calls (got {dispatchHandlerCount})")

        print("\n[P0.2] productionFormFieldset exists (submit handler can lock the form)")
        fieldsetExists = page.evaluate("!!document.getElementById('productionFormFieldset')")
        check(fieldsetExists, "#productionFormFieldset is present in the DOM")

        print("\n[P0.2] resetCreateForm() awaits the full cascade before resolving")
        # Only one active process exists, so Size->Model->Type->Process all
        # auto-select and the cascade runs all the way to the color
        # checklist, whose fetch (getProcessColorGroups) is artificially
        # delayed 300ms above.
        rowsRightAfterReset = page.evaluate("""
            (async () => {
                await App.Production.resetCreateForm();
                return document.querySelectorAll('#productionColorChecklist .production-color-row').length;
            })()
        """)
        check(rowsRightAfterReset == 2, f"checklist already has both color rows the instant resetCreateForm() resolves (got {rowsRightAfterReset} rows) — proves the delayed getProcessColorGroups fetch was awaited, not fired-and-forgotten")

        if console_errors:
            print("\nConsole/page errors:")
            for e in console_errors:
                print(f"    {e}")
            failures += len(console_errors)

        browser.close()


if __name__ == "__main__":
    run()
    print(f"\n{'ALL PASS' if failures == 0 else str(failures) + ' FAILURE(S)'}")
    sys.exit(0 if failures == 0 else 1)
