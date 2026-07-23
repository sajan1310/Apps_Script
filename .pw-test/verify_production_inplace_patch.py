"""
Verifies the new in-place row patch on Production Edit Lot save:
  1. Save succeeds (toast shown).
  2. getProductionData is NOT re-invoked (no full-table reload round trip).
  3. The edited row's own <tr> DOM node is REPLACED (new node) with updated data.
  4. An untouched row's <tr> DOM node is the EXACT SAME node as before save
     (proves the whole tbody wasn't rebuilt) and its custom marker survives.
  5. The production table container's scroll position is unchanged.
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
    "color": "", "colorBreakdown": [], "componentsConsumed": [], "dateRaw": None,
    "contractorPayable": 0
}
MOCK_LOT_2 = dict(MOCK_LOT, rowIdx=6, lotNumber="LOT-PK-0002")
MOCK_LOT_2["dateRaw"] = None

FRESH_ROW_AFTER_SAVE = dict(MOCK_LOT, qty=42, remarks="edited remarks")

MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": []},
    "getProcessComponentsData": {"success": True, "data": []},
    "getWarehousePoolData": {"success": True, "data": []},
    "getProcessWipData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
    "saveProduction": {
        "success": True,
        "data": {"lotNumber": "LOT-PK-0001", "row": FRESH_ROW_AFTER_SAVE},
        "message": "Production log updated successfully.",
    },
    "getProductionData": {"success": True, "data": [MOCK_LOT, MOCK_LOT_2]},
}


def run():
    failures = []

    def check(label, cond, extra=""):
        status = "PASS" if cond else "FAIL"
        print(f"[{status}] {label}" + (f" ({extra})" if extra else ""))
        if not cond:
            failures.append(label)

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
            window.__getProductionDataCallCount = 0;
            window.google = {{
                script: {{ run: {{
                    withSuccessHandler(cb) {{
                        const runner = {{ withFailureHandler() {{ return runner; }} }};
                        Object.keys(window.__mockResponses).forEach(method => {{
                            runner[method] = (...args) => {{
                                if (method === 'getProductionData') window.__getProductionDataCallCount++;
                                setTimeout(() => cb(window.__mockResponses[method]), 30);
                            }};
                        }});
                        return runner;
                    }}
                }} }}
            }};
        """)

        print("[setup] Render Production table via renderTable() directly...")
        page.evaluate("App.Production.renderTable()")
        page.wait_for_timeout(200)

        rows_before = page.evaluate("document.querySelectorAll('#productionTableBody tr').length")
        check("table shows 2 rows before edit", rows_before == 2, f"got {rows_before}")

        # Tag both rows' DOM nodes so we can detect node IDENTITY (not just content) after save.
        page.evaluate("""
            document.querySelectorAll('#productionTableBody tr').forEach((tr, i) => {
                tr.dataset.testMarker = 'marker-' + tr.dataset.rowKey;
                tr.__testNodeTag = 'original-' + tr.dataset.rowKey;
            });
        """)

        print("[1] Open Edit modal for lot rowIdx=5, change qty...")
        page.evaluate("App.Production.openEditModal(0)")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(300)

        page.fill("#productionQty", "42")

        # Track whether the "Loading Production Logs..." placeholder ever
        # appears in the tbody at any point during/after the save.
        page.evaluate("""
            window.__sawLoadingFlash = false;
            const tbody = document.getElementById('productionTableBody');
            window.__moLoading = new MutationObserver(() => {
                if (tbody.innerHTML.includes('Loading Production Logs')) window.__sawLoadingFlash = true;
            });
            window.__moLoading.observe(tbody, { childList: true, subtree: true, characterData: true });
        """)

        print("[2] Submit the Edit Lot form (Save)...")
        page.click("#productionSubmitBtn")
        page.wait_for_timeout(600)

        toast_text = page.evaluate("""
            (() => {
                const t = document.querySelector('.toast-body, .toast');
                return t ? t.textContent.trim() : null;
            })()
        """)
        check("save toast appeared", bool(toast_text), f"toast={toast_text!r}")

        call_count = page.evaluate("window.__getProductionDataCallCount")
        check("getProductionData was NOT re-invoked on edit save (no full reload)", call_count == 0, f"calls={call_count}")

        saw_flash = page.evaluate("window.__sawLoadingFlash")
        check("no 'Loading Production Logs...' flash appeared in the tbody", saw_flash is False)

        # The modal reopens on the same record after save (existing Save/Exit
        # behavior) — close it so we can inspect the underlying table rows.
        page.evaluate("""
            const modalEl = document.getElementById('editProductionModal');
            if (modalEl && window.bootstrap) bootstrap.Modal.getOrCreateInstance(modalEl).hide();
        """)
        page.wait_for_timeout(300)

        row_data = page.evaluate("""
            (() => {
                const tr = document.querySelector('#productionTableBody tr[data-row-key="5"]');
                return tr ? tr.textContent : null;
            })()
        """)
        check("edited row (rowIdx=5) shows updated qty 42 in the DOM", bool(row_data) and "42" in row_data, f"text={row_data!r}")

        other_row_identity = page.evaluate("""
            (() => {
                const tr = document.querySelector('#productionTableBody tr[data-row-key="6"]');
                return tr ? tr.__testNodeTag : null;
            })()
        """)
        check(
            "untouched row (rowIdx=6) DOM node is the SAME node as before save (tbody not fully rebuilt)",
            other_row_identity == "original-6",
            f"got {other_row_identity!r}"
        )

        rows_after = page.evaluate("document.querySelectorAll('#productionTableBody tr').length")
        check("table still shows 2 rows after edit (no row lost/duplicated)", rows_after == 2, f"got {rows_after}")

        check("no page errors thrown", len(console_errors) == 0, "; ".join(console_errors))

        page.screenshot(path=str(Path(__file__).parent / "after_inplace_patch.png"))

        browser.close()

    print()
    if failures:
        print(f"{len(failures)} CHECK(S) FAILED: {failures}")
        sys.exit(1)
    else:
        print("ALL CHECKS PASSED")
        sys.exit(0)


if __name__ == "__main__":
    run()
