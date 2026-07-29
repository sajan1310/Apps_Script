"""
Baseline regression check: after opening Edit Lot for an existing Production
lot, editing an ALREADY-CHECKED color's Qty in the "Colors to Produce"
checklist must live-update that color's cell in the Per-Process Pool
Components table. This already worked before the qty-sync fixes (see
verify_pool_group_qty_sync_new_color_edit.py /
verify_pool_group_qty_sync_percolor_ratio.py for the cases that didn't) --
kept here so a future change can't silently break the simple case while
fixing an edge case. Based on the same mock harness as
verify_pool_group_columns_edit_lot.py.

Run: python .pw-test/verify_pool_group_qty_sync_existing_color_edit.py
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
    "processId": "PRC-FTF", "processName": "Fitting Frame", "sequence": 5,
    "lotPrefix": "FTF", "outputItemName": "Fitted Frame Assembled", "isFinalStage": False,
    "active": True, "processType": "General", "primaryColorAxis": "Painted Frame Crysta 20 inch D/Gaddi"
}
MOCK_COMPONENTS = [
    {"itemName": "Painted Frame Crysta 20 inch D/Gaddi", "size": "", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
    {"itemName": "20 inch Mudguard", "size": "", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
]
MOCK_POOL = (
    [{"outputItemName": "Painted Frame Crysta 20 inch D/Gaddi", "color": c, "processId": "PRC-P", "qty": 10}
     for c in ["Blue-White", "Orange-White", "Pink-White"]]
    + [{"outputItemName": "20 inch Mudguard", "color": c, "processId": "PRC-M", "qty": 10}
       for c in ["Blue", "Orange", "Pink", "Black", "Grey", "Green", "Metallic Green", "B/T Green", "Silky Blue"]]
)
MOCK_AXES = {
    "success": True,
    "data": {
        "axes": [
            {"key": "pool:painted frame crysta 20 inch d/gaddi", "label": "Painted Frame Crysta 20 inch D/Gaddi",
             "colors": ["Blue-White", "Orange-White", "Pink-White"], "source": "pool"},
            {"key": "pool:20 inch mudguard", "label": "20 inch Mudguard",
             "colors": ["B/T Green", "Black", "Blue", "Green", "Grey", "Metallic Green", "Orange", "Pink", "Silky Blue"], "source": "pool"}
        ],
        "primaryColorAxis": "Painted Frame Crysta 20 inch D/Gaddi",
        "primaryAxisKey": "pool:painted frame crysta 20 inch d/gaddi"
    }
}

SAVED_LOT = {
    "lotNumber": "FTF-0001", "processId": "PRC-FTF", "status": "Completed",
    "qty": 10, "color": "Blue-White", "assignedTo": "Sanjay",
    "colorBreakdown": [{"color": "Blue-White", "qty": 10}, {"color": "Blue", "qty": 10}],
    "componentsConsumed": [
        {"itemName": "Painted Frame Crysta 20 inch D/Gaddi", "size": "", "sourceType": "POOL", "qty": 10, "colorGroup": "Blue-White"},
        {"itemName": "20 inch Mudguard", "size": "", "sourceType": "POOL", "qty": 10, "colorGroup": "Blue"},
    ],
}

MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": sorted(list({r["color"] for r in MOCK_POOL}))},
    "getProcessColorAxes": MOCK_AXES,
    "getProcessComponentsData": {"success": True, "data": MOCK_COMPONENTS},
    "getWarehousePoolData": {"success": True, "data": MOCK_POOL},
    "getProcessWipData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
}


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context()
        page = ctx.new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        page.evaluate(f"""
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
            App.State.globalItems = [];
            App.State.globalColors = [];
            App.State.globalContractors = [{{ contractorId: 'C1', name: 'Sanjay', active: true }}];
            App.State.globalProduction = [{json.dumps(SAVED_LOT)}];
            window.__mockResponses = {json.dumps(MOCK_API_RESPONSES)};
            window.google = {{
                script: {{
                    run: {{
                        withSuccessHandler(cb) {{
                            const runner = {{ withFailureHandler() {{ return runner; }} }};
                            Object.keys(window.__mockResponses).forEach(method => {{
                                runner[method] = (...args) => setTimeout(() => cb(window.__mockResponses[method]), 20);
                            }});
                            return runner;
                        }}
                    }}
                }}
            }};
        """)

        ok = True
        print("[Step 1] Open Edit modal for the saved lot...")
        page.evaluate("App.Production.openEditModal(0)")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(1000)

        before = page.evaluate("""
            (() => {
                const input = document.querySelector('#productionPoolColorGroupsContainer .pool-group-qty[data-color="Blue-White"]');
                return input ? input.value : null;
            })()
        """)
        print(f"  'Blue-White' pool cell BEFORE edit: {before}")

        print("[Step 2] Change Blue-White's checklist Qty from 10 to 20...")
        page.evaluate("""
            (() => {
                const row = Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row'))
                    .find(r => r.dataset.color === 'Blue-White');
                const qtyInput = row.querySelector('.production-color-qty');
                qtyInput.value = '20';
                qtyInput.dispatchEvent(new Event('input', { bubbles: true }));
            })()
        """)
        page.wait_for_timeout(300)

        after = page.evaluate("""
            (() => {
                const input = document.querySelector('#productionPoolColorGroupsContainer .pool-group-qty[data-color="Blue-White"]');
                return input ? input.value : null;
            })()
        """)
        print(f"  'Blue-White' pool cell AFTER edit: {after}")

        if after != '20':
            print(f"  FAIL: expected pool cell to update to 20 (qtyPerUnit=1 x qty=20), got {after}")
            ok = False
        else:
            print("  PASS: pool cell live-updated with new qty")

        if console_errors:
            print("\n  Console/page errors:")
            for e in console_errors:
                print(f"    {e}")
            ok = False

        browser.close()
        return ok


if __name__ == "__main__":
    ok = run()
    print("\n" + ("ALL PASS" if ok else "SOME FAILED"))
    sys.exit(0 if ok else 1)
