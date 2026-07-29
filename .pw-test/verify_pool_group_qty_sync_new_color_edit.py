"""
Covers a fixed bug in the Edit Lot path: checking a brand-new color (one
this lot's saved componentsConsumed never recorded before) inserted its
Per-Process Pool Components column with NO data-qty-per-unit stamped (see
syncPoolColorGroupColumns' old `def.mode === 'create'` gate on
qtyPerUnitAttr) -- so refreshPoolColorGroupCells could never fill it in
afterward, no matter what qty the operator typed into the checklist. Fixed
by stamping every cell's qty-per-unit mode-agnostically via the new
_poolCellQtyPerUnit helper.

Run: python .pw-test/verify_pool_group_qty_sync_new_color_edit.py
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
        print("[Step 1] Open Edit modal for the saved lot (only has Blue-White/Blue recorded)...")
        page.evaluate("App.Production.openEditModal(0)")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(1000)

        print("[Step 2] Check a brand-new color 'Orange-White' (never used on this lot before) and give it Qty 5...")
        result = page.evaluate("""
            (() => {
                const row = Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row'))
                    .find(r => r.dataset.color === 'Orange-White');
                if (!row) return { error: 'no Orange-White row found in checklist' };
                const chk = row.querySelector('.production-color-check');
                chk.checked = true;
                chk.dispatchEvent(new Event('change', { bubbles: true }));
                return { found: true };
            })()
        """)
        print(f"  check result: {result}")
        page.wait_for_timeout(300)

        page.evaluate("""
            (() => {
                const row = Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row'))
                    .find(r => r.dataset.color === 'Orange-White');
                const qtyInput = row.querySelector('.production-color-qty');
                qtyInput.value = '5';
                qtyInput.dispatchEvent(new Event('input', { bubbles: true }));
            })()
        """)
        page.wait_for_timeout(300)

        cell = page.evaluate("""
            (() => {
                const input = document.querySelector('#productionPoolColorGroupsContainer .pool-group-qty[data-color="Orange-White"]');
                return input ? { value: input.value, qtyPerUnit: input.dataset.qtyPerUnit } : null;
            })()
        """)
        print(f"  'Orange-White' pool cell after check+qty=5: {cell}")

        if not cell:
            print("  FAIL: no 'Orange-White' column/cell rendered in the Pool Components table at all")
            ok = False
        elif cell['value'] != '5':
            print(f"  FAIL: expected cell value '5' (qtyPerUnit=1 x qty=5), got '{cell['value']}' (qtyPerUnit attr: {cell['qtyPerUnit']})")
            ok = False
        else:
            print("  PASS: new color's pool cell reflects the checklist qty")

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
