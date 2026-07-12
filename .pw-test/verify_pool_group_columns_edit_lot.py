"""
Verifies the same "only checked/saved colors form columns" fix applies to
the EDIT LOT path (renderPoolColorGroupsFromAccum) — reopening an existing
saved Production lot should show Per-Process Pool Components columns only
for colors this lot actually has recorded values for, not every color that
item has ever had Warehouse Pool stock in.

Run: python .pw-test/verify_pool_group_columns_edit_lot.py
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

# The existing saved lot only ever used Blue-White (10) and its matched
# Mudguard color Blue (10) — none of the other 3+9 possible colors.
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

        cols = page.evaluate("""
            (() => {
                const out = {};
                document.querySelectorAll('#productionPoolColorGroupsContainer table.prod-color-table').forEach(t => {
                    const key = t.dataset.allColors;
                    out[key] = Array.from(t.querySelectorAll('thead th[data-color]')).map(th => th.dataset.color);
                });
                return out;
            })()
        """)
        print(f"  Rendered table columns: {json.dumps(cols, indent=2)}")

        painted_key = next((k for k in cols if 'Blue-White' in k), None)
        mudguard_key = next((k for k in cols if 'Metallic Green' in k), None)

        if not painted_key or cols[painted_key] != ['Blue-White']:
            print("  FAIL: Painted Frame table should show only ['Blue-White'] (this lot's saved color), not all 3 pool colors")
            ok = False
        else:
            print("  PASS: Painted Frame table scoped to just this lot's saved 'Blue-White'")

        if not mudguard_key or cols[mudguard_key] != ['Blue']:
            print("  FAIL: Mudguard table should show only ['Blue'] (this lot's saved color), not all 9 pool colors")
            ok = False
        else:
            print("  PASS: Mudguard table scoped to just this lot's saved 'Blue'")

        cell_val = page.evaluate("""
            (() => {
                const input = document.querySelector('#productionPoolColorGroupsContainer .pool-group-qty[data-color="Blue-White"]');
                return input ? input.value : null;
            })()
        """)
        print(f"  'Blue-White' cell value: {cell_val}")
        if cell_val != '10':
            print("  FAIL: expected the saved qty (10) preserved in the cell")
            ok = False
        else:
            print("  PASS: saved qty correctly shown")

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
