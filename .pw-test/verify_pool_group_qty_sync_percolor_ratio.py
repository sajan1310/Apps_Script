"""
Verifies each color of a Per-Process Pool Components row keeps its OWN
derived qty-per-unit ratio, not a shared row-level one borrowed from
whichever color happened to be processed first in populateComponentsConsumedDirect.

Scenario: "Painted Frame" was recorded on this lot for two colors with
DIFFERENT consumption ratios: Blue-White qty=10 for checklist qty=10 (ratio 1.0),
Orange-White qty=4.5 for checklist qty=5 (ratio 0.9). Editing Orange-White's
checklist qty to 10 should recompute its pool cell using ITS OWN ratio (0.9 x 10 = 9),
not Blue-White's ratio (1.0 x 10 = 10).

Run: python .pw-test/verify_pool_group_qty_sync_percolor_ratio.py
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
]
MOCK_POOL = [
    {"outputItemName": "Painted Frame Crysta 20 inch D/Gaddi", "color": c, "processId": "PRC-P", "qty": 10}
    for c in ["Blue-White", "Orange-White", "Pink-White"]
]
MOCK_AXES = {
    "success": True,
    "data": {
        "axes": [
            {"key": "pool:painted frame crysta 20 inch d/gaddi", "label": "Painted Frame Crysta 20 inch D/Gaddi",
             "colors": ["Blue-White", "Orange-White", "Pink-White"], "source": "pool"},
        ],
        "primaryColorAxis": "Painted Frame Crysta 20 inch D/Gaddi",
        "primaryAxisKey": "pool:painted frame crysta 20 inch d/gaddi"
    }
}

# Two colors on the same lot, each with its OWN historical consumption ratio.
SAVED_LOT = {
    "lotNumber": "FTF-0002", "processId": "PRC-FTF", "status": "Completed",
    "qty": 15, "color": "Blue-White / Orange-White", "assignedTo": "Sanjay",
    "colorBreakdown": [{"color": "Blue-White", "qty": 10}, {"color": "Orange-White", "qty": 5}],
    "componentsConsumed": [
        {"itemName": "Painted Frame Crysta 20 inch D/Gaddi", "size": "", "sourceType": "POOL", "qty": 10, "colorGroup": "Blue-White"},
        {"itemName": "Painted Frame Crysta 20 inch D/Gaddi", "size": "", "sourceType": "POOL", "qty": 4.5, "colorGroup": "Orange-White"},
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
        print("[Step 1] Open Edit modal (Blue-White ratio=1.0, Orange-White ratio=0.9)...")
        page.evaluate("App.Production.openEditModal(0)")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(1000)

        before = page.evaluate("""
            (() => {
                const bw = document.querySelector('#productionPoolColorGroupsContainer .pool-group-qty[data-color="Blue-White"]');
                const ow = document.querySelector('#productionPoolColorGroupsContainer .pool-group-qty[data-color="Orange-White"]');
                return {
                    blueWhite: bw ? { value: bw.value, qtyPerUnit: bw.dataset.qtyPerUnit } : null,
                    orangeWhite: ow ? { value: ow.value, qtyPerUnit: ow.dataset.qtyPerUnit } : null,
                };
            })()
        """)
        print(f"  Initial cells: {before}")
        if before['orangeWhite']['qtyPerUnit'] != '0.9':
            print(f"  FAIL: expected Orange-White's OWN ratio (4.5/5=0.9) stamped, got {before['orangeWhite']['qtyPerUnit']}")
            ok = False
        else:
            print("  PASS: Orange-White carries its own distinct ratio (0.9), not Blue-White's (1.0)")

        print("[Step 2] Change Orange-White's checklist Qty from 5 to 10...")
        page.evaluate("""
            (() => {
                const row = Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row'))
                    .find(r => r.dataset.color === 'Orange-White');
                const qtyInput = row.querySelector('.production-color-qty');
                qtyInput.value = '10';
                qtyInput.dispatchEvent(new Event('input', { bubbles: true }));
            })()
        """)
        page.wait_for_timeout(300)

        after = page.evaluate("""
            (() => {
                const bw = document.querySelector('#productionPoolColorGroupsContainer .pool-group-qty[data-color="Blue-White"]');
                const ow = document.querySelector('#productionPoolColorGroupsContainer .pool-group-qty[data-color="Orange-White"]');
                return { blueWhite: bw ? bw.value : null, orangeWhite: ow ? ow.value : null };
            })()
        """)
        print(f"  After edit: {after}")
        # Orange-White: 0.9 x 10 = 9 (its own ratio) -- NOT 1.0 x 10 = 10 (Blue-White's ratio)
        if after['orangeWhite'] != '9':
            print(f"  FAIL: expected Orange-White to recompute with ITS OWN ratio (0.9 x 10 = 9), got {after['orangeWhite']}")
            ok = False
        else:
            print("  PASS: Orange-White recomputed using its own ratio (0.9 x 10 = 9), unaffected by Blue-White's (1.0)")

        if after['blueWhite'] != '10':
            print(f"  FAIL: Blue-White cell should be untouched (still 10), got {after['blueWhite']}")
            ok = False
        else:
            print("  PASS: Blue-White cell unaffected by Orange-White's edit")

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
