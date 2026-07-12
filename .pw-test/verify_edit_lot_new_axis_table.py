"""
Regression test: reported as "the production form is unable to render a
new per-color pool component table" — reproduced when reopening an
EXISTING saved lot for edit, then checking a color belonging to a
pool-color axis/item the lot never used when it was originally saved
(e.g. adding Mudguard usage to a lot that only ever used Painted Frame).

Root cause: populateComponentsConsumedDirect (Script.html) only seeded a
Per-Process Pool Components table for items present in the lot's OWN saved
componentsConsumed array, not for every pool-color-aware item in the
process's current recipe. A brand-new axis this lot never touched before
had no table at all, so syncPoolColorGroupColumns had nowhere to attach a
newly-checked color's column.

Run: python .pw-test/verify_edit_lot_new_axis_table.py
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
PAINTED_COLORS = ["Blue-White", "Orange-White"]
MUDGUARD_COLORS = ["Blue", "Orange"]
MOCK_POOL = (
    [{"outputItemName": "Painted Frame Crysta 20 inch D/Gaddi", "color": c, "processId": "PRC-P", "qty": 10} for c in PAINTED_COLORS]
    + [{"outputItemName": "20 inch Mudguard", "color": c, "processId": "PRC-M", "qty": 10} for c in MUDGUARD_COLORS]
)
MOCK_AXES = {
    "success": True,
    "data": {
        "axes": [
            {"key": "pool:painted frame crysta 20 inch d/gaddi", "label": "Painted Frame Crysta 20 inch D/Gaddi", "colors": PAINTED_COLORS, "source": "pool"},
            {"key": "pool:20 inch mudguard", "label": "20 inch Mudguard", "colors": MUDGUARD_COLORS, "source": "pool"}
        ],
        "primaryColorAxis": "Painted Frame Crysta 20 inch D/Gaddi",
        "primaryAxisKey": "pool:painted frame crysta 20 inch d/gaddi"
    }
}
# This lot was originally saved using ONLY Painted Frame - Mudguard was
# never touched at all (no row for it in componentsConsumed).
SAVED_LOT = {
    "lotNumber": "FTF-0001", "processId": "PRC-FTF", "status": "Completed", "rowIdx": 2,
    "qty": 10, "color": "Blue-White", "assignedTo": "Sanjay",
    "colorBreakdown": [{"color": "Blue-White", "qty": 10}],
    "componentsConsumed": [
        {"itemName": "Painted Frame Crysta 20 inch D/Gaddi", "size": "", "sourceType": "POOL", "qty": 10, "colorGroup": "Blue-White"},
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
        page = browser.new_context().new_page()
        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.resolve().as_uri(), wait_until="domcontentloaded")
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
        print("[1] Open Edit modal for a lot that only ever used Painted Frame (never Mudguard)...")
        page.evaluate("App.Production.openEditModal(0)")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(800)

        tables_initial = page.evaluate("""
            Array.from(document.querySelectorAll('#productionPoolColorGroupsContainer table.prod-color-table'))
                .map(t => t.dataset.allColors)
        """)
        print(f"  tables present at open: {tables_initial}")
        has_mudguard_table = any('Blue' in t and 'Orange' in t and 'White' not in t for t in tables_initial)
        if not has_mudguard_table:
            print("  FAIL: expected a (hidden, 0-column) Mudguard table to already exist, ready for a new color")
            ok = False
        else:
            print("  PASS: Mudguard table exists (even though this lot never used it)")

        print("\n[2] Check a Mudguard color ('Blue') this lot never used before...")
        page.evaluate("""
            const row = Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row'))
                .find(r => r.dataset.color === 'Blue');
            const chk = row.querySelector('.production-color-check');
            chk.checked = true;
            App.Production.handleColorCheckToggle(chk);
            const q = row.querySelector('.production-color-qty');
            q.disabled = false;
            q.value = '5';
        """)
        page.wait_for_timeout(500)

        mudguard_cols = page.evaluate("""
            (() => {
                const table = Array.from(document.querySelectorAll('#productionPoolColorGroupsContainer table.prod-color-table'))
                    .find(t => t.dataset.allColors && t.dataset.allColors.includes('Blue') && t.dataset.allColors.includes('Orange') && !t.dataset.allColors.includes('White'));
                return table ? Array.from(table.querySelectorAll('thead th[data-color]')).map(th => th.dataset.color) : null;
            })()
        """)
        print(f"  Mudguard table columns after checking 'Blue': {mudguard_cols}")
        if mudguard_cols != ['Blue']:
            print("  FAIL: expected the Mudguard table to gain a 'Blue' column")
            ok = False
        else:
            print("  PASS: new axis's color correctly rendered a live column, ready to record consumption")

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
