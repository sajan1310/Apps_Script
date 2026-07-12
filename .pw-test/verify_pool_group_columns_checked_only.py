"""
Verifies the fix for: Per-Process Pool Components tables (the small
per-item color-quantity grids under "Log Production Lot") were showing a
column for EVERY color that item has ever had Warehouse Pool stock in,
even colors nobody checked in the "Colors to Produce" checklist above —
e.g. a Mudguard table showing 12 columns when only 5 colors were actually
being used on this lot. Columns should now only appear for colors that are
actually checked, and should add/remove live as checkboxes are toggled.

Run: python .pw-test/verify_pool_group_columns_checked_only.py
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

MOCK_ITEMS = []

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

        url = DIST_HTML.as_uri()
        page.goto(url, wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        page.evaluate(f"""
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
            App.State.globalItems = {json.dumps(MOCK_ITEMS)};
            App.State.globalColors = [];
            window.__mockResponses = {json.dumps(MOCK_API_RESPONSES)};
            window.google = {{
                script: {{
                    run: {{
                        withSuccessHandler(cb) {{
                            const runner = {{
                                withFailureHandler() {{ return runner; }}
                            }};
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

        print("[Step 1] Open Create modal, select Fitting Frame...")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.evaluate("""
            document.getElementById('productionProcessId').value = 'PRC-FTF';
            App.Production.handleProcessChange('PRC-FTF');
        """)
        page.wait_for_timeout(800)

        print("[Step 2] Before checking anything, Per-Process Pool tables should show NO color columns...")
        col_counts = page.evaluate("""
            Array.from(document.querySelectorAll('#productionPoolColorGroupsContainer table.prod-color-table'))
                .map(t => t.querySelectorAll('thead th[data-color]').length)
        """)
        print(f"  color-column counts per table: {col_counts}")
        if any(c > 0 for c in col_counts):
            print("  FAIL: a table already shows color columns before anything is checked")
            ok = False
        else:
            print("  PASS")

        print("[Step 3] Check 'Blue-White' (Painted Frame primary axis)...")
        page.evaluate("""
            const row = Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row'))
                .find(r => r.dataset.color === 'Blue-White');
            const chk = row.querySelector('.production-color-check');
            chk.checked = true;
            const qty = row.querySelector('.production-color-qty');
            if (qty) { qty.disabled = false; qty.value = '10'; }
            App.Production.handleColorCheckToggle(chk);
        """)
        page.wait_for_timeout(400)

        painted_cols = page.evaluate("""
            (() => {
                const table = Array.from(document.querySelectorAll('#productionPoolColorGroupsContainer table.prod-color-table'))
                    .find(t => t.dataset.allColors && t.dataset.allColors.includes('Blue-White'));
                if (!table) return null;
                return Array.from(table.querySelectorAll('thead th[data-color]')).map(th => th.dataset.color);
            })()
        """)
        print(f"  Painted Frame table columns now: {painted_cols}")
        if painted_cols != ['Blue-White']:
            print("  FAIL: expected exactly ['Blue-White']")
            ok = False
        else:
            print("  PASS")

        mudguard_cols_after_step3 = page.evaluate("""
            (() => {
                const table = Array.from(document.querySelectorAll('#productionPoolColorGroupsContainer table.prod-color-table'))
                    .find(t => t.dataset.allColors && t.dataset.allColors.includes('Metallic Green'));
                if (!table) return null;
                return Array.from(table.querySelectorAll('thead th[data-color]')).map(th => th.dataset.color);
            })()
        """)
        print(f"  Mudguard table columns after step 3 (auto-sync should have checked 'Blue'): {mudguard_cols_after_step3}")
        if mudguard_cols_after_step3 != ['Blue']:
            print("  FAIL: expected auto-sync to add exactly ['Blue'] to the Mudguard table")
            ok = False
        else:
            print("  PASS")

        print("[Step 4] Check 'Orange-White' too...")
        page.evaluate("""
            const row = Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row'))
                .find(r => r.dataset.color === 'Orange-White');
            const chk = row.querySelector('.production-color-check');
            chk.checked = true;
            const qty = row.querySelector('.production-color-qty');
            if (qty) { qty.disabled = false; qty.value = '7'; }
            App.Production.handleColorCheckToggle(chk);
        """)
        page.wait_for_timeout(400)
        painted_cols2 = page.evaluate("""
            (() => {
                const table = Array.from(document.querySelectorAll('#productionPoolColorGroupsContainer table.prod-color-table'))
                    .find(t => t.dataset.allColors && t.dataset.allColors.includes('Blue-White'));
                return Array.from(table.querySelectorAll('thead th[data-color]')).map(th => th.dataset.color);
            })()
        """)
        print(f"  Painted Frame table columns now: {painted_cols2}")
        if painted_cols2 != ['Blue-White', 'Orange-White']:
            print("  FAIL: expected sorted ['Blue-White', 'Orange-White'] (alphabetical insert position)")
            ok = False
        else:
            print("  PASS (new column inserted at correct sorted position, existing one preserved)")

        print("[Step 5] Uncheck 'Blue-White' — its column (and Mudguard's synced 'Blue') should disappear...")
        page.evaluate("""
            const row = Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row'))
                .find(r => r.dataset.color === 'Blue-White');
            const chk = row.querySelector('.production-color-check');
            chk.checked = false;
            App.Production.handleColorCheckToggle(chk);
        """)
        page.wait_for_timeout(400)
        painted_cols3 = page.evaluate("""
            (() => {
                const table = Array.from(document.querySelectorAll('#productionPoolColorGroupsContainer table.prod-color-table'))
                    .find(t => t.dataset.allColors && t.dataset.allColors.includes('Blue-White'));
                return Array.from(table.querySelectorAll('thead th[data-color]')).map(th => th.dataset.color);
            })()
        """)
        mudguard_cols_after_step5 = page.evaluate("""
            (() => {
                const table = Array.from(document.querySelectorAll('#productionPoolColorGroupsContainer table.prod-color-table'))
                    .find(t => t.dataset.allColors && t.dataset.allColors.includes('Metallic Green'));
                return Array.from(table.querySelectorAll('thead th[data-color]')).map(th => th.dataset.color);
            })()
        """)
        print(f"  Painted Frame table columns now: {painted_cols3}")
        print(f"  Mudguard table columns now: {mudguard_cols_after_step5}")
        if painted_cols3 != ['Orange-White']:
            print("  FAIL: expected only ['Orange-White'] left")
            ok = False
        else:
            print("  PASS")
        # "Orange" stays: it was auto-synced from "Orange-White" (still
        # checked from Step 4), independent of "Blue"'s own auto-sync being
        # torn down here. Only "Blue" (matched to the just-unchecked
        # "Blue-White") should disappear.
        if mudguard_cols_after_step5 != ['Orange']:
            print("  FAIL: expected Mudguard's 'Blue' column gone but 'Orange' (still matched to checked Orange-White) to remain")
            ok = False
        else:
            print("  PASS")

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
