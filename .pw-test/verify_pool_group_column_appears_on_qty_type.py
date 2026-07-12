"""
Verification script: a Per-Process Pool Components table's column must appear
as soon as a checked color gets a real quantity typed in, not just at the
moment its checkbox is toggled.

Reproduces the reported bug: a Primary-axis color row (e.g. a Painted Frame
color) has no auto-filled quantity when checked -- only non-primary rows do
(see handleColorCheckToggle) -- so at check time _checkedColorTokensLower()
(which requires qty > 0) doesn't yet count it, and syncPoolColorGroupColumns()
(only called on checkbox toggle) never adds its column. The operator's normal
workflow -- check the box, THEN type its quantity -- left the Per-Process Pool
Components table permanently empty, since onColorQtyChanged never re-synced
columns; only unchecking and rechecking after typing the qty would reveal it.

Run: python .pw-test/verify_pool_group_column_appears_on_qty_type.py
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
    "processId": "PRC-1110", "processName": "Painting Orbit Sports 16 inch", "sequence": 3,
    "lotPrefix": "PT", "outputItemName": "Painted Frame Orbit Sports 16 inch", "isFinalStage": False,
    "active": True, "processType": "General", "primaryColorAxis": "Painted Frame Orbit Sports 16 inch"
}
MOCK_COMPONENTS = [
    {"itemName": "Painted Frame Orbit Sports 16 inch", "size": "", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
    {"itemName": "Fitted Rim 16 inch", "size": "", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
]
MOCK_POOL = (
    [{"outputItemName": "Painted Frame Orbit Sports 16 inch", "color": c, "qty": 100}
     for c in ["Blue-White", "Orange-GREY", "Red-SeaGreen", "Red-Yellow"]]
    + [{"outputItemName": "Fitted Rim 16 inch", "color": c, "qty": 100} for c in ["BCP", "Black"]]
)
MOCK_AXES = {
    "success": True,
    "data": {
        "axes": [
            {"key": "pool:painted frame orbit sports 16 inch", "label": "Painted Frame Orbit Sports 16 inch",
             "colors": ["Blue-White", "Orange-GREY", "Red-SeaGreen", "Red-Yellow"], "source": "pool"},
            {"key": "pool:fitted rim 16 inch", "label": "Fitted Rim 16 inch",
             "colors": ["BCP", "Black"], "source": "pool"}
        ],
        "primaryColorAxis": "Painted Frame Orbit Sports 16 inch",
        "primaryAxisKey": "pool:painted frame orbit sports 16 inch"
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
        page = browser.new_context(viewport={"width": 1600, "height": 1000}).new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.resolve().as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        page.evaluate(f"""
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
            App.State.globalItems = [];
            App.State.globalColors = [];
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

        failures = []
        def check(cond, msg):
            print(("PASS: " if cond else "FAIL: ") + msg)
            if not cond:
                failures.append(msg)

        print("\n[Step 1] Open Create modal, select the process (2-axis, Painted Frame is Primary)...")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.evaluate("""
            document.getElementById('productionProcessId').value='PRC-1110';
            App.Production.handleProcessChange('PRC-1110');
        """)
        page.wait_for_timeout(800)

        print("\n[Step 2] Check 'Blue-White' WITHOUT typing a qty yet (real operator workflow: check first)...")
        page.evaluate("""
            const row = Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row'))
                .find(r => r.dataset.color === 'Blue-White');
            const chk = row.querySelector('.production-color-check');
            chk.checked = true;
            App.Production.handleColorCheckToggle(chk);
        """)
        page.wait_for_timeout(300)

        cols_before_qty = page.evaluate("""
            (() => {
                const table = Array.from(document.querySelectorAll('#productionPoolColorGroupsContainer table.prod-color-table'))
                    .find(t => t.dataset.allColors && t.dataset.allColors.includes('Blue-White'));
                if (!table) return null;
                return Array.from(table.querySelectorAll('thead th[data-color]')).map(th => th.dataset.color);
            })()
        """)
        print(f"  Painted Frame table columns right after check (qty still empty): {cols_before_qty}")

        print("\n[Step 3] NOW type the qty for 'Blue-White' (separate step, no re-toggle of the checkbox)...")
        page.evaluate("""
            const row = Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row'))
                .find(r => r.dataset.color === 'Blue-White');
            const qty = row.querySelector('.production-color-qty');
            qty.value = '12';
            qty.dispatchEvent(new Event('input', { bubbles: true }));
        """)
        page.wait_for_timeout(300)

        cols_after_qty = page.evaluate("""
            (() => {
                const table = Array.from(document.querySelectorAll('#productionPoolColorGroupsContainer table.prod-color-table'))
                    .find(t => t.dataset.allColors && t.dataset.allColors.includes('Blue-White'));
                if (!table) return null;
                return Array.from(table.querySelectorAll('thead th[data-color]')).map(th => th.dataset.color);
            })()
        """)
        print(f"  Painted Frame table columns after typing qty (no re-check): {cols_after_qty}")
        check(cols_after_qty == ['Blue-White'], "column appears once qty is typed, without needing to re-toggle the checkbox")

        wrapper_display = page.evaluate("""
            document.getElementById('productionPoolColorGroupsWrapper').style.display
        """)
        check(wrapper_display == '', f"outer section wrapper is visible (got display='{wrapper_display}')")

        cell_value = page.evaluate("""
            (() => {
                const table = Array.from(document.querySelectorAll('#productionPoolColorGroupsContainer table.prod-color-table'))
                    .find(t => t.dataset.allColors && t.dataset.allColors.includes('Blue-White'));
                const input = table.querySelector("input.pool-group-qty[data-color='Blue-White']");
                return input ? input.value : null;
            })()
        """)
        check(cell_value == '12', f"the new column's cell shows the typed qty (got {cell_value})")

        if console_errors:
            print("\n  Console/page errors:")
            for e in console_errors:
                print(f"    {e}")

        print("\n" + ("ALL TESTS PASSED" if not failures else f"{len(failures)} TEST(S) FAILED"))
        browser.close()
        return len(failures) == 0


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
