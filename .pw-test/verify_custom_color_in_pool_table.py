"""
Regression test: reported live — a custom color ("Purple-White") added via
"+ Add Custom Sub-Group" appeared correctly checked/quantified in the
Colors to Produce checklist and in the "Per-Color Components" matrix at the
top, but never got a column in the "Per-Process Pool Components" table at
the bottom — leaving no way to actually record how many units of that
color were produced/consumed from the pool. Reproduced on a process with
only ONE pool group (so the "which group?" picker never even showed).

Root cause: syncPoolColorGroupColumns only ever recognizes a color already
present in a pool item's real Warehouse Pool history (_poolColorGroupDefs
entry's own `colors`), and a custom name is by definition not part of any
item's real history. Fixed by having addCustomColorRow extend the
correlated pool def's own `colors` list when the custom color belongs to
(or, with only one pool group, unambiguously must belong to) a POOL-sourced
group — see _poolDefForCustomColor.

Run: python .pw-test/verify_custom_color_in_pool_table.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

# Mirrors the live report exactly: a single pool-color axis process, no
# Primary Axis needed since there's nothing to disambiguate.
MOCK_PROCESS = {
    "processId": "PRC-FF14", "processName": "Fitting Frame 14 inch", "sequence": 5,
    "lotPrefix": "FF14", "outputItemName": "Fitted Frame 14 inch Scooby Steel Rim", "isFinalStage": False,
    "active": True, "processType": "General", "primaryColorAxis": ""
}
MOCK_COMPONENTS = [
    {"itemName": "Painted Frame Rider 14 inch D/Gaddi", "size": "", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
]
COLORS = ["Blue-White", "Pink-White", "Red-White", "SeaGreen-White"]
MOCK_POOL = [{"outputItemName": "Painted Frame Rider 14 inch D/Gaddi", "color": c, "processId": "PRC-P", "qty": 10} for c in COLORS]
MOCK_AXES = {"success": True, "data": {"axes": [], "primaryColorAxis": "", "primaryAxisKey": ""}}
MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": COLORS},
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
        print("[1] Open Create modal, select the single-pool-group process...")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.evaluate("""
            document.getElementById('productionProcessId').value = 'PRC-FF14';
            App.Production.handleProcessChange('PRC-FF14');
        """)
        page.wait_for_timeout(700)

        picker_display = page.evaluate("getComputedStyle(document.getElementById('productionCustomColorGroupSelect')).display")
        if picker_display != 'none':
            print(f"  FAIL: expected the group picker hidden for a single-group process (got {picker_display})")
            ok = False
        else:
            print("  PASS: picker correctly hidden (only 1 pool group, no ambiguity)")

        print("\n[2] Check the 4 real colors...")
        page.evaluate("""
            [['Blue-White','14'],['Pink-White','14'],['Red-White','4'],['SeaGreen-White','14']].forEach(([color, qty]) => {
                const row = Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row'))
                    .find(r => r.dataset.color === color);
                const chk = row.querySelector('.production-color-check');
                chk.checked = true;
                App.Production.handleColorCheckToggle(chk);
                const q = row.querySelector('.production-color-qty');
                q.disabled = false;
                q.value = qty;
            });
        """)
        page.wait_for_timeout(400)

        print("\n[3] Add a custom color 'Purple-White' (no group choice possible/needed)...")
        page.evaluate("""
            document.getElementById('productionCustomColorInput').add(new Option('Purple-White', 'Purple-White', true, true));
            App.Production.addCustomColorRow();
            const row = Array.from(document.querySelectorAll('.production-color-row')).find(r => r.dataset.color === 'Purple-White');
            row.querySelector('.production-color-qty').disabled = false;
            row.querySelector('.production-color-qty').value = '14';
        """)
        page.wait_for_timeout(400)

        pool_cols = page.evaluate("""
            Array.from(document.querySelectorAll('#productionPoolColorGroupsContainer table.prod-color-table'))
                .map(t => Array.from(t.querySelectorAll('thead th[data-color]')).map(th => th.dataset.color))
        """)
        print(f"  Per-Process Pool Components columns: {pool_cols}")
        if not pool_cols or "Purple-White" not in pool_cols[0]:
            print("  FAIL: expected 'Purple-White' to appear as a column in the pool table")
            ok = False
        else:
            print("  PASS: custom color correctly rendered as a Pool Components column")
            expected_order = sorted(["Blue-White", "Pink-White", "Red-White", "SeaGreen-White", "Purple-White"])
            if pool_cols[0] != expected_order:
                print(f"  FAIL: expected alphabetically sorted {expected_order}, got {pool_cols[0]}")
                ok = False
            else:
                print("  PASS: inserted at the correct alphabetical position")

        matrix_cols = page.evaluate("""
            Array.from(document.querySelectorAll('#productionColorMatrixHeaderRow th[data-color]')).map(th => th.dataset.color)
        """)
        print(f"  Per-Color Components (top matrix) columns: {matrix_cols}")
        if "Purple-White" not in matrix_cols:
            print("  FAIL: custom color should still also appear in the top matrix (unchanged prior behavior)")
            ok = False
        else:
            print("  PASS: top matrix still shows it too")

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
