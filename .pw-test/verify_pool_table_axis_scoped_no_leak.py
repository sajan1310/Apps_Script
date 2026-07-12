"""
Verification script: Per-Process Pool Components tables must stay strictly
scoped to their OWN sub-group/axis — a color checked under one axis must
never "slide in" as a column in a DIFFERENT axis's table, and two axes
coincidentally sharing a literal color name (e.g. a "Painted Rim" pool item
and an independent "Painted Frame" pool item each having their own
"Purple") must never have their quantities summed into each other's table.

This is the front-end half of a class of bug already fixed server-side for
the lot's own total Qty (see bug_axis_color_name_collision_qty in project
memory) and for the checklist's own row-identity (see
verify_edit_lot_color_restore.py) — this script covers the remaining gap:
_checkedPoolGroupColors / syncPoolColorGroupColumns / _poolGroupCellValue /
refreshPoolColorGroupCells in Script.html previously derived a Pool
Components table's visible columns AND suggested quantities from EVERY
checked color in the WHOLE checklist regardless of which axis it came from,
letting an unrelated axis's same-named color both (a) make an irrelevant
column appear in a table that never had that color checked at all, and
(b) get summed into another axis's own suggested (and, if unnoticed,
actually submitted) Warehouse Pool consumption.

Run: python .pw-test/verify_pool_table_axis_scoped_no_leak.py
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
    "processId": "PRC-1", "processName": "Fitted Frame Assembly", "sequence": 1,
    "lotPrefix": "FFA", "outputItemName": "Fitted Frame", "isFinalStage": False,
    "active": True, "processType": "General"
}

MOCK_COMPONENTS = [
    {"itemName": "Painted Rim", "size": "", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
    {"itemName": "Painted Frame", "size": "", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
]

MOCK_POOL_ROWS = [
    {"outputItemName": "Painted Rim", "color": "Blue", "processId": "PRC-R", "qty": 50},
    {"outputItemName": "Painted Rim", "color": "Purple", "processId": "PRC-R", "qty": 50},
    {"outputItemName": "Painted Frame", "color": "Green", "processId": "PRC-F", "qty": 50},
    {"outputItemName": "Painted Frame", "color": "Purple", "processId": "PRC-F", "qty": 50},
]

MOCK_AXES = {
    "axes": [
        {"key": "pool:painted rim", "label": "Painted Rim", "colors": ["Blue", "Purple"], "source": "pool"},
        {"key": "pool:painted frame", "label": "Painted Frame", "colors": ["Green", "Purple"], "source": "pool"},
    ],
    "primaryColorAxis": "Painted Rim",
    "primaryAxisKey": "pool:painted rim",
}

MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": ["Blue", "Green", "Purple"]},
    "getProcessColorAxes": {"success": True, "data": MOCK_AXES},
    "getProcessComponentsData": {"success": True, "data": MOCK_COMPONENTS},
    "getWarehousePoolData": {"success": True, "data": MOCK_POOL_ROWS},
    "getProcessWipData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
}


def get_pool_table_cells(page):
    return page.evaluate("""
        Array.from(document.querySelectorAll('#productionPoolColorGroupsContainer table')).map(table => ({
            axisKey: table.dataset.axisKey,
            columns: Array.from(table.querySelectorAll('thead th[data-color]')).map(th => th.dataset.color),
            cells: Array.from(table.querySelectorAll('.pool-group-qty')).map(inp => ({ color: inp.dataset.color, value: inp.value }))
        }))
    """)


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
            window.__mockResponses = {json.dumps(MOCK_API_RESPONSES)};
            window.google = {{
                script: {{
                    run: {{
                        withSuccessHandler(cb) {{
                            const runner = {{ withFailureHandler() {{ return runner; }} }};
                            Object.keys(window.__mockResponses).forEach(method => {{
                                runner[method] = (...args) => setTimeout(() => cb(window.__mockResponses[method]), 30);
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

        print("[Step 1] Open Create modal, select the 2-pool-axis process (Rim + Frame, both having a 'Purple')...")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.evaluate("""
            document.getElementById('productionProcessId').value = 'PRC-1';
            App.Production.handleProcessChange('PRC-1');
        """)
        page.wait_for_timeout(800)

        print("\n[Step 2] Check ONLY Rim's Blue=10 (no Purple checked anywhere yet)...")
        page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Blue"] .production-color-check').click()
        """)
        page.wait_for_timeout(150)
        page.locator("#productionColorChecklist .production-color-row[data-color='Blue'] .production-color-qty").fill("10")
        page.wait_for_timeout(300)

        tables = get_pool_table_cells(page)
        print("  tables:", json.dumps(tables))
        rim_table = next((t for t in tables if t['axisKey'] == 'pool:painted rim'), None)
        frame_table = next((t for t in tables if t['axisKey'] == 'pool:painted frame'), None)
        check(rim_table is not None and 'Blue' in rim_table['columns'], f"Rim table shows its own Blue column (got {rim_table})")
        check(frame_table is None or frame_table['columns'] == [], f"Frame table has NO columns yet — nothing of its own is checked (got {frame_table})")

        print("\n[Step 3] Check Rim's Purple=6 (Primary axis) — Frame's own matching 'Purple' auto-syncs (expected, real feature)...")
        page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Purple"][data-group="pool:painted rim"] .production-color-check').click()
        """)
        page.wait_for_timeout(150)
        page.locator('#productionColorChecklist .production-color-row[data-color="Purple"][data-group="pool:painted rim"] .production-color-qty').fill("6")
        page.wait_for_timeout(300)

        frame_purple_row = page.evaluate("""
            (() => {
                const row = document.querySelector('#productionColorChecklist .production-color-row[data-color="Purple"][data-group="pool:painted frame"]');
                return row ? { checked: row.querySelector('.production-color-check').checked, qty: row.querySelector('.production-color-qty').value } : null;
            })()
        """)
        check(frame_purple_row is not None and frame_purple_row['checked'] and frame_purple_row['qty'] == '6',
              f"Frame's own Purple auto-checked and auto-filled to match Rim's 6 (real, unrelated feature) (got {frame_purple_row})")

        print("\n[Step 4] Both tables' OWN 'Purple' column shows 6 — NOT double-counted to 12...")
        tables = get_pool_table_cells(page)
        print("  tables:", json.dumps(tables))
        rim_table = next((t for t in tables if t['axisKey'] == 'pool:painted rim'), None)
        frame_table = next((t for t in tables if t['axisKey'] == 'pool:painted frame'), None)
        rim_purple_cell = next((c for c in (rim_table or {}).get('cells', []) if c['color'] == 'Purple'), None)
        frame_purple_cell = next((c for c in (frame_table or {}).get('cells', []) if c['color'] == 'Purple'), None)
        check(rim_purple_cell is not None and rim_purple_cell['value'] == '6', f"Rim table's own Purple cell = 6, not summed with Frame's (got {rim_purple_cell})")
        check(frame_purple_cell is not None and frame_purple_cell['value'] == '6', f"Frame table's own Purple cell = 6, not summed with Rim's (got {frame_purple_cell})")

        print("\n[Step 5] Uncheck Frame's Purple (break the auto-sync) — Frame's table must lose its Purple column even though Rim's stays checked at 6...")
        page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Purple"][data-group="pool:painted frame"] .production-color-check').click()
        """)
        page.wait_for_timeout(300)
        tables = get_pool_table_cells(page)
        frame_table = next((t for t in tables if t['axisKey'] == 'pool:painted frame'), None)
        rim_table = next((t for t in tables if t['axisKey'] == 'pool:painted rim'), None)
        check(frame_table is None or 'Purple' not in frame_table['columns'],
              f"Frame table's Purple column disappears once its OWN Purple is unchecked, despite Rim's Purple still being checked (got {frame_table})")
        check(rim_table is not None and 'Purple' in rim_table['columns'] and
              next(c for c in rim_table['cells'] if c['color'] == 'Purple')['value'] == '6',
              f"Rim table's own Purple column is untouched by Frame's uncheck (got {rim_table})")

        print("\n[Step 6] Re-check Frame's Purple and manually override it to 9 (diverging from Rim's 6) — each table keeps its own distinct value...")
        page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Purple"][data-group="pool:painted frame"] .production-color-check').click()
        """)
        page.wait_for_timeout(150)
        page.locator('#productionColorChecklist .production-color-row[data-color="Purple"][data-group="pool:painted frame"] .production-color-qty').fill("9")
        page.wait_for_timeout(300)

        tables = get_pool_table_cells(page)
        print("  tables:", json.dumps(tables))
        rim_table = next((t for t in tables if t['axisKey'] == 'pool:painted rim'), None)
        frame_table = next((t for t in tables if t['axisKey'] == 'pool:painted frame'), None)
        rim_purple_cell = next((c for c in (rim_table or {}).get('cells', []) if c['color'] == 'Purple'), None)
        frame_purple_cell = next((c for c in (frame_table or {}).get('cells', []) if c['color'] == 'Purple'), None)
        check(rim_purple_cell is not None and rim_purple_cell['value'] == '6', f"Rim's Purple cell stays 6, unaffected by Frame's override (got {rim_purple_cell})")
        check(frame_purple_cell is not None and frame_purple_cell['value'] == '9', f"Frame's Purple cell shows its own overridden 9, not summed/overwritten (got {frame_purple_cell})")

        if console_errors:
            print("\n⚠️  Console/page errors:")
            for e in console_errors:
                print(f"    {e}")
            failures.append("console errors present")

        print("\n" + ("ALL TESTS PASSED" if not failures else f"{len(failures)} TEST(S) FAILED"))
        browser.close()
        return len(failures) == 0


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
