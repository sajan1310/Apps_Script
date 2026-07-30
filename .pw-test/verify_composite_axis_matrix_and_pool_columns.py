"""
Verification script: a 2-axis process whose PRIMARY axis's pool colors are
already COMPOSITES ("Blue-White / BCP") — the real "Fitting Frame" shape.

Two reported bugs, both visible in one screenshot of the Log Production Lot
form after checking 4 primary composite colors at 15 each:

  Problem 1 — Per-Color Components rendered 4 EXTRA, permanently-empty
  columns (RED / PURPLE / PINK / BLUE) next to the 4 real composite ones.
  Those come from the non-primary Mudguard axis rows that the primary
  cascade auto-checks (_syncMatchingNonPrimaryRows); they describe the SAME
  physical units as their primary counterpart, and their real consumption
  lives in their own Per-Process Pool Components table, so a duplicate
  column here is pure clutter at best and a double stock debit at worst.
  Fixed by _pruneRedundantMatrixColumns — which must still LEAVE a
  non-primary column alone when the recipe genuinely populated it (a recipe
  row tagged colorGroup "Purple" only ever matches a plain "Purple" column,
  never the composite — see _matchedColorToken). Step 4 covers that.

  Problem 2 — the Frame axis's Per-Process Pool Components table rendered a
  single "PINK-WHITE" column instead of its 4 checked composite ones.
  syncPoolColorGroupColumns derived visible columns by TOKEN matching only,
  so checking "Blue-White / BCP" produced tokens {blue-white, bcp}, neither
  of which equals the full column string — every composite column was
  dropped, and only a legacy plain column that happened to equal one token
  ("Pink-White") survived. Since that resync runs after every checkbox
  toggle, it overwrote the correct initial render. Fixed by reusing
  _checkedPoolGroupColors (exact-wins-then-tokens), the same rule the first
  render already used.

Run: python .pw-test/verify_composite_axis_matrix_and_pool_columns.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

MUDGUARD = "14 inch Broad Mudguard Painted Rib"
FRAME = "Fitted Frame 14 inch Crysta Steel Rim"

MOCK_PROCESS = {
    "processId": "PRC-1", "processName": "Fitting Frame", "sequence": 2,
    "lotPrefix": "FF", "outputItemName": "Fitted Cycle", "isFinalStage": False,
    "active": True, "processType": "General"
}

# PRC-2 is identical except the recipe ALSO carries a genuinely
# mudguard-color-specific non-pool item, tagged with the plain color "Purple".
MOCK_PROCESS_2 = dict(MOCK_PROCESS, processId="PRC-2", processName="Fitting Frame (w/ Purple sticker)")

BASE_COMPONENTS = [
    {"itemName": "CRYSTA Side Flap", "size": "GENERAL", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "COMMON"},
    {"itemName": MUDGUARD, "size": "", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
    {"itemName": FRAME, "size": "", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
]
PURPLE_SPECIFIC = {"itemName": "Purple Mudguard Sticker", "size": "GENERAL", "sourceType": "ITEM",
                   "qtyPerUnit": 2, "colorGroup": "Purple"}

MUDGUARD_COLORS = ["Blue", "Pink", "Purple", "Red"]
FRAME_COLORS = ["Blue-White / BCP", "Pink-White", "Pink-White / BCP",
                "Purple-White / BCP", "Red-White / BCP"]
# The 4 the operator actually checks (all primary, 15 each) — "Pink-White"
# (the plain legacy bucket) is deliberately left UNCHECKED, since that is
# precisely the column the old token matching wrongly showed on its own.
CHECKED_PRIMARY = ["Blue-White / BCP", "Pink-White / BCP",
                   "Purple-White / BCP", "Red-White / BCP"]

MOCK_POOL_ROWS = (
    [{"outputItemName": MUDGUARD, "color": c, "processId": "PRC-M", "qty": 1000} for c in MUDGUARD_COLORS]
    + [{"outputItemName": FRAME, "color": c, "processId": "PRC-F", "qty": 50} for c in FRAME_COLORS]
)

MOCK_AXES = {
    "axes": [
        {"key": "pool:mudguard", "label": MUDGUARD, "colors": MUDGUARD_COLORS, "source": "pool"},
        {"key": "pool:frame", "label": FRAME, "colors": FRAME_COLORS, "source": "pool"},
    ],
    "primaryColorAxis": FRAME,
    "primaryAxisKey": "pool:frame",
}

MOCK_ITEMS = [
    {"name": "CRYSTA Side Flap", "size": "GENERAL"},
    {"name": "Purple Mudguard Sticker", "size": "GENERAL"},
]


def responses(components):
    return {
        "getProcessColorGroups": {"success": True, "data": MUDGUARD_COLORS + FRAME_COLORS},
        "getProcessColorAxes": {"success": True, "data": MOCK_AXES},
        "getProcessComponentsData": {"success": True, "data": components},
        "getWarehousePoolData": {"success": True, "data": MOCK_POOL_ROWS},
        "getProcessWipData": {"success": True, "data": []},
        "getStockData": {"success": True, "data": []},
        "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
    }


def matrix_columns(page):
    return page.evaluate("""
        Array.from(document.querySelectorAll('#productionColorMatrixHeaderRow th[data-color]'))
             .map(th => th.dataset.color)
    """)


def pool_tables(page):
    return page.evaluate("""
        Array.from(document.querySelectorAll('#productionPoolColorGroupsContainer table')).map(t => ({
            axisKey: t.dataset.axisKey,
            columns: Array.from(t.querySelectorAll('thead th[data-color]')).map(th => th.dataset.color),
            cells: Array.from(t.querySelectorAll('.pool-group-qty')).map(i => ({color: i.dataset.color, value: i.value}))
        }))
    """)


def open_process(page, process_id):
    page.evaluate("App.Production.openCreateModal()")
    page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
    page.evaluate(f"""
        document.getElementById('productionProcessId').value = '{process_id}';
        App.Production.handleProcessChange('{process_id}');
    """)
    page.wait_for_timeout(900)


def check_primaries(page):
    """Check the 4 primary composites and set each to 15, one at a time —
    exactly how the operator did it (each click fires its own cascade)."""
    for color in CHECKED_PRIMARY:
        page.evaluate(f"""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="{color}"] .production-color-check').click()
        """)
        page.wait_for_timeout(200)
        page.locator(f'#productionColorChecklist .production-color-row[data-color="{color}"] .production-color-qty').fill("15")
        page.wait_for_timeout(200)
    page.wait_for_timeout(400)


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_context().new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        page.evaluate(f"""
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}, {json.dumps(MOCK_PROCESS_2)}];
            App.State.globalItems = {json.dumps(MOCK_ITEMS)};
            App.State.globalColors = [];
            window.__mock = {{
                'PRC-1': {json.dumps(responses(BASE_COMPONENTS))},
                'PRC-2': {json.dumps(responses(BASE_COMPONENTS + [PURPLE_SPECIFIC]))}
            }};
            window.__activeProcess = 'PRC-1';
            window.google = {{ script: {{ run: {{
                withSuccessHandler(cb) {{
                    const runner = {{ withFailureHandler() {{ return runner; }} }};
                    Object.keys(window.__mock['PRC-1']).forEach(method => {{
                        runner[method] = () => setTimeout(() => cb(window.__mock[window.__activeProcess][method]), 30);
                    }});
                    return runner;
                }}
            }} }} }};
        """)

        failures = []

        def check(cond, msg):
            print(("PASS: " if cond else "FAIL: ") + msg)
            if not cond:
                failures.append(msg)

        print("[Step 1] Open Log Production Lot on the 2-pool-axis process (Mudguard + Frame, Frame primary)...")
        open_process(page, "PRC-1")
        axes_rendered = page.evaluate("""
            Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row'))
                 .map(r => r.dataset.group + '|' + r.dataset.color + '|' + r.dataset.primary)
        """)
        print("  checklist rows:", json.dumps(axes_rendered, indent=None))
        check(any(r.startswith("pool:frame|Blue-White / BCP|true") for r in axes_rendered),
              "Frame axis rendered as the primary group with composite colors")

        print("\n[Step 2] Check the 4 primary composites at 15 each (Mudguard rows auto-cascade)...")
        check_primaries(page)

        cascaded = page.evaluate("""
            Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row[data-group="pool:mudguard"]'))
                 .filter(r => r.querySelector('.production-color-check').checked)
                 .map(r => r.dataset.color + '=' + r.querySelector('.production-color-qty').value)
        """)
        print("  auto-checked mudguard rows:", cascaded)
        check(sorted(cascaded) == ["Blue=15", "Pink=15", "Purple=15", "Red=15"],
              f"All 4 matching Mudguard rows auto-checked at 15 (the real cascade feature) (got {cascaded})")

        print("\n[Problem 1] Per-Color Components must show ONLY the 4 composite columns...")
        cols = matrix_columns(page)
        print("  matrix columns:", cols)
        check(sorted(cols) == sorted(CHECKED_PRIMARY),
              f"Matrix has exactly the 4 checked composite columns, no duplicate plain ones (got {cols})")
        for dup in MUDGUARD_COLORS:
            check(dup not in cols, f"No redundant empty '{dup}' column (duplicate of its composite counterpart)")

        print("\n[Problem 2] Frame pool table must show all 4 checked composite columns, not just 'Pink-White'...")
        tables = pool_tables(page)
        print("  pool tables:", json.dumps(tables))
        frame = next((t for t in tables if t["axisKey"] == "pool:frame"), None)
        mud = next((t for t in tables if t["axisKey"] == "pool:mudguard"), None)
        check(frame is not None and sorted(frame["columns"]) == sorted(CHECKED_PRIMARY),
              f"Frame table shows its 4 checked composite columns (got {frame and frame['columns']})")
        check(frame is not None and "Pink-White" not in frame["columns"],
              "Frame table does NOT show the unchecked plain 'Pink-White' column")
        check(frame is not None and all(c["value"] == "15" for c in frame["cells"]),
              f"Every Frame composite cell suggests 15 (got {frame and frame['cells']})")
        check(mud is not None and sorted(mud["columns"]) == sorted(MUDGUARD_COLORS)
              and all(c["value"] == "15" for c in mud["cells"]),
              f"Mudguard table still shows its own 4 plain columns at 15 (got {mud})")

        print("\n[Step 3] Uncheck one primary composite — its column leaves BOTH tables, the rest stay...")
        page.evaluate("""
            document.querySelector('#productionColorChecklist .production-color-row[data-color="Red-White / BCP"] .production-color-check').click()
        """)
        page.wait_for_timeout(600)
        cols = matrix_columns(page)
        frame = next((t for t in pool_tables(page) if t["axisKey"] == "pool:frame"), None)
        check("Red-White / BCP" not in cols and sorted(cols) == sorted(CHECKED_PRIMARY[:3]),
              f"Matrix drops only the unchecked composite (got {cols})")
        check(frame is not None and sorted(frame["columns"]) == sorted(CHECKED_PRIMARY[:3]),
              f"Frame pool table drops only the unchecked composite (got {frame and frame['columns']})")

        print("\n[Step 4] A non-primary color the recipe REALLY has a per-color item for must keep its column...")
        page.evaluate("App.Nav && App.Nav.reset && App.Nav.reset(); window.__activeProcess = 'PRC-2';")
        page.evaluate("document.getElementById('editProductionModal') && bootstrap.Modal.getInstance(document.getElementById('editProductionModal')).hide()")
        page.wait_for_timeout(500)
        open_process(page, "PRC-2")
        check_primaries(page)
        cols = matrix_columns(page)
        print("  matrix columns (PRC-2):", cols)
        purple_cell_has_qty = page.evaluate("""
            (() => {
                const ths = Array.from(document.querySelectorAll('#productionColorMatrixHeaderRow th'));
                const idx = ths.findIndex(th => (th.dataset.color || '').toLowerCase() === 'purple');
                if (idx === -1) return null;
                return Array.from(document.querySelectorAll('#productionColorMatrixBody tr'))
                    .map(r => (r.children[idx]?.querySelector('.matrix-qty') || {}).value || '')
                    .filter(Boolean);
            })()
        """)
        print("  Purple column quantities:", purple_cell_has_qty)
        check("Purple" in cols,
              f"Non-primary 'Purple' column SURVIVES — the recipe has a real Purple-only item for it (got {cols})")
        check(purple_cell_has_qty and any(v for v in purple_cell_has_qty),
              f"...and it is genuinely populated, not blank (got {purple_cell_has_qty})")
        for dup in ["Blue", "Pink", "Red"]:
            check(dup not in cols, f"'{dup}' still pruned — nothing in the recipe populates it")

        check(not console_errors, f"No uncaught page errors (got {console_errors})")

        browser.close()

        print("\n" + "=" * 60)
        if failures:
            print(f"{len(failures)} FAILURE(S):")
            for f in failures:
                print("  - " + f)
            sys.exit(1)
        print("ALL CHECKS PASSED")


if __name__ == "__main__":
    run()
