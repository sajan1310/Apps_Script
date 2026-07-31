"""
Reopening a saved Production lot whose colors are COMPOSITES
("Blue-White / Black") must NOT render extra, permanently-blank Per-Process
Pool Components columns for each composite's SEGMENTS ("Blue-White",
"Black") — none of those plain colors was ever checked, and the segment
column can never hold a value because _poolGroupCellValue's exact branch
wins for the composite. renderPoolColorGroupsFromAccum used raw token
matching (_checkedColorTokensLower) where syncPoolColorGroupColumns used
exact-wins (_checkedPoolGroupColors), so a 6-composite lot opened with 13
columns and collapsed back to 6 on the first checkbox toggle.

Step 2 pins that agreement: the initial render and a later resync must
produce the identical column set.

Run: python .pw-test/verify_pool_table_composite_segment_columns.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

POOL_ITEM = "Fitted Frame 20 inch Crysta S/Rim"

# Exactly the checklist from the screenshot: plain colors, "/ BCP"
# composites, and the "/ Black" composites the lot actually produced.
POOL_COLORS = [
    "Black", "Blue", "Blue-White", "Blue-White / BCP", "Blue-White / Black",
    "Orange", "Orange-White", "Orange-White / BCP", "Orange-White / Black",
    "Pink", "Pink-White", "Pink-White / BCP", "Pink-White / Black",
    "Purple", "Purple-White", "Purple-White / BCP", "Purple-White / Black",
    "Red", "Red-White", "Red-White / BCP", "Red-White / Black",
    "SeaGreen", "SeaGreen-White", "SeaGreen-White / BCP", "SeaGreen-White / Black",
]

PRODUCED = [
    ("Blue-White / Black", 5), ("Orange-White / Black", 5), ("Pink-White / Black", 8),
    ("Purple-White / Black", 7), ("Red-White / Black", 10), ("SeaGreen-White / Black", 5),
]

MOCK_PROCESS = {
    "processId": "PRC-FF", "processName": "Fitting Frame", "sequence": 6,
    "lotPrefix": "FF", "outputItemName": "Fitted Frame Assembled", "isFinalStage": False,
    "active": True, "processType": "General", "primaryColorAxis": POOL_ITEM,
}
MOCK_COMPONENTS = [
    {"itemName": POOL_ITEM, "size": "", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
]
MOCK_POOL = [{"outputItemName": POOL_ITEM, "color": c, "processId": "PRC-UP", "qty": 30} for c in POOL_COLORS]
MOCK_AXES = {
    "success": True,
    "data": {
        "axes": [{"key": "pool:" + POOL_ITEM.lower(), "label": POOL_ITEM, "colors": POOL_COLORS, "source": "pool"}],
        "primaryColorAxis": POOL_ITEM,
        "primaryAxisKey": "pool:" + POOL_ITEM.lower(),
    },
}

SAVED_LOT = {
    "lotNumber": "FF-0001", "processId": "PRC-FF", "status": "Completed",
    "qty": sum(q for _, q in PRODUCED), "color": PRODUCED[0][0], "assignedTo": "Sanjay",
    "colorBreakdown": [{"color": c, "qty": q} for c, q in PRODUCED],
    "componentsConsumed": [
        {"itemName": POOL_ITEM, "size": "", "sourceType": "POOL", "qty": q, "colorGroup": c}
        for c, q in PRODUCED
    ],
}

MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": POOL_COLORS},
    "getProcessColorAxes": MOCK_AXES,
    "getProcessComponentsData": {"success": True, "data": MOCK_COMPONENTS},
    "getWarehousePoolData": {"success": True, "data": MOCK_POOL},
    "getProcessWipData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
}

READ_COLS = """
    (() => Array.from(document.querySelectorAll(
        '#productionPoolColorGroupsContainer table.prod-color-table thead th[data-color]'))
        .map(th => ({ color: th.dataset.color })))()
"""
READ_CELLS = """
    (() => Array.from(document.querySelectorAll(
        '#productionPoolColorGroupsContainer .pool-group-qty'))
        .map(i => [i.dataset.color, i.value]))()
"""


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_context().new_page()
        errors = []
        page.on("pageerror", lambda e: errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        page.evaluate(f"""
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
            App.State.globalItems = [];
            App.State.globalColors = [];
            App.State.globalContractors = [{{ contractorId: 'C1', name: 'Sanjay', active: true }}];
            App.State.globalProduction = [{json.dumps(SAVED_LOT)}];
            window.__mockResponses = {json.dumps(MOCK_API_RESPONSES)};
            window.google = {{ script: {{ run: {{
                withSuccessHandler(cb) {{
                    const runner = {{ withFailureHandler() {{ return runner; }} }};
                    Object.keys(window.__mockResponses).forEach(m => {{
                        runner[m] = () => setTimeout(() => cb(window.__mockResponses[m]), 20);
                    }});
                    return runner;
                }}
            }} }} }};
        """)

        ok = True
        print("[Step 1] Open Edit Lot for the saved composite-color lot")
        page.evaluate("App.Production.openEditModal(0)")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(1200)

        cols = [c["color"] for c in page.evaluate(READ_COLS)]
        cells = dict(page.evaluate(READ_CELLS))
        produced = [c for c, _ in PRODUCED]
        blanks = [c for c in cols if not cells.get(c)]
        print(f"  columns ({len(cols)}): {cols}")
        print(f"  blank columns ({len(blanks)}): {blanks}")

        if blanks:
            print(f"  FAIL: {len(blanks)} never-selected column(s) rendered blank")
            ok = False
        else:
            print("  PASS: no segment columns on open")
        if sorted(cols) != sorted(produced):
            print(f"  FAIL: expected exactly the {len(produced)} produced composites, got {len(cols)}")
            ok = False
        else:
            print("  PASS: columns are exactly this lot's produced composites")

        print("\n[Step 2] Toggle a checkbox to force a column resync")
        page.evaluate("""
            (() => {
                const cb = document.querySelector('#productionColorChecklist input.production-color-check:not(:checked)');
                if (cb) { cb.checked = true; cb.dispatchEvent(new Event('change', { bubbles: true })); }
                if (cb) { cb.checked = false; cb.dispatchEvent(new Event('change', { bubbles: true })); }
            })()
        """)
        page.wait_for_timeout(800)
        cols2 = [c["color"] for c in page.evaluate(READ_COLS)]
        print(f"  columns after resync ({len(cols2)}): {cols2}")
        if cols2 != cols:
            print("  FAIL: resync disagrees with the initial render (inconsistent rule)")
            ok = False
        else:
            print("  PASS: resync agrees with the initial render")

        if errors:
            print("\n  Page errors:")
            for e in errors:
                print(f"    {e}")
            ok = False

        browser.close()
        return ok


if __name__ == "__main__":
    ok = run()
    print("\n" + ("ALL PASS" if ok else "SOME FAILED"))
    sys.exit(0 if ok else 1)
