"""
Verifies the 2026-07-22 "always show every Color Master color" change: the
"Colors to Produce" checklist on the Production form must offer the FULL
Color Master list once a process has any color sub-group at all, not just
the colors this specific process's own recipe/pool history happens to have
touched so far (see computeColorGroupsForProcess in module_process.js).

Supersedes the old repro_color_checklist_full_list.py, which asserted the
OPPOSITE (that the checklist must stay scoped to just the process's own 2
colors) — that was the deliberate pre-2026-07-22 behavior; the user asked
for it to be relaxed specifically because a not-yet-produced/not-yet-tagged
Color Master color was invisible on the form. This script mocks
getProcessColorGroups as the real (now-widened) backend would answer: the
process's own 2 tag-based colors plus the rest of a 5-color Color Master.

App.State.globalColors is seeded with 5 colors; getProcessColorGroups is
mocked to return all 5 (simulating the widened backend response) for a
process whose recipe only explicitly tags 2 of them. All 5 must render.
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
    "processId": "PRC-1", "processName": "Frame Painting", "sequence": 1,
    "lotPrefix": "FP", "outputItemName": "Painted Frame", "isFinalStage": False,
    "active": True, "processType": "General"
}

MOCK_COMPONENTS = [
    {"itemName": "Brush", "size": "", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "COMMON"},
    {"itemName": "Frame---Blue", "size": "General", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "Blue"},
    {"itemName": "Frame---Orange-White", "size": "General", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "Orange-White"},
]

MOCK_ITEMS = [
    {"name": "Brush", "size": ""},
    {"name": "Frame---Blue", "size": "General"},
    {"name": "Frame---Orange-White", "size": "General"},
]

# Color Master has 5 colors total; the process's recipe only explicitly
# tags 2 of them ("Blue", "Orange-White") — the other 3 must still show up
# on the checklist now that the backend always widens to the full master
# list once a process is color-enabled at all.
MOCK_GLOBAL_COLORS = [
    {"name": "Blue", "remarks": ""},
    {"name": "Orange-White", "remarks": ""},
    {"name": "Red", "remarks": ""},
    {"name": "Green", "remarks": ""},
    {"name": "Black", "remarks": ""},
]
ALL_COLOR_NAMES = {c["name"] for c in MOCK_GLOBAL_COLORS}

MOCK_API_RESPONSES = {
    # Simulates the real (post-2026-07-22) computeColorGroupsForProcess:
    # the 2 recipe-tagged colors unioned with the full Color Master list.
    "getProcessColorGroups": {"success": True, "data": sorted(ALL_COLOR_NAMES)},
    "getProcessComponentsData": {"success": True, "data": MOCK_COMPONENTS},
    "getWarehousePoolData": {"success": True, "data": []},
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
            App.State.globalColors = {json.dumps(MOCK_GLOBAL_COLORS)};
            window.__mockResponses = {json.dumps(MOCK_API_RESPONSES)};
            window.google = {{
                script: {{
                    run: {{
                        withSuccessHandler(cb) {{
                            const runner = {{
                                withFailureHandler() {{ return runner; }}
                            }};
                            Object.keys(window.__mockResponses).forEach(method => {{
                                runner[method] = (...args) => setTimeout(() => cb(window.__mockResponses[method]), 50);
                            }});
                            return runner;
                        }}
                    }}
                }}
            }};
        """)

        print("[Step 1] Open Create modal, select the process (recipe tags 2 of 5 Color Master colors)...")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)

        page.evaluate("""
            document.getElementById('productionProcessId').value = 'PRC-1';
            App.Production.handleProcessChange('PRC-1');
        """)
        page.wait_for_timeout(800)

        rows = page.locator("#productionColorChecklist .production-color-row")
        colors_shown = page.evaluate("""
            Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row'))
                .map(r => r.dataset.color)
        """)
        print(f"  Rows rendered: {rows.count()}")
        print(f"  Colors shown: {colors_shown}")
        print(f"  Color Master total size: {len(MOCK_GLOBAL_COLORS)}")

        ok = rows.count() == len(ALL_COLOR_NAMES) and set(colors_shown) == ALL_COLOR_NAMES
        if ok:
            print("  ✅ Checklist shows the full Color Master list, not just this process's 2 tagged colors.")
        else:
            print("  ❌ Checklist did not widen to the full Color Master list as expected.")

        addColorsBtnVisible = page.evaluate("""
            (() => { const el = document.getElementById('productionAddColorsBtn'); return el && el.style.display !== 'none'; })()
        """)
        print(f"  '+ Add Colors to this Lot' button visible: {addColorsBtnVisible} (should be False when process already has color sub-groups)")
        ok = ok and not addColorsBtnVisible

        if console_errors:
            print("\n⚠️  Console/page errors:")
            for e in console_errors:
                print(f"    {e}")

        browser.close()
        return ok


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
