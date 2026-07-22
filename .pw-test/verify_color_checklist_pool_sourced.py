"""
Same as verify_color_checklist_full_list.py but for a process whose color
sub-groups come from the Warehouse Pool (a multi-color upstream item), not
from explicit per-color recipe rows — this path goes through
renderGroupedColorChecklist's pool-item-grouping fallback (see
getPoolColorAwareItemNames), not the simpler flat-list path.

Supersedes the old repro_color_checklist_pool_sourced.py, which asserted
the checklist must stay scoped to just this pool item's 2 stocked colors —
that was the deliberate pre-2026-07-22 behavior; see
verify_color_checklist_full_list.py's docstring for why it was relaxed.
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
    "processId": "PRC-2", "processName": "Frame Assembly", "sequence": 2,
    "lotPrefix": "FA", "outputItemName": "Assembled Frame", "isFinalStage": False,
    "active": True, "processType": "General"
}

MOCK_COMPONENTS = [
    {"itemName": "Screw", "size": "", "sourceType": "ITEM", "qtyPerUnit": 4, "colorGroup": "COMMON"},
    {"itemName": "Painted Frame", "size": "General", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
]

MOCK_ITEMS = [
    {"name": "Screw", "size": ""},
]

# Warehouse Pool has stock of "Painted Frame" in 2 colors (Blue, Orange-White) —
# these 2 are still expected to render (grouped under "Painted Frame"), but
# the rest of Color Master must also render now, in a trailing "Other" bucket.
MOCK_POOL = [
    {"outputItemName": "Painted Frame", "color": "Blue", "processId": "PRC-1", "qty": 10},
    {"outputItemName": "Painted Frame", "color": "Orange-White", "processId": "PRC-1", "qty": 10},
]

# Color Master has 5 colors total; only 2 are actually in stock for this pool item.
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
    # the 2 pool-detected colors unioned with the full Color Master list.
    "getProcessColorGroups": {"success": True, "data": sorted(ALL_COLOR_NAMES)},
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

        print("[Step 1] Open Create modal, select the pool-sourced-color process...")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)

        page.evaluate("""
            document.getElementById('productionProcessId').value = 'PRC-2';
            App.Production.handleProcessChange('PRC-2');
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
            print("  ✅ Checklist shows the full Color Master list, not just this pool item's 2 stocked colors.")
        else:
            print("  ❌ Checklist did not widen to the full Color Master list as expected.")

        if console_errors:
            print("\n⚠️  Console/page errors:")
            for e in console_errors:
                print(f"    {e}")

        browser.close()
        return ok


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
