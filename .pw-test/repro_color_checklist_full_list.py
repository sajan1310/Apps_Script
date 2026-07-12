"""
Repro script: does the "Colors to Produce" checklist in the Production
form show every Color Master entry, instead of just the process's own
configured color sub-groups, when the process DOES have color sub-groups?

App.State.globalColors is seeded with 5 colors; getProcessColorGroups is
mocked to return only 2 of them ("Blue", "Orange-White") for the selected
process. If the checklist is correctly scoped, exactly 2 rows should
render. If the bug is real, 5 (or more) rows will render.
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

# Color Master has 5 colors total; the process only uses 2 of them.
MOCK_GLOBAL_COLORS = [
    {"name": "Blue", "remarks": ""},
    {"name": "Orange-White", "remarks": ""},
    {"name": "Red", "remarks": ""},
    {"name": "Green", "remarks": ""},
    {"name": "Black", "remarks": ""},
]

MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": ["Blue", "Orange-White"]},
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

        print("[Step 1] Open Create modal, select the process (which HAS 2 color sub-groups)...")
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

        bug = rows.count() != 2 or set(colors_shown) != {"Blue", "Orange-White"}
        if bug:
            print("  ❌ BUG REPRODUCED: checklist does not match the process's own 2 color sub-groups.")
        else:
            print("  ✅ Checklist correctly scoped to just this process's 2 color sub-groups.")

        addColorsBtnVisible = page.evaluate("""
            (() => { const el = document.getElementById('productionAddColorsBtn'); return el && el.style.display !== 'none'; })()
        """)
        print(f"  '+ Add Colors to this Lot' button visible: {addColorsBtnVisible} (should be False when process already has color sub-groups)")

        if console_errors:
            print("\n⚠️  Console/page errors:")
            for e in console_errors:
                print(f"    {e}")

        browser.close()
        return not bug


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
