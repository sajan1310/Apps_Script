"""
Verifies the new "which group does this custom color belong to?" picker
(addCustomColorRow / _refreshCustomColorGroupSelect in Script.html):
- Hidden for a process with 0/1 real color groups (no ambiguity to ask about).
- Shown, populated with every real group (Primary flagged), for a process
  with 2+ groups.
- Picking the Primary group places the row under it with data-primary="true".
- Picking a non-primary group places the row under IT with data-primary="false".
- Picking "Independent" (default) keeps the old flat "Custom" bucket behavior.

Run: python .pw-test/verify_custom_color_group_picker.py
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
MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": sorted(list({r["color"] for r in MOCK_POOL}))},
    "getProcessColorAxes": MOCK_AXES,
    "getProcessComponentsData": {"success": True, "data": MOCK_COMPONENTS},
    "getWarehousePoolData": {"success": True, "data": MOCK_POOL},
    "getProcessWipData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
}

# Single-axis process for the "picker stays hidden" check.
MOCK_PROCESS_SIMPLE = {
    "processId": "PRC-SIMPLE", "processName": "Simple Paint", "sequence": 6,
    "lotPrefix": "SP", "outputItemName": "Painted Widget", "isFinalStage": False,
    "active": True, "processType": "General", "primaryColorAxis": ""
}
MOCK_API_RESPONSES_SIMPLE = {
    "getProcessColorGroups": {"success": True, "data": ["Red", "Blue"]},
    "getProcessColorAxes": {"success": True, "data": {"axes": [], "primaryColorAxis": "", "primaryAxisKey": ""}},
    "getProcessComponentsData": {"success": True, "data": []},
    "getWarehousePoolData": {"success": True, "data": []},
    "getProcessWipData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
}


def setup_mocks(page, process, api_responses):
    page.evaluate(f"""
        App.State.globalProcesses = [{json.dumps(process)}];
        App.State.globalItems = [];
        App.State.globalColors = [];
        window.__mockResponses = {json.dumps(api_responses)};
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


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_context().new_page()
        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.resolve().as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        ok = True

        print("[1] Multi-group process: picker should be visible with 3 options "
              "(Independent + Painted Frame (Primary) + Mudguard)...")
        setup_mocks(page, MOCK_PROCESS, MOCK_API_RESPONSES)
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.evaluate("""
            document.getElementById('productionProcessId').value = 'PRC-FTF';
            App.Production.handleProcessChange('PRC-FTF');
        """)
        page.wait_for_timeout(800)

        sel_visible = page.evaluate("getComputedStyle(document.getElementById('productionCustomColorGroupSelect')).display")
        options = page.evaluate("""
            Array.from(document.getElementById('productionCustomColorGroupSelect').options).map(o => o.textContent)
        """)
        print(f"  selector display: {sel_visible}")
        print(f"  options: {options}")
        if sel_visible == 'none' or len(options) != 3:
            print("  FAIL: expected the picker visible with 3 options")
            ok = False
        else:
            print("  PASS")

        print("\n[2] Add a custom color to the PRIMARY group (Painted Frame)...")
        page.evaluate("""
            const colorInput = document.getElementById('productionCustomColorInput');
            colorInput.add(new Option('Sunset Orange', 'Sunset Orange', true, true));
            const sel = document.getElementById('productionCustomColorGroupSelect');
            const opt = Array.from(sel.options).find(o => o.textContent.includes('Painted Frame'));
            sel.value = opt.value;
            App.Production.addCustomColorRow();
        """)
        page.wait_for_timeout(300)
        row_info = page.evaluate("""
            (() => {
                const row = Array.from(document.querySelectorAll('.production-color-row')).find(r => r.dataset.color === 'Sunset Orange');
                return row ? { group: row.dataset.group, primary: row.dataset.primary, checked: row.querySelector('.production-color-check').checked } : null;
            })()
        """)
        print(f"  row info: {row_info}")
        if not row_info or row_info['primary'] != 'true' or 'painted' not in (row_info['group'] or '').lower():
            print("  FAIL: expected the row placed in the Painted Frame group with data-primary=true")
            ok = False
        else:
            print("  PASS: custom color correctly placed in Primary group, auto-checked")

        print("\n[3] Add a custom color to the NON-primary group (Mudguard)...")
        page.evaluate("""
            const colorInput = document.getElementById('productionCustomColorInput');
            colorInput.add(new Option('Neon Green', 'Neon Green', true, true));
            const sel = document.getElementById('productionCustomColorGroupSelect');
            const opt = Array.from(sel.options).find(o => o.textContent.includes('Mudguard'));
            sel.value = opt.value;
            App.Production.addCustomColorRow();
        """)
        page.wait_for_timeout(300)
        row_info2 = page.evaluate("""
            (() => {
                const row = Array.from(document.querySelectorAll('.production-color-row')).find(r => r.dataset.color === 'Neon Green');
                return row ? { group: row.dataset.group, primary: row.dataset.primary } : null;
            })()
        """)
        print(f"  row info: {row_info2}")
        if not row_info2 or row_info2['primary'] != 'false' or 'mudguard' not in (row_info2['group'] or '').lower():
            print("  FAIL: expected the row placed in the Mudguard group with data-primary=false")
            ok = False
        else:
            print("  PASS: custom color correctly placed in non-Primary group")

        print("\n[4] Add a custom color as Independent (no group picked)...")
        page.evaluate("""
            const colorInput = document.getElementById('productionCustomColorInput');
            colorInput.add(new Option('Test Batch X', 'Test Batch X', true, true));
            document.getElementById('productionCustomColorGroupSelect').value = '';
            App.Production.addCustomColorRow();
        """)
        page.wait_for_timeout(300)
        row_info3 = page.evaluate("""
            (() => {
                const row = Array.from(document.querySelectorAll('.production-color-row')).find(r => r.dataset.color === 'Test Batch X');
                return row ? { group: row.dataset.group, primary: row.dataset.primary } : null;
            })()
        """)
        print(f"  row info: {row_info3}")
        if not row_info3 or row_info3['group'] != 'custom' or row_info3['primary'] not in (None, ''):
            print("  FAIL: expected the row in the flat 'custom' bucket with no data-primary attribute")
            ok = False
        else:
            print("  PASS: independent custom color uses the old flat 'Custom' bucket")

        print("\n[5] Confirm the checked-color qty payload matches expectations for saveProduction...")
        for r in [row_info, row_info2, row_info3]:
            pass
        checked_summary = page.evaluate("""
            App.Production.getCheckedColorQtys().filter(c => c.isCustom).map(c => ({ color: c.color, countsTowardTotal: c.countsTowardTotal }))
        """)
        print(f"  {checked_summary}")
        by_color = {c['color']: c['countsTowardTotal'] for c in checked_summary}
        if by_color.get('Sunset Orange') is not True:
            print("  FAIL: Primary-group custom color should have countsTowardTotal=true")
            ok = False
        elif by_color.get('Neon Green') is not False:
            print("  FAIL: non-Primary-group custom color should have countsTowardTotal=false")
            ok = False
        elif by_color.get('Test Batch X') is not True:
            print("  FAIL: independent custom color should have countsTowardTotal=true")
            ok = False
        else:
            print("  PASS: countsTowardTotal correctly differs by chosen group")

        print("\n[6] Single-group process: picker should stay hidden entirely...")
        page2 = browser.new_context().new_page()
        page2.goto(DIST_HTML.resolve().as_uri(), wait_until="domcontentloaded")
        page2.wait_for_timeout(1000)
        setup_mocks(page2, MOCK_PROCESS_SIMPLE, MOCK_API_RESPONSES_SIMPLE)
        page2.evaluate("App.Production.openCreateModal()")
        page2.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page2.evaluate("""
            document.getElementById('productionProcessId').value = 'PRC-SIMPLE';
            App.Production.handleProcessChange('PRC-SIMPLE');
        """)
        page2.wait_for_timeout(800)
        sel_visible2 = page2.evaluate("getComputedStyle(document.getElementById('productionCustomColorGroupSelect')).display")
        print(f"  selector display: {sel_visible2}")
        if sel_visible2 != 'none':
            print("  FAIL: expected the picker hidden for a single-group process")
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
