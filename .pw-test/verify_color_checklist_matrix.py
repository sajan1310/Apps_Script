"""
Verification script: Production Lot form's Color checklist + Per-Color
Components matrix.

Reproduces the user's report:
  1. Checking a color's checkbox in the checklist should enable its qty
     text field — verify whether it actually does.
  2. Checked colors' quantities should feed into the Common Components
     suggested-qty total — verify the total adds up.
  3. The Per-Color Components matrix table layout (for CSS/width work).

Api.call is monkey-patched (no google.script.run available outside Apps
Script) to return canned data for a process with 2 color sub-groups and a
mix of Common + per-color recipe components, mirroring the real shape.
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

        # `Api` (the google.script.run wrapper) is a closure-private const,
        # not reachable from here — but window.google.script.run IS the
        # thing it calls into, and the app's own built-in local-preview mock
        # (Script.html's "MOCK ENVIRONMENT" block) already installed a
        # window.google before this script ever runs. Replace it with a
        # fuller mock covering the methods this flow needs.
        page.evaluate(f"""
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
            App.State.globalItems = {json.dumps(MOCK_ITEMS)};
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

        print("\n[Step 1] Open Create modal, select the process...")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)

        page.evaluate("""
            document.getElementById('productionProcessId').value = 'PRC-1';
            App.Production.handleProcessChange('PRC-1');
        """)
        page.wait_for_timeout(800)

        rows = page.locator("#productionColorChecklist .production-color-row")
        print(f"  Color checklist rows found: {rows.count()}")
        assert rows.count() == 2, f"Expected 2 color rows, got {rows.count()}"

        blue_row = page.locator("#productionColorChecklist .production-color-row[data-color='Blue']")
        assert blue_row.count() == 1, "Blue row not found"

        qty_input = blue_row.locator(".production-color-qty")
        before = qty_input.is_disabled()
        print(f"\n[Step 2] Blue qty field disabled BEFORE checking checkbox: {before}")
        assert before, "Expected qty field to start disabled"

        checkbox = blue_row.locator(".production-color-check")
        checkbox.check()
        page.wait_for_timeout(500)

        after = qty_input.is_disabled()
        print(f"  Blue qty field disabled AFTER checking checkbox: {after}")
        if after:
            print("  ❌ BUG CONFIRMED: qty field stayed disabled after checking the checkbox.")
        else:
            print("  ✅ qty field became enabled.")

        if not after:
            print("\n[Step 3] Type qty=5 for Blue, check Orange-White, type qty=3, verify totals...")
            qty_input.fill("5")
            page.wait_for_timeout(300)

            ow_row = page.locator("#productionColorChecklist .production-color-row[data-color='Orange-White']")
            ow_checkbox = ow_row.locator(".production-color-check")
            ow_checkbox.check()
            page.wait_for_timeout(500)
            ow_qty = ow_row.locator(".production-color-qty")
            print(f"  Orange-White qty field disabled after check: {ow_qty.is_disabled()}")
            ow_qty.fill("3")
            page.wait_for_timeout(500)

            checked = page.evaluate("App.Production.getCheckedColorQtys()")
            print(f"  getCheckedColorQtys() -> {checked}")
            total = sum(c["qty"] for c in checked)
            print(f"  Total checked qty: {total} (expected 8)")
            assert total == 8, f"Expected total 8, got {total}"

            # Common Components table: Brush is COMMON, qtyPerUnit 1 -> suggested qty should be 8
            brush_row = page.locator("#productionComponentsBody tr").filter(has=page.locator(".prod-comp-qty"))
            qty_values = page.evaluate("""
                Array.from(document.querySelectorAll('#productionComponentsBody .prod-comp-qty')).map(i => i.value)
            """)
            print(f"  Common Components suggested qty values: {qty_values}")

            # Per-Color matrix: Blue column qty for Frame---Blue should be 5, Orange-White column for Frame---Orange-White should be 3
            matrix_info = page.evaluate("""
                (() => {
                    const headerColors = App.Production.getMatrixColors().map(c => c.color);
                    const rows = Array.from(document.querySelectorAll('#productionColorMatrixBody tr')).map(row => {
                        const name = row.querySelector('.prod-comp-display-name')?.value || '(non-merged)';
                        const cellVals = Array.from(row.querySelectorAll('.matrix-qty')).map(i => i.value);
                        return { name, cellVals };
                    });
                    return { headerColors, rows };
                })()
            """)
            print(f"  Matrix header colors: {matrix_info['headerColors']}")
            for r in matrix_info["rows"]:
                print(f"    Row '{r['name']}': {r['cellVals']}")

        print("\n[Step 4] Inspect Per-Color matrix table layout (widths)...")
        layout = page.evaluate("""
            (() => {
                const table = document.querySelector('#productionColorMatrixWrapper table');
                const wrapper = document.querySelector('#productionColorMatrixWrapper .table-responsive');
                const header = document.getElementById('productionColorMatrixHeaderRow');
                const ths = Array.from(header.children).map(th => ({
                    text: th.textContent.trim(),
                    width: th.getBoundingClientRect().width
                }));
                return {
                    tableWidth: table.getBoundingClientRect().width,
                    wrapperWidth: wrapper.getBoundingClientRect().width,
                    scrollWidth: wrapper.scrollWidth,
                    needsScroll: wrapper.scrollWidth > wrapper.clientWidth + 1,
                    columns: ths
                };
            })()
        """)
        print(f"  Table width: {layout['tableWidth']:.0f}px, wrapper width: {layout['wrapperWidth']:.0f}px, scrollWidth: {layout['scrollWidth']:.0f}px")
        print(f"  Needs horizontal scroll: {layout['needsScroll']}")
        for c in layout["columns"]:
            print(f"    Column '{c['text']}': {c['width']:.0f}px")

        page.screenshot(path=str(Path(__file__).parent / "color_matrix_before.png"), full_page=True)
        print("\n  Screenshot saved: .pw-test/color_matrix_before.png")

        if console_errors:
            print("\n⚠️  Console/page errors:")
            for e in console_errors:
                print(f"    {e}")

        browser.close()
        return not after


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
