"""
Verification script: Warehouse Pool sub-tab bulk-print feature (Stock tab).

Checks:
  1. Each Process row in the Warehouse Pool table has a checkbox.
  2. "Select All" checks every row and shows "Print Selected (N)".
  3. Unchecking one row updates the count and label.
  4. Clicking a row's checkbox does NOT open the process breakdown modal;
     clicking elsewhere on the row DOES open it.
  5. Clicking "Print Selected" populates the shared print pivot container
     with rows for only the selected processes' Warehouse Pool buckets.

Run: python .pw-test/verify_warehouse_pool_bulk_print.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

MOCK_PROCESSES = [
    {
        "processId": "PRC-1", "processName": "Painting Frame 16 inch", "sequence": 1,
        "lotPrefix": "PF", "outputItemName": "Painted Frame 16 inch", "isFinalStage": False,
        "active": True, "processType": "General"
    },
    {
        "processId": "PRC-2", "processName": "Fitting Rim 16 inch", "sequence": 2,
        "lotPrefix": "FR", "outputItemName": "Fitted Rim 16 inch", "isFinalStage": False,
        "active": True, "processType": "General"
    }
]
MOCK_POOL_ROWS = [
    {"outputItemName": "Painted Frame 16 inch", "processId": "PRC-1", "productTag": "", "color": "Red", "producedQty": 100, "consumedQty": 20, "availableQty": 80},
    {"outputItemName": "Painted Frame 16 inch", "processId": "PRC-1", "productTag": "", "color": "Blue", "producedQty": 50, "consumedQty": 10, "availableQty": 40},
    {"outputItemName": "Fitted Rim 16 inch", "processId": "PRC-2", "productTag": "", "color": "", "producedQty": 30, "consumedQty": 5, "availableQty": 25},
]
MOCK_API_RESPONSES = {
    "getProcessData": {"success": True, "data": MOCK_PROCESSES},
    "getWarehousePoolData": {"success": True, "data": MOCK_POOL_ROWS},
    "getAllProcessColorGroups": {"success": True, "data": {}},
    "getStockData": {"success": True, "data": []},
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

        print("\n[Step 1] Navigate to Stock tab > Warehouse Pool sub-tab...")
        page.evaluate("App.Navigation.showTab('stockTab')")
        page.wait_for_timeout(500)
        page.locator("#btn-warehousePoolSubTab").click()
        page.wait_for_timeout(800)

        rows = page.locator("#warehousePoolTableBody tr.pool-process-row")
        check(rows.count() == 2, f"2 process rows rendered (got {rows.count()})")

        chks = page.locator(".warehousepool-select-chk")
        check(chks.count() == 2, f"2 row checkboxes rendered (got {chks.count()})")

        print_btn = page.locator("#btnBulkPrintWarehousePool")
        check("d-none" in (print_btn.get_attribute("class") or ""), "Print Selected button hidden before any selection")

        print("\n[Step 2] Click Select All...")
        page.locator("#warehousePoolSubTab").locator("button:has-text('Select All')").click()
        page.wait_for_timeout(300)

        check("d-none" not in (print_btn.get_attribute("class") or ""), "Print Selected button visible after Select All")
        check("(2)" in print_btn.inner_text(), f"Print Selected button shows count 2 (got '{print_btn.inner_text()}')")
        check(chks.nth(0).is_checked() and chks.nth(1).is_checked(), "Both row checkboxes checked after Select All")

        print("\n[Step 3] Uncheck one row...")
        chks.nth(0).uncheck()
        page.wait_for_timeout(300)
        check("(1)" in print_btn.inner_text(), f"Print Selected button shows count 1 after unchecking one row (got '{print_btn.inner_text()}')")

        print("\n[Step 4] Clicking a row checkbox must NOT open the breakdown modal...")
        chks.nth(0).check()
        page.wait_for_timeout(300)
        modal_visible = page.locator("#warehousePoolProcessModal").is_visible()
        check(not modal_visible, "Breakdown modal stayed closed after clicking a row's checkbox")

        print("\n[Step 5] Clicking elsewhere on the row DOES open the breakdown modal...")
        rows.nth(0).locator("td").nth(2).click()
        page.locator("#warehousePoolProcessModal").wait_for(state="visible", timeout=TIMEOUT)
        check(page.locator("#warehousePoolProcessModal").is_visible(), "Breakdown modal opened after clicking row body")
        page.locator("#warehousePoolProcessModal .btn-close").click()
        page.wait_for_timeout(500)

        print("\n[Step 6] Click Print Selected and inspect the pivot table...")
        # Re-select both rows (state persists, but confirm before printing).
        check(chks.nth(0).is_checked() and chks.nth(1).is_checked(), "Both rows still selected before printing")
        # window.print() is a no-op in headless Chromium and fires 'afterprint'
        # synchronously, so App.Print.trigger()'s cleanup (which clears
        # 'active-print' and resets the title) can already have run by the
        # time this evaluate() call returns — check the populated content
        # instead of the transient active-print class.
        page.evaluate("App.Stock.bulkPrintWarehousePool()")

        subtitle = page.evaluate("document.getElementById('print-low-stock-subtitle').innerText")
        check(subtitle == 'Selected Warehouse Pool Register', f"Subtitle reads 'Selected Warehouse Pool Register' (got '{subtitle}')")

        body_rows = page.evaluate("document.getElementById('print-low-stock-body').innerText")
        print(f"  Print body text:\n{body_rows}")
        check("Painted Frame 16 inch" in body_rows, "Printout includes Painted Frame pool rows")
        check("Fitted Rim 16 inch" in body_rows, "Printout includes Fitted Rim pool row")
        check("[Red]" in body_rows and "[Blue]" in body_rows, "Printout breaks out Red/Blue color buckets as separate rows")

        if console_errors:
            print("\n⚠️  Console/page errors:")
            for e in console_errors:
                print(f"    {e}")

        print("\n" + ("ALL TESTS PASSED" if not failures else f"{len(failures)} TEST(S) FAILED"))
        browser.close()
        return len(failures) == 0


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
