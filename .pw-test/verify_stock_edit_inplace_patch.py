"""
Verifies that editing a Stock row (Items sub-tab of the Stock tab) patches
ONLY that row instead of reloading the whole list -- so the page you're on,
your active search and your checkbox selection all survive the edit and you
don't have to hunt the item down again
(Script_Stock.html#App.Stock.patchRowInPlace).

Covers the three edit paths:
  1. "Correct stock manually" modal (handleAdjustSubmit) -- no getStockData
     refetch, stays on the current page, updates Current AND Initial Stock.
  2. Threshold input (saveThreshold) -- no getStockData refetch, recomputes
     the Low Stock badge/color + the Alerts panel in place.
  3. Inline click-to-edit Current Stock (editStockCell) -- keeps an active
     search + checkbox selection intact.

Run: python .pw-test/verify_stock_edit_inplace_patch.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"

# 45 items -> 3 pages at 20 rows/page, so a reset-to-page-1 reload is visible.
STOCK = [
    {
        "name": f"Item {i:02d}",
        "size": "GENERAL",
        "initialStock": 100 + i,
        "currentStock": 50 + i,
        "threshold": 5,
        "isLowStock": False,
        "deadStock": False,
    }
    for i in range(1, 46)
]

failures = 0


def check(cond, msg):
    global failures
    if cond:
        print(f"  PASS: {msg}")
    else:
        failures += 1
        print(f"  FAIL: {msg}")


def setup_mocks(page, extra_routes=None):
    routes = {
        "getStockData": {"success": True, "data": STOCK},
        "getWarehousePoolData": {"success": True, "data": []},
        "adjustStockManually": {"success": True, "message": "Stock adjusted successfully.", "data": {}},
        "updateThreshold": {"success": True, "message": "Threshold updated successfully."},
    }
    if extra_routes:
        routes.update(extra_routes)
    page.evaluate(f"""
        window.__mockRoutes = {json.dumps(routes)};
        window.__lastArgs = window.__lastArgs || {{}};
        window.__callCounts = window.__callCounts || {{}};
        window.google = {{
            script: {{
                run: {{
                    withSuccessHandler(cb) {{
                        const runner = {{ withFailureHandler() {{ return runner; }} }};
                        Object.keys(window.__mockRoutes).forEach(method => {{
                            runner[method] = (...args) => {{
                                window.__lastArgs[method] = args;
                                window.__callCounts[method] = (window.__callCounts[method] || 0) + 1;
                                setTimeout(() => cb(window.__mockRoutes[method]), 10);
                            }};
                        }});
                        return runner;
                    }}
                }}
            }}
        }};
    """)


def loads(page):
    return page.evaluate("window.__callCounts.getStockData || 0")


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_context().new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(500)

        setup_mocks(page)
        page.click("#btn-stockTab")
        page.wait_for_timeout(300)

        print("\n[1] 'Correct stock manually' modal patches the row without a reload")
        page.evaluate("App.Stock.changePage(2)")
        page.wait_for_timeout(150)
        first_row = page.locator("#stockTableBody tr").nth(0)
        check(first_row.locator("td").nth(1).inner_text().strip() == "Item 21",
              "Page 2 starts at Item 21")

        # Tag the live DOM nodes so we can prove the patch repaints only the
        # affected CELLS -- the <tr>, its checkbox and its threshold input
        # must all be the very same elements afterwards, never re-created.
        page.evaluate("""
            const tr = document.querySelectorAll('#stockTableBody tr')[0];
            tr.__probe = 'row';
            tr.querySelector('.stock-select-chk').__probe = 'chk';
            tr.querySelector('.threshold-input').__probe = 'thr';
        """)

        loads_before = loads(page)
        page.locator("#stockTableBody tr").nth(0).locator("button[title='Correct stock manually']").click()
        page.wait_for_timeout(300)
        check(page.input_value("#adjustStockOldValue") == "71",
              "Modal prefills the CURRENT qty (71) read from state")
        page.fill("#adjustStockNewValue", "999")
        page.fill("#adjustStockReason", "Physical count")
        page.click("#adjustStockSubmitBtn")
        page.wait_for_timeout(400)

        check(loads(page) == loads_before, "getStockData was NOT re-called (no full list reload)")
        probes = page.evaluate("""
            (() => {
                const tr = document.querySelectorAll('#stockTableBody tr')[0];
                return [tr.__probe, tr.querySelector('.stock-select-chk').__probe,
                        tr.querySelector('.threshold-input').__probe];
            })()
        """)
        check(probes == ["row", "chk", "thr"],
              f"Row, checkbox and threshold input are the SAME DOM nodes — only cells repainted (got {probes!r})")
        check(page.evaluate("App.State.stockCurrentPage") == 2, "Still on page 2 after the edit")
        first_row = page.locator("#stockTableBody tr").nth(0)
        check(first_row.locator("td").nth(1).inner_text().strip() == "Item 21",
              "Page 2 still starts at Item 21 (list not re-paged)")
        check(first_row.locator(".stock-current-display").inner_text().strip() == "999",
              "Edited row's Current Stock shows 999")
        # Initial Stock is back-solved server-side; the client mirrors the delta.
        # Item 21: initial 121, current 71 -> +928 -> 1049
        check(first_row.locator("td").nth(3).inner_text().strip() == "1049",
              f"Initial Stock moved by the same delta (expected 1049, got {first_row.locator('td').nth(3).inner_text().strip()!r})")
        check(page.evaluate("App.State.globalStock.find(i => i.name === 'Item 21').currentStock") == 999,
              "globalStock patched too")

        # Re-opening must show 999, not the value that was baked into the
        # button's markup when the row was first rendered.
        page.locator("#stockTableBody tr").nth(0).locator("button[title='Correct stock manually']").click()
        page.wait_for_timeout(300)
        check(page.input_value("#adjustStockOldValue") == "999",
              f"Re-opened modal shows the patched qty, not a stale one (got {page.input_value('#adjustStockOldValue')!r})")
        page.click("#adjustStockModal .btn-close")
        page.wait_for_timeout(300)

        print("\n[2] Threshold edit patches the badge + Alerts panel without a reload")
        loads_before = loads(page)
        thr = page.locator("#stockTableBody tr").nth(1).locator(".threshold-input")
        thr.fill("99999")
        thr.blur()
        page.wait_for_timeout(400)

        check(loads(page) == loads_before, "getStockData was NOT re-called for a threshold edit")
        check(page.evaluate("App.State.stockCurrentPage") == 2, "Still on page 2 after the threshold edit")
        row_b = page.locator("#stockTableBody tr").nth(1)
        check(row_b.locator("td").nth(1).inner_text().strip() == "Item 22", "Row 2 is still Item 22")
        check("Low Stock" in row_b.locator("td").nth(6).inner_text(),
              "Item 22 now shows the Low Stock badge")
        check("Item 22" in page.locator("#lowStockAlertSection").inner_text(),
              "Item 22 appears in the Low Stock alerts panel")
        check(row_b.locator(".threshold-input").input_value() == "99999",
              "Threshold input keeps the new value and is re-enabled")
        check(row_b.locator(".threshold-input").is_enabled(), "Threshold input is not left disabled")

        print("\n[3] Inline cell edit keeps an active search + selection")
        page.fill("#searchStock", "Item 33")
        page.locator("#searchStock").press("End")  # filterData is wired to onkeyup
        page.wait_for_timeout(400)
        check(page.locator("#stockTableBody tr").count() == 1, "Search narrows to a single row")
        page.locator("#stockTableBody tr").nth(0).locator(".stock-select-chk").check()
        page.wait_for_timeout(100)

        loads_before = loads(page)
        page.locator("#stockTableBody tr").nth(0).locator(".stock-current-display").click()
        cell_input = page.locator("#stockTableBody tr").nth(0).locator("td").nth(4).locator("input[type=number]")
        cell_input.fill("7")
        cell_input.press("Enter")
        page.wait_for_timeout(400)

        check(loads(page) == loads_before, "getStockData was NOT re-called for an inline edit")
        check(page.evaluate("App.State.stockSearchTerm") == "Item 33", "Search term survives the edit")
        check(page.locator("#stockTableBody tr").count() == 1, "Table still shows only the searched row")
        row_c = page.locator("#stockTableBody tr").nth(0)
        check(row_c.locator(".stock-current-display").inner_text().strip() == "7", "Cell shows the new value 7")
        check(row_c.locator(".stock-select-chk").is_checked(), "Row stays checked after the patch")
        check(page.evaluate("App.State.selectedStock.length") == 1, "Selection state preserved")
        check(page.evaluate("App.State.filteredStock.find(i => i.name === 'Item 33').currentStock") == 7,
              "filteredStock patched in sync with globalStock")

        check(len(console_errors) == 0, f"No uncaught page errors (got {console_errors})")

        print(f"\n{'='*60}")
        print("ALL CHECKS PASSED" if failures == 0 else f"{failures} CHECK(S) FAILED")
        browser.close()
        return failures


if __name__ == "__main__":
    sys.exit(1 if run() > 0 else 0)
