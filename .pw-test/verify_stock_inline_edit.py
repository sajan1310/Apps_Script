"""
Verifies the new click-to-edit Current Stock cell on the Stock tab's Items
sub-tab (Script_Stock.html#App.Stock.editStockCell), added so correcting
stock no longer requires the "Correct stock manually" modal for quick edits
-- same interaction as the existing Warehouse Pool table's click-to-edit
Available Qty cell (editPoolStockCell).

Covers:
  1. Clicking the Current Stock value swaps it for a number input, prefilled
     with the current value.
  2. Committing a new value (blur) calls adjustStockManually(name, size,
     newQty, <fixed reason>) and, on success, updates the cell text, the
     row's low/well-stocked color+badge, and the Alerts panel -- without a
     full table reload.
  3. Escape reverts to the original display without calling the server.
  4. A same-value server rejection (stale on-screen number) reconciles the
     cell to the server's authoritative value instead of getting stuck.

Run: python .pw-test/verify_stock_inline_edit.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

ITEM_A = {"name": "Brake Pads", "size": "Standard", "initialStock": 50, "currentStock": 40, "threshold": 10, "isLowStock": False, "deadStock": False}
ITEM_B = {"name": "Chain Link", "size": "GENERAL", "initialStock": 20, "currentStock": 15, "threshold": 20, "isLowStock": True, "deadStock": False}

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
        "getStockData": {"success": True, "data": [ITEM_A, ITEM_B]},
        "getWarehousePoolData": {"success": True, "data": []},
        "adjustStockManually": {"success": True, "message": "Stock adjusted successfully.", "data": {}},
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


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context()
        page = ctx.new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        url = DIST_HTML.as_uri()
        page.goto(url, wait_until="domcontentloaded")
        page.wait_for_timeout(500)

        setup_mocks(page)
        page.click("#btn-stockTab")
        page.wait_for_timeout(300)

        row_a = page.locator("#stockTableBody tr").nth(0)
        span_a = row_a.locator(".stock-current-display")
        check(span_a.count() == 1, "Current Stock cell renders as a click-to-edit span")
        check(span_a.inner_text().strip() == "40", "Span shows the initial Current Stock value (40)")

        print("\n[1] Click swaps the span for a number input prefilled with the current value")
        span_a.click()
        input_a = row_a.locator("td").nth(4).locator("input[type=number]")
        check(input_a.count() == 1, "A number input replaces the span")
        check(input_a.input_value() == "40", "Input is prefilled with 40")

        print("\n[2] Committing a new value calls adjustStockManually and updates the row")
        input_a.fill("55")
        input_a.press("Enter")
        page.wait_for_timeout(200)

        args = page.evaluate("window.__lastArgs.adjustStockManually")
        check(args == ["Brake Pads", "Standard", 55, "Inline edit via Stock table"],
              f"adjustStockManually called with (name, size, 55, fixed reason) (got {args!r})")

        row_a = page.locator("#stockTableBody tr").nth(0)
        cell_a_text = row_a.locator(".stock-current-display").inner_text().strip()
        check(cell_a_text == "55", f"Cell now displays 55 (got {cell_a_text!r})")
        toast_msg = page.locator("#toastMessage").inner_text()
        check(toast_msg == "Stock updated.", f"Success toast shown (got {toast_msg!r})")

        print("\n[3] Escape reverts without calling the server")
        calls_before = page.evaluate("window.__callCounts.adjustStockManually || 0")
        span_a = row_a.locator(".stock-current-display")
        span_a.click()
        input_a = row_a.locator("td").nth(4).locator("input[type=number]")
        input_a.fill("999")
        input_a.press("Escape")
        page.wait_for_timeout(150)
        calls_after = page.evaluate("window.__callCounts.adjustStockManually || 0")
        check(calls_after == calls_before, "Escape does not call adjustStockManually")
        row_a = page.locator("#stockTableBody tr").nth(0)
        cell_a_text = row_a.locator(".stock-current-display").inner_text().strip()
        check(cell_a_text == "55", f"Cell still shows 55 after Escape (got {cell_a_text!r})")

        print("\n[4] Editing a low-stock item that crosses back above threshold updates color/badge + Alerts")
        row_b = page.locator("#stockTableBody tr").nth(1)
        alert_before = page.locator("#lowStockAlertSection").inner_text()
        check("Chain Link" in alert_before, "Chain Link starts in the Low Stock alert panel")

        span_b = row_b.locator(".stock-current-display")
        span_b.click()
        input_b = row_b.locator("td").nth(4).locator("input[type=number]")
        input_b.fill("25")
        input_b.press("Enter")
        page.wait_for_timeout(200)

        row_b = page.locator("#stockTableBody tr").nth(1)
        td_b = row_b.locator("td").nth(4)
        check("text-success" in (td_b.get_attribute("class") or ""), "Chain Link's cell is now text-success (above threshold)")
        alert_after = page.locator("#lowStockAlertSection").inner_text()
        check("Chain Link" not in alert_after, "Chain Link no longer appears in the Low Stock alert panel")

        print("\n[5] A same-value server rejection reconciles to the server's authoritative value")
        setup_mocks(page, {
            "adjustStockManually": {"success": False, "data": {"oldCurrentStock": 33, "newCurrentStock": 33}, "message": "New stock value is the same as the current value — nothing to adjust."}
        })
        row_b = page.locator("#stockTableBody tr").nth(1)
        span_b = row_b.locator(".stock-current-display")
        span_b.click()
        input_b = row_b.locator("td").nth(4).locator("input[type=number]")
        input_b.fill("40")
        input_b.press("Enter")
        page.wait_for_timeout(200)
        row_b = page.locator("#stockTableBody tr").nth(1)
        cell_b_text = row_b.locator(".stock-current-display").inner_text().strip()
        check(cell_b_text == "33", f"Cell reconciles to the server's authoritative value 33 (got {cell_b_text!r})")

        check(len(console_errors) == 0, f"No uncaught page errors (got {console_errors})")

        print(f"\n{'='*60}")
        if failures == 0:
            print("ALL CHECKS PASSED")
        else:
            print(f"{failures} CHECK(S) FAILED")
        browser.close()
        return failures


if __name__ == "__main__":
    sys.exit(1 if run() > 0 else 0)
