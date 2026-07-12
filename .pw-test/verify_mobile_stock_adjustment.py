"""
Verification script — Mobile Stock Adjustment write flow (Stock tab > expand
an item > "Adjust stock").

This is the first WRITE action added to the mobile Stock screen (previously
read-only: list + per-item movement history). Mirrors desktop's
App.Stock.handleAdjustSubmit / adjustStockModal (View_Stock.html) exactly:
same 3 fields (item label, current value, corrected value + reason), same
server call (module_stock.js#adjustStockManually, unchanged), same
"negative values allowed" exception for stock corrections.

Covers: opening the sheet pre-fills the right item/value, client-side
validation (missing reason, non-numeric value) blocks the API call, a
successful save closes the sheet + shows a success toast + reloads the
list, and a rejected save keeps the sheet open with an error toast so the
user can retry without re-entering everything.

Run: python .pw-test/verify_mobile_stock_adjustment.py
"""
import sys
import io
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "mobile.html"

failures = 0
def check(cond, msg):
    global failures
    if cond:
        print(f"  PASS: {msg}")
    else:
        failures += 1
        print(f"  FAIL: {msg}")


# Any method not explicitly set in __mockRoutes defaults to an empty success
# response -- lets _ensureLedgerSources' 6-way Promise.all (bills/returns/
# wastage/issues/production/adjustment history) resolve without each test
# having to stub every one of them individually.
MOCK_RUNNER_JS = """
    window.__mockRoutes = {};
    window.__callCounts = {};
    window.google = {
        script: {
            run: {
                withSuccessHandler(successCb) {
                    const state = { successCb, failureCb: null };
                    let proxy;
                    const runner = {
                        withFailureHandler(failureCb) {
                            state.failureCb = failureCb;
                            return proxy;
                        }
                    };
                    proxy = new Proxy(runner, { get(target, prop) {
                        if (prop in target) return target[prop];
                        return (...args) => {
                            window.__callCounts[prop] = (window.__callCounts[prop] || 0) + 1;
                            window.__lastArgs = window.__lastArgs || {};
                            window.__lastArgs[prop] = args;
                            setTimeout(() => {
                                const resp = (prop in window.__mockRoutes) ? window.__mockRoutes[prop] : { success: true, data: [] };
                                if (resp && resp.__throwError) {
                                    if (state.failureCb) state.failureCb(new Error(resp.__throwError));
                                } else {
                                    successCb(resp);
                                }
                            }, 60);
                        };
                    }});
                    return proxy;
                }
            }
        }
    };
"""

SAMPLE_STOCK = """
    ({
      success: true,
      data: [
        { name: 'Brake Pads', size: 'Standard', initialStock: 100, currentStock: 5, threshold: 20, isLowStock: true, deadStock: false },
        { name: 'Chain Link', size: '', initialStock: 50, currentStock: 40, threshold: 10, isLowStock: false, deadStock: false }
      ]
    })
"""
SAMPLE_ITEMS = """
    ({
      success: true,
      data: [
        { name: 'Brake Pads', size: 'Standard', narration: '', baseUnit: 'Pcs', purchaseUnit: 'Pcs', weightPerBaseUnit: 0, vendors: [] },
        { name: 'Chain Link', size: '', narration: '', baseUnit: 'Pcs', purchaseUnit: 'Box', weightPerBaseUnit: 0, vendors: [] }
      ]
    })
"""


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context()
        page = ctx.new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(300)
        page.evaluate(MOCK_RUNNER_JS)
        page.evaluate(f"window.__mockRoutes.getStockData = {SAMPLE_STOCK};")
        page.evaluate(f"window.__mockRoutes.getItemsData = {SAMPLE_ITEMS};")

        print("\n[Navigate] Stock tab loads both items")
        page.evaluate("MApp.Shell.showTab('stock')")
        page.wait_for_timeout(300)
        check(page.locator("#stock-list .mb-card").count() == 2, "2 stock cards rendered")

        print("\n[Expand] tapping Brake Pads reveals its movement panel with an 'Adjust stock' action")
        page.click("#stock-list .mb-card >> nth=0")
        page.wait_for_timeout(250)
        adjust_btn = page.locator("#stock-expand-0 >> text=Adjust stock")
        check(adjust_btn.count() == 1, "an 'Adjust stock' button is present in the expanded panel")

        print("\n[Open sheet] pre-fills item label and current value from the tapped item")
        adjust_btn.click()
        page.wait_for_timeout(50)
        sheet_open = page.evaluate("document.getElementById('sheet-stock-adjust').classList.contains('open')")
        check(sheet_open, "sheet-stock-adjust is opened (has 'open' class)")
        item_label = page.evaluate("document.getElementById('stock-adjust-item-label').value")
        check(item_label == 'Brake Pads (Standard)', f"item label pre-filled correctly (got {item_label!r})")
        old_val = page.evaluate("document.getElementById('stock-adjust-old-value').value")
        check(str(old_val) == '5', f"current value shown as 5 (got {old_val!r})")
        new_val_default = page.evaluate("document.getElementById('stock-adjust-new-value').value")
        check(str(new_val_default) == '5', f"corrected value input defaults to the current value (got {new_val_default!r})")

        print("\n[Validation] blank reason blocks the save, no API call fires")
        page.fill("#stock-adjust-new-value", "-3")
        page.fill("#stock-adjust-reason", "")
        page.click("#stock-adjust-save-btn")
        page.wait_for_timeout(50)
        calls_before = page.evaluate("window.__callCounts.adjustStockManually || 0")
        check(calls_before == 0, f"adjustStockManually NOT called yet (blank reason) (got {calls_before} call(s))")
        error_toast = page.locator("#mapp-toast-stack .mb-toast-error").last
        check("reason" in (error_toast.text_content() or "").lower(), f"toast asks for a reason (got {error_toast.text_content()!r})")

        print("\n[Validation] an empty corrected-value field also blocks the save")
        page.fill("#stock-adjust-reason", "Physical recount")
        page.fill("#stock-adjust-new-value", "")
        page.click("#stock-adjust-save-btn")
        page.wait_for_timeout(50)
        calls_after_bad_number = page.evaluate("window.__callCounts.adjustStockManually || 0")
        check(calls_after_bad_number == 0, f"adjustStockManually still not called (empty value) (got {calls_after_bad_number} call(s))")

        print("\n[Negative value allowed] a negative corrected value is NOT blocked client-side (intentional exception)")
        page.evaluate("window.__mockRoutes.adjustStockManually = { success: true, message: 'Stock adjusted successfully.' };")
        page.fill("#stock-adjust-new-value", "-3")
        page.fill("#stock-adjust-reason", "Physical recount found a shortfall")
        page.click("#stock-adjust-save-btn")
        page.wait_for_timeout(250)
        args = page.evaluate("window.__lastArgs.adjustStockManually")
        check(args == ['Brake Pads', 'Standard', -3, 'Physical recount found a shortfall'],
              f"adjustStockManually called with (name, size, -3, reason) (got {args!r})")

        print("\n[Success] sheet closes, success toast shown, list reloads")
        sheet_closed = page.evaluate("!document.getElementById('sheet-stock-adjust').classList.contains('open')")
        check(sheet_closed, "sheet-stock-adjust closes on success")
        success_toast = page.locator("#mapp-toast-stack .mb-toast-success").last
        check("adjusted" in (success_toast.text_content() or "").lower(), f"success toast shown (got {success_toast.text_content()!r})")
        stock_reload_calls = page.evaluate("window.__callCounts.getStockData || 0")
        check(stock_reload_calls == 2, f"getStockData re-fetched after a successful adjustment (got {stock_reload_calls} call(s))")

        print("\n[Failure] a rejected adjustment keeps the sheet open with an error toast (so the user can retry)")
        # render() re-opens the previously-expanded panel automatically after a
        # reload (matched by item key, see MApp.Stock.render()'s tail) -- the
        # panel for Brake Pads is therefore already expanded at this point,
        # no need to tap the card again.
        page.wait_for_timeout(350)
        page.click("#stock-expand-0 >> text=Adjust stock")
        page.wait_for_timeout(50)
        page.evaluate("window.__mockRoutes.adjustStockManually = { success: false, message: 'New stock value is the same as the current value — nothing to adjust.' };")
        page.fill("#stock-adjust-new-value", "5")
        page.fill("#stock-adjust-reason", "Testing failure path")
        page.click("#stock-adjust-save-btn")
        page.wait_for_timeout(250)
        sheet_still_open = page.evaluate("document.getElementById('sheet-stock-adjust').classList.contains('open')")
        check(sheet_still_open, "sheet-stock-adjust stays open after a failed adjustment (nothing lost)")
        fail_toast = page.locator("#mapp-toast-stack .mb-toast-error").last
        check("nothing to adjust" in (fail_toast.text_content() or "").lower(), f"server's rejection message shown verbatim (got {fail_toast.text_content()!r})")
        save_btn_reenabled = page.evaluate("!document.getElementById('stock-adjust-save-btn').disabled")
        check(save_btn_reenabled, "Save button re-enabled after a failed save")

        print("\n[Cancel] closing the sheet manually discards the in-progress edit")
        page.evaluate("MApp.Stock.closeAdjustSheet()")
        page.wait_for_timeout(50)
        check(page.evaluate("!document.getElementById('sheet-stock-adjust').classList.contains('open')"), "sheet-stock-adjust closes cleanly via Cancel/close")

        if console_errors:
            print("\nConsole/page errors:")
            for e in console_errors:
                print(f"    {e}")
            globals()['failures'] += len(console_errors)

        browser.close()


if __name__ == "__main__":
    run()
    print(f"\n{'ALL PASS' if failures == 0 else str(failures) + ' FAILURE(S)'}")
    sys.exit(0 if failures == 0 else 1)
