"""
Verification script — Mobile Items Lookup stock-on-hand enhancement (More
tab > Items lookup sheet).

sheet-items-lookup already existed (getItemsData only); this verifies the
new stock-on-hand + low-stock cross-reference added on top of it
(module_stock.js#getStockData, joined client-side by (name, size)), plus
the "fail open" requirement -- a stock-load failure must not break the
item lookup itself, just omit the stock figure.

Run: python .pw-test/verify_mobile_item_lookup_stock.py
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


# Routes each mocked method independently via window.__mockRoutes[method], so
# getItemsData and getStockData can be configured (and made to fail
# independently of each other) within the same page session.
MOCK_RUNNER_JS = """
    window.__mockRoutes = {};
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
                            setTimeout(() => {
                                const resp = window.__mockRoutes[prop];
                                if (resp && resp.__throwError) {
                                    if (state.failureCb) state.failureCb(new Error(resp.__throwError));
                                } else {
                                    successCb(resp);
                                }
                            }, 100);
                        };
                    }});
                    return proxy;
                }
            }
        }
    };
"""

SAMPLE_ITEMS = """
    ({
      success: true,
      data: [
        { name: 'Brake Pads', size: 'Standard', narration: '', baseUnit: 'Pcs', purchaseUnit: 'Pcs', weightPerBaseUnit: 0, vendors: [] },
        { name: 'Chain Link', size: '', narration: 'Heavy duty', baseUnit: 'Pcs', purchaseUnit: 'Box', weightPerBaseUnit: 0, vendors: [] }
      ]
    })
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

        print("\n[Success] items + stock both resolve -> stock-on-hand shown per row")
        page.evaluate(f"window.__mockRoutes.getItemsData = {SAMPLE_ITEMS};")
        page.evaluate(f"window.__mockRoutes.getStockData = {SAMPLE_STOCK};")
        page.evaluate("MApp.Items.openLookupSheet()")
        page.wait_for_timeout(400)

        card_count = page.locator("#items-lookup-list .mb-card").count()
        check(card_count == 2, f"2 item cards rendered (got {card_count})")

        card_texts = page.evaluate("""
            Object.fromEntries([...document.querySelectorAll('#items-lookup-list .mb-card')].map(card => {
                const name = card.querySelector('.mb-card-title').textContent.trim();
                return [name, card.textContent];
            }))
        """)
        check('5' in card_texts.get('Brake Pads', ''), f"Brake Pads shows its stock-on-hand (5) (got {card_texts.get('Brake Pads')!r})")
        check('Low stock' in card_texts.get('Brake Pads', ''), "Brake Pads (below threshold) shows the 'Low stock' chip")
        check('40' in card_texts.get('Chain Link', ''), f"Chain Link shows its stock-on-hand (40) (got {card_texts.get('Chain Link')!r})")
        check('Low stock' not in card_texts.get('Chain Link', ''), "Chain Link (above threshold) shows no 'Low stock' chip")

        print("\n[Fail-open] getStockData rejects -> items still render, just without a stock figure")
        page.evaluate(f"window.__mockRoutes.getItemsData = {SAMPLE_ITEMS};")
        page.evaluate("window.__mockRoutes.getStockData = { __throwError: 'Simulated stock failure' };")
        page.evaluate("MApp.Items.openLookupSheet()")
        page.wait_for_timeout(400)

        card_count2 = page.locator("#items-lookup-list .mb-card").count()
        check(card_count2 == 2, f"items still render when stock fetch fails (got {card_count2} cards)")
        no_stock_texts = page.evaluate("""
            [...document.querySelectorAll('#items-lookup-list .mb-card-number')].length
        """)
        check(no_stock_texts == 0, "no stock-on-hand figure is shown when the stock fetch failed (not a stale/wrong number)")

        print("\n[Failure] getItemsData itself rejects -> error state, not a blank/broken sheet")
        page.evaluate("window.__mockRoutes.getItemsData = { __throwError: 'Simulated items failure' };")
        page.evaluate("MApp.Items.openLookupSheet()")
        page.wait_for_timeout(400)
        check(page.locator("#items-lookup-list .mb-state-error").count() > 0, "error state renders when getItemsData itself fails")

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
