"""
Verification script — App.Dispatch.populateDispatchOrderSelect() performance
refactor (Script_Dispatch.html). Dispatched qty per (orderNumber, productId)
is now pre-aggregated into a Map once, instead of filtering the full
globalDispatch array per order line inside the nested orders x lines loop.
Confirms: pending-qty math is unchanged, an order line that's fully
dispatched is hidden, and editing an existing dispatch still excludes its
OWN row from the "already dispatched" total (so its current order stays
selectable/correctly pending).

Run: python .pw-test/verify_dispatch_order_select_perf.py
"""
import sys
import io
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"

failures = 0
def check(cond, msg):
    global failures
    if cond:
        print(f"  PASS: {msg}")
    else:
        failures += 1
        print(f"  FAIL: {msg}")


SEED = """
    App.State.globalOrders = [
      { orderNumber: 'ORD-1', clientName: 'Client A', status: 'Order Confirmed',
        lines: [
          { productId: 'P1', productName: 'Widget', qty: 10 },
          { productId: 'P2', productName: 'Gadget', qty: 5 }
        ] },
      { orderNumber: 'ORD-2', clientName: 'Client B', status: 'Order Confirmed',
        lines: [ { productId: 'P3', productName: 'Thingamajig', qty: 3 } ] }
    ];
    App.State.globalDispatch = [
      { rowIdx: 10, orderNumber: 'ORD-1', productId: 'P1', qty: 4 },
      { rowIdx: 11, orderNumber: 'ORD-1', productId: 'P2', qty: 5 },  // fully dispatched
      { rowIdx: 12, orderNumber: 'ORD-2', productId: 'P3', qty: 1 }
    ];
"""


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context()
        page = ctx.new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(500)

        page.evaluate("""
            window.google = {
                script: {
                    run: {
                        withSuccessHandler(cb) {
                            const runner = { withFailureHandler() { return runner; } };
                            return new Proxy(runner, { get(target, prop) {
                                if (prop in target) return target[prop];
                                return (...args) => {};
                            }});
                        }
                    }
                }
            };
        """)

        print("\n[Seed] 2 orders (3 lines total), partial dispatch history")
        page.evaluate(SEED)

        print("\n=== populateDispatchOrderSelect (no row being edited) ===")
        page.evaluate("""
            document.getElementById('dispatchRowIdx').value = '';
            App.Dispatch.populateDispatchOrderSelect('');
        """)
        options = page.locator("#dispatchOrderSelect option")
        texts = [options.nth(i).inner_text() for i in range(options.count())]
        joined = " | ".join(texts)

        check(any("ORD-1" in t and "Widget" in t and "Pending: 6" in t for t in texts),
              f"ORD-1/Widget shows pending 6 (10-4) (options: {joined})")
        check(not any("ORD-1" in t and "Gadget" in t for t in texts),
              f"ORD-1/Gadget (fully dispatched, 5-5=0) is hidden (options: {joined})")
        check(any("ORD-2" in t and "Pending: 2" in t for t in texts),
              f"ORD-2/Thingamajig shows pending 2 (3-1) (options: {joined})")

        print("\n=== populateDispatchOrderSelect (editing dispatch rowIdx=10 itself) ===")
        page.evaluate("""
            document.getElementById('dispatchRowIdx').value = '10';
            App.Dispatch.populateDispatchOrderSelect('ORD-1|P1');
        """)
        options2 = page.locator("#dispatchOrderSelect option")
        texts2 = [options2.nth(i).inner_text() for i in range(options2.count())]
        joined2 = " | ".join(texts2)
        check(any("ORD-1" in t and "Widget" in t and "Pending: 10" in t for t in texts2),
              f"editing rowIdx=10 excludes its own 4-unit dispatch, pending back to 10 (options: {joined2})")

        check(len(console_errors) == 0, f"no page errors ({console_errors})")

        browser.close()


if __name__ == "__main__":
    run()
    print("\nALL TESTS PASSED" if failures == 0 else f"\n{failures} TEST(S) FAILED")
    sys.exit(1 if failures else 0)
