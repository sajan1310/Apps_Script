"""
Verification script — App.Utils.getPendingByItem() / App.Bill._getBilledQty()
performance refactor (Script_Core.html / Script_Bill.html).

_getBilledQty used to rescan every Bill x every Bill line on EVERY call, and
getPendingByItem calls it once per PO line across every open PO. It's now
backed by a Map built once from App.State.globalBills and cached by array
identity (rebuilt only when globalBills is reassigned, e.g. after a reload).
This test proves: (1) the aggregated pending-by-item numbers are unchanged,
(2) excludeBillNumber still works, and (3) the cache correctly picks up a
reassigned globalBills instead of serving stale data.

Run: python .pw-test/verify_pending_by_item_perf.py
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


SEED_PO_AND_BILLS = """
    App.State.globalPOs = [
      { poNumber: 'TEST-PO-1', vendor: 'Vendor A',
        items: [ { name: 'Widget', size: '', narration: '', baseQty: 10 } ] },
      { poNumber: 'TEST-PO-2', vendor: 'Vendor B',
        items: [ { name: 'Gadget', size: 'L', narration: '', baseQty: 20 } ] }
    ];
    App.State.globalBills = [
      { billNumber: 'BILL-1', vendor: 'Vendor A',
        items: [ { poNumber: 'TEST-PO-1', name: 'Widget', size: '', narration: '', baseQty: 4 } ] },
      { billNumber: 'BILL-2', vendor: 'Vendor A',
        items: [ { poNumber: 'TEST-PO-1', name: 'Widget', size: '', narration: '', baseQty: 3 } ] }
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

        print("\n[Seed] 2 open PO lines (Widget qty 10, Gadget qty 20); Widget billed 4+3=7 across 2 bills")
        page.evaluate(SEED_PO_AND_BILLS)

        print("\n=== getPendingByItem aggregation ===")
        pending_widget = page.evaluate("""
            (() => {
              const map = App.Utils.getPendingByItem();
              const e = map.get('widget|');
              return e ? { qty: e.qty, poCount: e.poNumbers.size } : null;
            })()
        """)
        check(pending_widget is not None, "Widget entry present in pending map")
        if pending_widget:
            check(abs(pending_widget["qty"] - 3) < 0.001, f"Widget pending qty is 10-7=3 (got {pending_widget['qty']})")
            check(pending_widget["poCount"] == 1, f"Widget pending tracked against 1 PO (got {pending_widget['poCount']})")

        pending_gadget = page.evaluate("""
            (() => {
              const map = App.Utils.getPendingByItem();
              const e = map.get('gadget|l');
              return e ? { qty: e.qty } : null;
            })()
        """)
        check(pending_gadget is not None, "Gadget entry present (no bills against it)")
        if pending_gadget:
            check(abs(pending_gadget["qty"] - 20) < 0.001, f"Gadget pending qty is 20 (got {pending_gadget['qty']})")

        print("\n=== _getBilledQty direct + excludeBillNumber ===")
        billed_total = page.evaluate("App.Bill._getBilledQty('TEST-PO-1', 'Widget', '', '')")
        check(abs(billed_total - 7) < 0.001, f"billed total across both bills is 7 (got {billed_total})")

        billed_excl = page.evaluate("App.Bill._getBilledQty('TEST-PO-1', 'Widget', '', '', 'BILL-1')")
        check(abs(billed_excl - 3) < 0.001, f"excluding BILL-1 leaves 3 (BILL-2 only) (got {billed_excl})")

        print("\n=== Cache invalidation on globalBills reassignment ===")
        # Prime the cache, then reassign globalBills entirely (simulating a
        # reload) with a DIFFERENT billed qty for the same PO line, and
        # confirm the cache picks up the new array instead of serving stale
        # totals from the old one.
        page.evaluate("App.Bill._getBilledQty('TEST-PO-1', 'Widget', '', '');")  # prime cache
        page.evaluate("""
            App.State.globalBills = [
              { billNumber: 'BILL-3', vendor: 'Vendor A',
                items: [ { poNumber: 'TEST-PO-1', name: 'Widget', size: '', narration: '', baseQty: 1 } ] }
            ];
        """)
        billed_after_reassign = page.evaluate("App.Bill._getBilledQty('TEST-PO-1', 'Widget', '', '')")
        check(abs(billed_after_reassign - 1) < 0.001,
              f"cache reflects reassigned globalBills, not stale (got {billed_after_reassign})")

        check(len(console_errors) == 0, f"no page errors ({console_errors})")

        browser.close()


if __name__ == "__main__":
    run()
    print("\nALL TESTS PASSED" if failures == 0 else f"\n{failures} TEST(S) FAILED")
    sys.exit(1 if failures else 0)
