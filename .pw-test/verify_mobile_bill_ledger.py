"""
Verification script — Mobile Bill Ledger (More tab > Bill Ledger sheet).

Same technique as verify_mobile_po_ledger.py: drives the compiled mobile
shell (dist/mobile.html) via Playwright, mocking google.script.run so
getBillData() responses are fully controlled from the test -- covering
loading/empty/success/failure states, search, and print-field population
(reusing the SAME #print-bill-container markup from View_Print.html the
desktop Bill Ledger populates). Bills have no status field, so there are no
filter chips to test here (unlike PO Ledger).

Run: python .pw-test/verify_mobile_bill_ledger.py
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


MOCK_RUNNER_JS = """
    window.__mockResponse = { success: true, data: [] };
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
                                const resp = window.__mockResponse;
                                if (resp && resp.__throwError) {
                                    if (state.failureCb) state.failureCb(new Error(resp.__throwError));
                                } else {
                                    successCb(resp);
                                }
                            }, 250);
                        };
                    }});
                    return proxy;
                }
            }
        }
    };
"""

SAMPLE_BILLS = """
    ({
      success: true,
      data: [
        { billNumber: 'BILL-901', billDate: '01/01/2026', vendor: 'Vendor A', totalQty: 10, totalAmount: 1180,
          remarks: '', contact: '9876543210', poNumbers: ['1001'],
          items: [ { name: 'Item A', narration: '', size: '', unit: 'Pcs', qty: 10, price: 100, gstRatePct: 18, lineTotal: 1180 } ] },
        { billNumber: 'BILL-902', billDate: '02/01/2026', vendor: 'Vendor B', totalQty: 5, totalAmount: 500,
          remarks: 'Urgent', contact: '', poNumbers: ['DIRECT'],
          items: [ { name: 'Item B', narration: '', size: '', unit: 'Pcs', qty: 5, price: 100, gstRatePct: 0, lineTotal: 500 } ] }
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

        print("\n[Navigate] More tab -> open Bill Ledger sheet")
        page.evaluate("MApp.Shell.showTab('more')")
        page.wait_for_timeout(50)

        print("\n[Loading state] skeleton shows before the mock resolves")
        page.evaluate("window.__mockResponse = " + SAMPLE_BILLS)
        page.evaluate("MApp.Bill.openLedgerSheet(); void 0;")
        skeleton_count = page.locator("#bill-ledger-list .mb-skel").count()
        check(skeleton_count > 0, f"skeleton placeholders render before data arrives (got {skeleton_count})")
        sheet_open = page.evaluate("document.getElementById('sheet-bill-ledger').classList.contains('open')")
        check(sheet_open, "sheet-bill-ledger is opened (has 'open' class)")

        page.wait_for_timeout(350)

        print("\n[Success state] 2 bills render as cards")
        card_count = page.locator("#bill-ledger-list .mb-card").count()
        check(card_count == 2, f"2 bill cards rendered (got {card_count})")

        card_texts = page.evaluate("""
            Object.fromEntries([...document.querySelectorAll('#bill-ledger-list .mb-card')].map(card => {
                const num = card.querySelector('.mb-card-title').textContent.trim();
                return [num, card.textContent];
            }))
        """)
        check('PO-1001' in card_texts.get('BILL-901', ''), f"BILL-901 shows its PO reference (got {card_texts.get('BILL-901')!r})")
        check('Direct' in card_texts.get('BILL-902', ''), f"BILL-902 (DIRECT) shows 'Direct', not the raw sentinel (got {card_texts.get('BILL-902')!r})")

        print("\n[Search] narrows to matching bill/vendor")
        page.fill("#bill-ledger-search", "Vendor A")
        page.evaluate("MApp.Bill.onSearch(document.getElementById('bill-ledger-search').value)")
        page.wait_for_timeout(50)
        check(page.locator("#bill-ledger-list .mb-card").count() == 1, "searching 'Vendor A' narrows to 1 card")
        page.fill("#bill-ledger-search", "")
        page.evaluate("MApp.Bill.onSearch('')")
        page.wait_for_timeout(50)
        check(page.locator("#bill-ledger-list .mb-card").count() == 2, "clearing search restores both cards")

        print("\n[Print] tapping a bill's print icon populates the shared print container")
        page.evaluate("MApp.Bill.print(0)")
        page.wait_for_timeout(50)
        printed = page.evaluate("""
            ({
                billnum: document.getElementById('print-bill-number').innerText,
                vendor: document.getElementById('print-bill-vendor').innerText,
                poRef: document.getElementById('print-bill-po-ref').innerText,
                rows: document.querySelectorAll('#print-bill-items-body tr').length,
                total: document.getElementById('print-bill-grand-total').innerText
            })
        """)
        check(printed['billnum'] == 'BILL-901', f"print container populated with BILL-901's number (got {printed['billnum']!r})")
        check(printed['vendor'] == 'Vendor A', f"print container populated with Vendor A (got {printed['vendor']!r})")
        check('PO-1001' in printed['poRef'], f"print container shows PO-1001 reference (got {printed['poRef']!r})")
        check(printed['rows'] == 1, f"print container has 1 item row (got {printed['rows']})")
        check(printed['total'] == '1180.00', f"print container grand total is 1180.00 (got {printed['total']!r})")

        print("\n[Empty state] zero bills shows the empty message, not a blank list")
        page.evaluate("window.__mockResponse = { success: true, data: [] }")
        page.evaluate("MApp.Bill.openLedgerSheet()")
        page.wait_for_timeout(350)
        check(page.locator("#bill-ledger-list .mb-state").count() > 0, "empty state renders when there are zero bills")
        check(page.locator("#bill-ledger-list .mb-card").count() == 0, "no stale cards remain from the previous load")

        print("\n[Failure state] a rejected getBillData shows the retry error state")
        page.evaluate("window.__mockResponse = { __throwError: 'Simulated network failure' }")
        page.evaluate("MApp.Bill.openLedgerSheet()")
        page.wait_for_timeout(350)
        check(page.locator("#bill-ledger-list .mb-state-error").count() > 0, "error state renders on a rejected getBillData call")
        check(page.locator("#bill-ledger-list .mb-state-retry").count() > 0, "a Retry button is present in the failure state")

        print("\n[Failure state] Retry recovers once the mock is fixed")
        page.evaluate("window.__mockResponse = " + SAMPLE_BILLS)
        page.click("#bill-ledger-list .mb-state-retry")
        page.wait_for_timeout(350)
        check(page.locator("#bill-ledger-list .mb-card").count() == 2, "clicking Retry re-loads and shows both cards")

        page.evaluate("MApp.Bill.closeLedgerSheet()")
        page.wait_for_timeout(50)
        sheet_closed = page.evaluate("!document.getElementById('sheet-bill-ledger').classList.contains('open')")
        check(sheet_closed, "sheet-bill-ledger closes cleanly")

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
