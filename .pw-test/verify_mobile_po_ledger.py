"""
Verification script — Mobile PO Ledger (More tab > PO Ledger sheet).

Drives the compiled mobile shell (dist/mobile.html, produced by
compile_preview.py from Mobile_Index.html) directly via Playwright. Mocks
google.script.run so getPOData() responses are fully controlled from the
test -- covering the loading/empty/success/failure states
MApp.PO.openLedgerSheet() must degrade through (per the "every new mobile
screen needs empty, loading, and failure states" requirement), plus search,
status filtering, and print-field population (reusing the SAME
#print-po-container markup from View_Print.html the desktop PO Ledger
populates).

Run: python .pw-test/verify_mobile_po_ledger.py
"""
import sys
import io
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "mobile.html"
TIMEOUT = 8000

failures = 0
def check(cond, msg):
    global failures
    if cond:
        print(f"  PASS: {msg}")
    else:
        failures += 1
        print(f"  FAIL: {msg}")


# Captures BOTH the success and failure callbacks per call (unlike
# verify_po_status_bar.py's mock, which only ever needed the success path),
# so window.__mockResponse can drive either outcome per test phase.
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
                            return proxy; // must stay Proxy-wrapped, or the next
                                          // .methodName(...) call below has no
                                          // catch-all trap and resolves undefined
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
                            }, 250); // long enough to reliably observe the
                                      // loading/skeleton state before it
                                      // resolves, despite Python<->browser
                                      // round-trip jitter between evaluate() calls
                        };
                    }});
                    return proxy;
                }
            }
        }
    };
"""

SAMPLE_POS = """
    ({
      success: true,
      data: [
        { poNumber: 'PO-501', poDate: '01/01/2026', vendor: 'Vendor A', totalQty: 10, grandTotal: 1000,
          status: 'PO Issued', poDescription: '', poRemarks: '', supplierRemarks: '', contact: '9876543210',
          items: [ { name: 'Item A', narration: '', size: '', unit: 'Pcs', qty: 10, price: 100, receivedQty: 0, pendingQty: 10 } ] },
        { poNumber: 'PO-502', poDate: '02/01/2026', vendor: 'Vendor B', totalQty: 20, grandTotal: 2000,
          status: 'Partially Received', poDescription: '', poRemarks: '', supplierRemarks: '', contact: '',
          items: [ { name: 'Item B', narration: '', size: '', unit: 'Pcs', qty: 20, price: 100, receivedQty: 8, pendingQty: 12 } ] },
        { poNumber: 'PO-503', poDate: '03/01/2026', vendor: 'Vendor C', totalQty: 5, grandTotal: 500,
          status: 'Completed', poDescription: '', poRemarks: '', supplierRemarks: '', contact: '',
          items: [ { name: 'Item C', narration: '', size: '', unit: 'Pcs', qty: 5, price: 100, receivedQty: 5, pendingQty: 0 } ] }
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

        print("\n[Navigate] More tab -> open PO Ledger sheet")
        page.evaluate("MApp.Shell.showTab('more')")
        page.wait_for_timeout(50)

        print("\n[Loading state] skeleton shows immediately, before the mock resolves")
        page.evaluate("window.__mockResponse = " + SAMPLE_POS)
        # openLedgerSheet() is async; page.evaluate() awaits any returned
        # Promise by default, which would skip straight past the transient
        # loading state. Appending "void 0" makes the expression's own value
        # a plain undefined -- Playwright doesn't wait for it -- while the
        # call itself still fires and keeps running in the background.
        page.evaluate("MApp.PO.openLedgerSheet(); void 0;")
        skeleton_count = page.locator("#po-ledger-list .mb-skel").count()
        check(skeleton_count > 0, f"skeleton placeholders render before data arrives (got {skeleton_count})")

        sheet_open = page.evaluate("document.getElementById('sheet-po-ledger').classList.contains('open')")
        check(sheet_open, "sheet-po-ledger is opened (has 'open' class)")

        page.wait_for_timeout(350)

        print("\n[Success state] 3 POs render as cards with correct status chips")
        card_count = page.locator("#po-ledger-list .mb-card").count()
        check(card_count == 3, f"3 PO cards rendered (got {card_count})")

        chip_classes = page.evaluate("""
            Object.fromEntries([...document.querySelectorAll('#po-ledger-list .mb-card')].map(card => {
                const poNum = card.querySelector('.mb-card-title').textContent.trim();
                const chip = card.querySelector('.mb-chip');
                return [poNum, chip.className];
            }))
        """)
        check('mb-chip-inprogress' in chip_classes.get('PO-501', ''), f"PO-501 (PO Issued) uses mb-chip-inprogress (got {chip_classes.get('PO-501')})")
        check('mb-chip-pending' in chip_classes.get('PO-502', ''), f"PO-502 (Partially Received) uses mb-chip-pending (got {chip_classes.get('PO-502')})")
        check('mb-chip-completed' in chip_classes.get('PO-503', ''), f"PO-503 (Completed) uses mb-chip-completed (got {chip_classes.get('PO-503')})")

        print("\n[Pending qty detail] PO-502's partial line shows its pending qty")
        po502_text = page.evaluate("""
            [...document.querySelectorAll('#po-ledger-list .mb-card')]
                .find(c => c.querySelector('.mb-card-title').textContent.includes('PO-502')).textContent
        """)
        check('12' in po502_text and 'pending' in po502_text, f"PO-502 card mentions 12 ... pending (got {po502_text!r})")
        po501_text = page.evaluate("""
            [...document.querySelectorAll('#po-ledger-list .mb-card')]
                .find(c => c.querySelector('.mb-card-title').textContent.includes('PO-503')).textContent
        """)
        check('pending' not in po501_text, "PO-503 (fully completed) card shows no pending line")

        print("\n[Search] narrows to matching PO/vendor")
        page.fill("#po-ledger-search", "Vendor A")
        page.evaluate("MApp.PO.onSearch(document.getElementById('po-ledger-search').value)")
        page.wait_for_timeout(50)
        check(page.locator("#po-ledger-list .mb-card").count() == 1, "searching 'Vendor A' narrows to 1 card")
        page.fill("#po-ledger-search", "")
        page.evaluate("MApp.PO.onSearch('')")
        page.wait_for_timeout(50)

        print("\n[Status filter] tapping 'Completed' chip shows only that PO")
        page.evaluate("MApp.PO.filterByStatus('Completed')")
        page.wait_for_timeout(50)
        check(page.locator("#po-ledger-list .mb-card").count() == 1, "filtering to 'Completed' shows 1 card")
        active_chip = page.evaluate("document.querySelector('#po-ledger-status-bar .mb-filter-chip.active')?.dataset.status")
        check(active_chip == 'Completed', f"'Completed' chip carries the active class (got {active_chip!r})")

        print("\n[Status filter] 'All' resets")
        page.evaluate("MApp.PO.filterByStatus('all')")
        page.wait_for_timeout(50)
        check(page.locator("#po-ledger-list .mb-card").count() == 3, "'all' filter shows all 3 cards again")

        print("\n[Print] tapping a PO's print icon populates the shared print container")
        page.evaluate("MApp.PO.print(0)")
        page.wait_for_timeout(50)
        printed = page.evaluate("""
            ({
                ponum: document.getElementById('print-ponum').innerText,
                vendor: document.getElementById('print-vendor').innerText,
                rows: document.querySelectorAll('#print-items-body tr').length
            })
        """)
        check(printed['ponum'] == 'PO-501', f"print container populated with PO-501's number (got {printed['ponum']!r})")
        check(printed['vendor'] == 'Vendor A', f"print container populated with Vendor A (got {printed['vendor']!r})")
        check(printed['rows'] == 1, f"print container has 1 item row (got {printed['rows']})")

        print("\n[Empty state] zero POs shows the empty message, not a blank list")
        page.evaluate("window.__mockResponse = { success: true, data: [] }")
        page.evaluate("MApp.PO.openLedgerSheet()")
        page.wait_for_timeout(350)
        empty_state = page.locator("#po-ledger-list .mb-state").count()
        check(empty_state > 0, "empty state renders when there are zero POs")
        check(page.locator("#po-ledger-list .mb-card").count() == 0, "no stale cards remain from the previous load")

        print("\n[Failure state] a rejected getPOData shows the retry error state, not a blank/frozen sheet")
        page.evaluate("window.__mockResponse = { __throwError: 'Simulated network failure' }")
        page.evaluate("MApp.PO.openLedgerSheet()")
        page.wait_for_timeout(350)
        error_state = page.locator("#po-ledger-list .mb-state-error").count()
        check(error_state > 0, "error state renders on a rejected getPOData call")
        retry_btn = page.locator("#po-ledger-list .mb-state-retry").count()
        check(retry_btn > 0, "a Retry button is present in the failure state")

        print("\n[Failure state] Retry recovers once the mock is fixed")
        page.evaluate("window.__mockResponse = " + SAMPLE_POS)
        page.click("#po-ledger-list .mb-state-retry")
        page.wait_for_timeout(350)
        check(page.locator("#po-ledger-list .mb-card").count() == 3, "clicking Retry re-loads and shows all 3 cards")

        page.evaluate("MApp.PO.closeLedgerSheet()")
        page.wait_for_timeout(50)
        sheet_closed = page.evaluate("!document.getElementById('sheet-po-ledger').classList.contains('open')")
        check(sheet_closed, "sheet-po-ledger closes cleanly")

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
