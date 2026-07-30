"""
End-to-end: create a MULTI-ITEM dispatch bill through the real form, save it via
the real onsubmit handler, reload the ledger, click the row's actual
"Print Challan" button, and assert every challan field rendered and the browser
print dialog opened.

Replaces verify_p1_1_dispatch_print_e2e.py, which drove the pre-refactor
single-item form (#dispatchProductId / #dispatchQty, and a
populateDispatchPrintData that took a flat-array index). Those fields no longer
exist, so that test had been failing against removed UI rather than testing
anything — see the multi-item bill model in module_dispatch.js's header comment.

Deliberately exercises what the old single-item test structurally could not:
TWO line items on one Dispatch Number, so the bill grouping
(buildDispatchBills), the per-line serialization (serializeDispatchLines) and
the multi-row challan item table are all covered on the real path.

window.print() being called is the automated proxy for "the browser print
dialog opens" — a real OS dialog can't be observed from a script, and a called
window.print() is the only observable signal that the browser would open one.

Run: python .pw-test/verify_dispatch_bill_create_print_e2e.py
"""
import sys
import io
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

failures = 0


def check(cond, msg):
    global failures
    if cond:
        print(f"  PASS: {msg}")
    else:
        failures += 1
        print(f"  FAIL: {msg}")


# Mocked backend. saveDispatch echoes back the same {dispatchNumber, rows}
# shape the real one does (see its freshRows read-back), because the client's
# create path calls loadDispatchData() and the edit path patches from rows.
MOCK_BACKEND = """
    window.__savedRows = [];
    window.__saveCalls = [];
    window.__printCalled = false;
    window.__printTitleAtPrint = null;
    window.print = () => {
        window.__printCalled = true;
        // Captured at call time: App.Print.trigger sets document.title to the
        // challan filename for the duration of the print job and resets it a
        // second later, so this is the only moment it can be observed.
        window.__printTitleAtPrint = document.title;
    };
    window.google = {
      script: {
        run: {
          withSuccessHandler(cb) {
            const runner = { withFailureHandler() { return this; } };
            return new Proxy(runner, {
              get(target, prop) {
                if (prop in target) return target[prop];
                return (...args) => {
                  if (prop === 'saveDispatch') {
                    const formData = args[0];
                    window.__saveCalls.push(formData);
                    const lines = JSON.parse(formData.lines || '[]');
                    const dispatchNumber = formData.dispatchNumber
                      || ('DSP-' + (1001 + new Set(window.__savedRows.map(r => r.dispatchNumber)).size));
                    window.__savedRows = window.__savedRows.filter(r => r.dispatchNumber !== dispatchNumber);
                    const rows = lines.map((l, i) => ({
                      rowIdx: 2 + window.__savedRows.length + i,
                      dispatchNumber,
                      dispatchDate: formData.dispatchDate || '',
                      dateRaw: null,
                      orderNumber: formData.orderNumber || '',
                      clientName: formData.clientName || '',
                      productId: l.productId,
                      productName: l.productName,
                      qty: Number(l.qty) || 0,
                      transport: formData.transport || '',
                      remarks: formData.remarks || '',
                      invoiceNumber: formData.invoiceNumber || '',
                      privateMark: formData.privateMark || '',
                      grNumber: formData.grNumber || '',
                      logisticsContractor: formData.logisticsContractor || '',
                      logisticsRate: 0,
                      logisticsCost: 0,
                      rate: Number(l.rate) || 0,
                      amount: (Number(l.qty) || 0) * (Number(l.rate) || 0)
                    }));
                    window.__savedRows = window.__savedRows.concat(rows);
                    setTimeout(() => cb({ success: true, message: 'Dispatch bill recorded successfully.',
                                          data: { dispatchNumber, rows } }), 15);
                    return;
                  }
                  if (prop === 'getDispatchData') {
                    setTimeout(() => cb({ success: true, data: window.__savedRows }), 15);
                    return;
                  }
                  if (prop === 'getReadyToDispatchData') {
                    setTimeout(() => cb({ success: true, data: App.State.globalReadyToDispatch }), 15);
                    return;
                  }
                  setTimeout(() => cb({ success: true, data: [] }), 15);
                };
              }
            });
          }
        }
      }
    };
"""

SEED = """
    App.State.globalClients = [{ name: 'Acme Cycles', contact: '9999999999',
        address: '221B Baker Street, Pune', gstin: '27ABCDE1234F1Z5', remarks: '' }];
    App.State.globalContractors = [];
    App.State.globalOrders = [];
    App.State.globalItems = [];
    // Two dispatchable products, so the bill can carry two real line items.
    App.State.globalReadyToDispatch = [
        { key: '__output__crysta 20 inch', productId: 'PRD-1', productName: 'Crysta 20 inch',
          differentiator: '', producedQty: 25, dispatchedQty: 0, readyQty: 25, colorBreakdown: [] },
        { key: '__output__ranger 24 inch', productId: 'PRD-2', productName: 'Ranger 24 inch',
          differentiator: '', producedQty: 10, dispatchedQty: 0, readyQty: 10, colorBreakdown: [] }
    ];
    App.State.filteredReadyToDispatch = App.State.globalReadyToDispatch;
"""


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_context().new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(600)

        print("=== Setup: mock backend + a Ready-to-Dispatch balance ===")
        page.evaluate(SEED)
        page.evaluate(MOCK_BACKEND)

        print("\n=== Fill the real create-bill form (2 line items) ===")
        page.evaluate("App.Dispatch.openCreateDispatchModal('PRD-1')")
        page.locator("#dispatchModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(250)

        line_count = page.evaluate("document.querySelectorAll('#dispatchLinesBody tr').length")
        check(line_count == 1, f"opening from a Ready row starts the bill with 1 line (got {line_count})")

        prefilled = page.evaluate(
            "document.querySelector('#dispatchLinesBody .dispatch-line-product')?.value")
        check(prefilled == 'PRD-1',
              f"that line is prefilled with the product clicked on Ready to Dispatch (got {prefilled!r})")

        # Second line — the whole point of the multi-item bill.
        page.evaluate("App.Dispatch.addDispatchLineRow()")
        page.wait_for_timeout(150)
        line_count = page.evaluate("document.querySelectorAll('#dispatchLinesBody tr').length")
        check(line_count == 2, f"a second item line can be added (got {line_count})")

        page.evaluate("""
            const rows = Array.from(document.querySelectorAll('#dispatchLinesBody tr'));
            const set = (tr, pid, qty, rate) => {
                const sel = tr.querySelector('.dispatch-line-product');
                sel.value = pid;
                App.Dispatch.handleDispatchLineProductChange(sel);
                tr.querySelector('.dispatch-line-qty').value = qty;
                tr.querySelector('.dispatch-line-rate').value = rate;
            };
            set(rows[0], 'PRD-1', '4', '1200');
            set(rows[1], 'PRD-2', '3', '900');

            document.getElementById('dispatchDate').value = '2026-07-11';
            document.getElementById('dispatchClientSelect').value = 'Acme Cycles';
            document.getElementById('dispatchTransport').value = 'MH-12-AB-1234 / Sharma Transport';
            document.getElementById('dispatchInvoiceNumber').value = 'INV-99';
            document.getElementById('dispatchGrNumber').value = 'GR-77';
            document.getElementById('dispatchRemarks').value = 'Handle with care';
            App.Dispatch.recalcDispatchLinesTotal();
        """)

        serialized = page.evaluate("App.Dispatch.serializeDispatchLines()")
        check(len(serialized) == 2, f"both lines serialize for the payload (got {len(serialized)})")
        check([l['qty'] for l in serialized] == [4, 3],
              f"each line carries its own qty (got {[l['qty'] for l in serialized]})")

        print("\n=== Submit through the real onsubmit handler ===")
        page.evaluate("""
            (async () => {
                const form = document.getElementById('dispatchForm');
                await form.onsubmit(new Event('submit', { cancelable: true }));
            })()
        """)
        page.wait_for_timeout(400)

        calls = page.evaluate("window.__saveCalls")
        check(len(calls) == 1, f"saveDispatch was called exactly once (got {len(calls)})")
        if calls:
            check(calls[0].get('clientName') == 'Acme Cycles',
                  f"header client reached the payload (got {calls[0].get('clientName')!r})")
            check(calls[0].get('grNumber') == 'GR-77',
                  f"header GR number reached the payload (got {calls[0].get('grNumber')!r})")
            check(len(__import__('json').loads(calls[0].get('lines') or '[]')) == 2,
                  "both item lines reached the payload as `lines`")

        rows = page.evaluate("window.__savedRows")
        check(len(rows) == 2, f"the bill persisted as 2 line rows sharing one Dispatch Number (got {len(rows)})")
        check(len({r['dispatchNumber'] for r in rows}) == 1,
              f"both rows share one Dispatch Number (got {sorted({r['dispatchNumber'] for r in rows})})")

        print("\n=== The ledger groups those rows into ONE bill ===")
        bills = page.evaluate("(App.State.globalDispatchBills || []).map(b => ({n: b.dispatchNumber, items: b.items.length, qty: b.totalQty}))")
        check(len(bills) == 1, f"one bill row in the ledger, not two (got {bills})")
        check(bills and bills[0]['items'] == 2 and bills[0]['qty'] == 7,
              f"that bill holds both items and totals 7 units (got {bills})")

        print("\n=== Click the row's real 'Print Challan' button ===")
        clicked = page.evaluate("""
            (() => {
                const btns = Array.from(document.querySelectorAll('#dispatchTableBody button'))
                    .filter(b => b.textContent.trim() === 'Print Challan');
                btns.forEach(b => b.click());
                return btns.length;
            })()
        """)
        check(clicked == 1, f"exactly one Print Challan button was found and clicked (got {clicked})")
        page.wait_for_timeout(150)

        fields = page.evaluate("""
            (() => ({
                number: document.getElementById('print-dispatch-number')?.innerText,
                date: document.getElementById('print-dispatch-date')?.innerText,
                client: document.getElementById('print-dispatch-client')?.innerText,
                clientAddress: document.getElementById('print-dispatch-client-address')?.innerText,
                clientGstin: document.getElementById('print-dispatch-client-gstin')?.innerText,
                transport: document.getElementById('print-dispatch-transport')?.innerText,
                grRef: document.getElementById('print-dispatch-gr-ref')?.innerText,
                remarks: document.getElementById('print-dispatch-remarks')?.innerText,
                itemsHtml: document.getElementById('print-dispatch-items-body')?.innerHTML,
                itemRows: document.querySelectorAll('#print-dispatch-items-body tr').length,
                printCalled: window.__printCalled,
                printTitle: window.__printTitleAtPrint
            }))()
        """)

        expected_number = rows[0]['dispatchNumber'] if rows else 'DSP-1001'
        check(fields["number"] == expected_number,
              f"challan number rendered (got {fields['number']!r}, expected {expected_number!r})")
        check(fields["client"] == "Acme Cycles", f"consignee name rendered (got {fields['client']!r})")
        check(fields["clientAddress"] == "221B Baker Street, Pune",
              f"consignee address looked up from Client Master (got {fields['clientAddress']!r})")
        check(fields["clientGstin"] == "27ABCDE1234F1Z5",
              f"consignee GSTIN looked up from Client Master (got {fields['clientGstin']!r})")
        check("MH-12-AB-1234" in (fields["transport"] or ""),
              f"vehicle/transport rendered (got {fields['transport']!r})")
        check("INV-99" in (fields["grRef"] or "") and "GR-77" in (fields["grRef"] or ""),
              f"invoice + GR reference rendered (got {fields['grRef']!r})")
        check("Handle with care" in (fields["remarks"] or ""),
              f"remarks rendered (got {fields['remarks']!r})")

        # The multi-item assertion the old single-item test could not make.
        check("Crysta 20 inch" in (fields["itemsHtml"] or "")
              and "Ranger 24 inch" in (fields["itemsHtml"] or ""),
              "BOTH item lines render in the challan item table")
        check(fields["itemRows"] == 2,
              f"the challan item table has exactly one <tr> per line item (got {fields['itemRows']})")

        check(fields["printCalled"] is True,
              "window.print() was invoked (the browser print dialog opens)")
        check("Delivery_Challan" in (fields["printTitle"] or "")
              and "Acme_Cycles" in (fields["printTitle"] or ""),
              f"the print job is titled per-bill, so a Save-as-PDF gets a meaningful filename (got {fields['printTitle']!r})")

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
