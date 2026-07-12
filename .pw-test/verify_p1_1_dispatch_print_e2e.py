"""
P3 Verification — P1.1 proof: "create a dispatch, print the challan,
confirm all fields render and the browser print dialog opens."

Drives the REAL create-dispatch form (not direct state injection) against a
mocked backend, submits it via the real onsubmit handler, then clicks the
resulting row's actual "Print Challan" button and asserts every challan
field is populated and window.print() was invoked (the automated proxy for
"the browser print dialog opens" — a real OS print dialog can't be
observed/asserted from a script; a called window.print() is the standard
and only observable signal that the browser would have opened one).

Run: python .pw-test/verify_p1_1_dispatch_print_e2e.py
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

        print("=== Setup: mock backend + seed a Ready-to-Dispatch balance ===")
        page.evaluate("""
            App.State.globalClients = [{ name: 'Acme Cycles', contact: '9999999999', address: '221B Baker Street, Pune', gstin: '27ABCDE1234F1Z5', remarks: '' }];
            App.State.globalItems = [{ name: 'Crysta 20 inch', size: 'General' }];
            App.State.globalContractors = [];
            App.State.globalReadyToDispatch = [
                { productId: 'PRD-1', productName: 'Crysta 20 inch', readyQty: 25 }
            ];
            window.__savedDispatches = [];
            window.__printCalled = false;
            window.print = () => { window.__printCalled = true; };
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
                                            const dispatchNumber = 'DSP-' + (1000 + window.__savedDispatches.length + 1);
                                            window.__savedDispatches.push({
                                                rowIdx: 2 + window.__savedDispatches.length,
                                                dispatchNumber,
                                                dispatchDate: formData.dispatchDate,
                                                orderNumber: formData.orderNumber || '',
                                                clientName: formData.clientName || '',
                                                productId: formData.productId,
                                                productName: formData.productName,
                                                qty: Number(formData.qty),
                                                transport: formData.transport || '',
                                                remarks: formData.remarks || '',
                                                invoiceNumber: formData.invoiceNumber || '',
                                                grNumber: formData.grNumber || '',
                                                logisticsContractor: formData.logisticsContractor || '',
                                                logisticsRate: 0, logisticsCost: 0
                                            });
                                            setTimeout(() => cb({ success: true, message: 'Dispatch recorded successfully.', data: { dispatchNumber } }), 15);
                                            return;
                                        }
                                        if (prop === 'getDispatchData') {
                                            setTimeout(() => cb({ success: true, data: window.__savedDispatches }), 15);
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
        """)

        print("\n=== Create a dispatch through the real form ===")
        page.evaluate("App.Dispatch.openCreateDispatchModal('PRD-1');")
        page.locator("#dispatchModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(200)

        productId = page.evaluate("document.getElementById('dispatchProductId').value")
        check(productId == 'PRD-1', f"product auto-selected from the Ready-to-Dispatch prefill (got {productId!r})")

        page.evaluate("""
            document.getElementById('dispatchQty').value = '4';
            document.getElementById('dispatchTransport').value = 'MH-12-AB-1234 / Sharma Transport';
            document.getElementById('dispatchInvoiceNumber').value = 'INV-99';
            document.getElementById('dispatchGrNumber').value = 'GR-77';
            document.getElementById('dispatchRemarks').value = 'Handle with care';
        """)

        save_result = page.evaluate("""
            (async () => {
                const formEl = document.getElementById('dispatchForm');
                return await formEl.onsubmit(new Event('submit', { cancelable: true }));
            })()
        """)

        saved_count = page.evaluate("window.__savedDispatches.length")
        check(saved_count == 1, f"exactly 1 dispatch record was saved (got {saved_count})")

        print("\n=== Reload the dispatched-goods table and print the new row's challan ===")
        page.evaluate("App.Dispatch.loadDispatchData();")
        page.wait_for_timeout(200)

        row_count = page.evaluate("(App.State.globalDispatch || []).length")
        check(row_count == 1, f"the newly created dispatch appears in the table's backing data (got {row_count} row(s))")

        # Click the real per-row "Print Challan" button, matching what a
        # user would actually click — not calling App.Dispatch.print()
        # directly.
        page.evaluate("""
            document.querySelectorAll('#dispatchTableBody button').forEach(b => {
                if (b.textContent.trim() === 'Print Challan') b.click();
            });
        """)
        page.wait_for_timeout(100)

        fields = page.evaluate("""
            (() => ({
                number: document.getElementById('print-dispatch-number')?.innerText,
                date: document.getElementById('print-dispatch-date')?.innerText,
                client: document.getElementById('print-dispatch-client')?.innerText,
                clientAddress: document.getElementById('print-dispatch-client-address')?.innerText,
                clientGstin: document.getElementById('print-dispatch-client-gstin')?.innerText,
                transport: document.getElementById('print-dispatch-transport')?.innerText,
                itemsHtml: document.getElementById('print-dispatch-items-body')?.innerHTML,
                printCalled: window.__printCalled
            }))()
        """)

        check(fields["number"] == "DSP-1001", f"challan number rendered (got {fields['number']!r})")
        check("Acme Cycles" == fields["client"], f"consignee name rendered (got {fields['client']!r})")
        check(fields["clientAddress"] == "221B Baker Street, Pune", f"consignee address rendered (got {fields['clientAddress']!r})")
        check(fields["clientGstin"] == "27ABCDE1234F1Z5", f"consignee GSTIN rendered (got {fields['clientGstin']!r})")
        check("MH-12-AB-1234" in (fields["transport"] or ""), f"vehicle/transport rendered (got {fields['transport']!r})")
        check("Crysta 20 inch" in (fields["itemsHtml"] or ""), "item row rendered in the challan item table")
        check(fields["printCalled"] is True, "window.print() was invoked (the browser print dialog opens)")

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
