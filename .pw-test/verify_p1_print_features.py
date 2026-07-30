"""
Verifies the P1.1/P1.2 print features added this session:
  - App.Dispatch.print (Delivery Challan) — new print-dispatch-container
  - App.Bill.print (per-row, existing print-bill-container)
  - App.Return.print (per-row, via triggerBulk + buildReturnPrintPageHtml)
  - App.Issue.print (per-row, via triggerBulk + existing buildIssuePrintPageHtml)
  - App.Client.printOrder (PI/Estimate print, via triggerBulk + buildOrderPrintPageHtml)

Run: python .pw-test/verify_p1_print_features.py
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

        page.evaluate("""
            App.State.globalDispatch = [{
                rowIdx: 2, dispatchNumber: 'DSP-1001', dispatchDate: '11/07/2026',
                orderNumber: 'PI-5001', clientName: 'Acme Cycles', productId: 'PRD-1',
                productName: 'Crysta 20 inch', qty: 25, transport: 'MH-12-AB-1234 / Sharma Transport',
                remarks: 'Handle with care', invoiceNumber: 'INV-99', grNumber: 'GR-77',
                logisticsContractor: '', logisticsRate: 0, logisticsCost: 0
            }];
            App.State.globalClients = [{ name: 'Acme Cycles', contact: '9999999999', address: '221B Baker Street, Pune', gstin: '27ABCDE1234F1Z5', remarks: '' }];
            App.State.globalItems = [{ name: 'Crysta 20 inch', size: 'General' }];

            App.State.globalBills = [{
                billNumber: 'BILL-1', billDate: '10/07/2026', vendor: 'Steel Traders',
                contact: '8888888888', remarks: 'Test bill', poNumbers: ['PO-1'],
                totalQty: 10, totalAmount: 5000,
                items: [{ name: 'Rod', narration: 'MS Rod', size: '12mm', qty: 10, unit: 'Pcs', price: 500, gstRatePct: 18, lineTotal: 5000 }]
            }];

            App.State.globalReturns = [{
                returnNumber: 'RET-1', returnDate: '09/07/2026', vendor: 'Steel Traders',
                contact: '8888888888', billNumber: 'BILL-1', remarks: 'Defective batch',
                totalQty: 2, totalAmount: 1000,
                items: [{ name: 'Rod', narration: 'MS Rod', size: '12mm', qty: 2, unit: 'Pcs', price: 500, reason: 'Damaged', lineTotal: 1000 }]
            }];

            App.State.globalIssues = [{
                issueId: 'ISS-1', date: '08/07/2026', issuedTo: 'Maintenance', reference: 'REF-1',
                totalQty: 3, remarks: 'For repair',
                items: [{ name: 'Bolt', size: 'M8', qty: 3, unit: 'Pcs' }]
            }];

            App.State.globalOrders = [{
                orderNumber: 'PI-5001', orderDate: '07/07/2026', clientName: 'Acme Cycles',
                status: 'Estimate', orderRemarks: 'Urgent',
                lines: [{ productId: 'PRD-1', productName: 'Crysta 20 inch', qty: 25, lineRemarks: 'Blue only' }]
            }];
        """)

        print("[P1.1] App.Dispatch.print / populateDispatchPrintData")
        # populateDispatchPrintData now takes a Dispatch NUMBER and reads the
        # grouped globalDispatchBills (a bill is N line rows), not a flat-array
        # INDEX. Called with 0 against an unbuilt bills array it found no bill,
        # returned null early and left every print field blank — so all six
        # checks below failed for that reason rather than anything about the
        # challan layout they exist to verify.
        page.evaluate("App.Dispatch.buildDispatchBills()")
        page.evaluate("App.Dispatch.populateDispatchPrintData('DSP-1001')")
        check(page.evaluate("document.getElementById('print-dispatch-number').innerText") == 'DSP-1001', "challan number populated")
        check(page.evaluate("document.getElementById('print-dispatch-client').innerText") == 'Acme Cycles', "consignee name populated")
        check(page.evaluate("document.getElementById('print-dispatch-client-gstin').innerText") == '27ABCDE1234F1Z5', "consignee GSTIN looked up from Client Master")
        check(page.evaluate("document.getElementById('print-dispatch-client-address').innerText") == '221B Baker Street, Pune', "consignee address looked up from Client Master")
        check('MH-12-AB-1234' in page.evaluate("document.getElementById('print-dispatch-transport').innerText"), "vehicle/transport populated")
        check('Crysta 20 inch' in page.evaluate("document.getElementById('print-dispatch-items-body').innerHTML"), "item row rendered")

        print("\n[P1.2] App.Bill.print / populatePrintData (per-row)")
        page.evaluate("App.Bill.populatePrintData(0)")
        check(page.evaluate("document.getElementById('print-bill-number').innerText") == 'BILL-1', "bill number populated")
        check(page.evaluate("document.getElementById('print-bill-vendor').innerText") == 'Steel Traders', "vendor populated")
        check('Rod' in page.evaluate("document.getElementById('print-bill-items-body').innerHTML"), "item row rendered")
        check(page.evaluate("document.getElementById('print-bill-grand-total').innerText") == '5000.00', "grand total populated")

        print("\n[P1.2] App.Return.print (per-row, via triggerBulk)")
        page.evaluate("App.Return.print(0)")
        returnHtml = page.evaluate("document.getElementById('print-bulk-body').innerHTML")
        check('RET-1' in returnHtml, "return number rendered in bulk-print body")
        check('Damaged' in returnHtml, "reason column rendered")

        print("\n[P1.2] App.Issue.print (per-row, via triggerBulk)")
        page.evaluate("App.Issue.print('ISS-1')")
        issueHtml = page.evaluate("document.getElementById('print-bulk-body').innerHTML")
        check('ISS-1' in issueHtml, "issue id rendered in bulk-print body")
        check('Bolt' in issueHtml, "item rendered")

        print("\n[P1.2] App.Client.printOrder (per-row, via triggerBulk)")
        page.evaluate("App.Client.printOrder(0)")
        orderHtml = page.evaluate("document.getElementById('print-bulk-body').innerHTML")
        check('PI-5001' in orderHtml, "order number rendered in bulk-print body")
        check('Crysta 20 inch' in orderHtml, "product line rendered")

        print("\n[P1.2] Modal edit-mode Print buttons exist and are wired")
        check(page.evaluate("!!document.getElementById('poModalPrintBtn')"), "#poModalPrintBtn exists")
        check(page.evaluate("!!document.getElementById('billModalPrintBtn')"), "#billModalPrintBtn exists")
        check(page.evaluate("!!document.getElementById('returnModalPrintBtn')"), "#returnModalPrintBtn exists")
        check(page.evaluate("!!document.getElementById('clientOrderPrintBtn')"), "#clientOrderPrintBtn exists")
        check(page.evaluate("typeof App.PO.printCurrent === 'function'"), "App.PO.printCurrent defined")
        check(page.evaluate("typeof App.Bill.printCurrent === 'function'"), "App.Bill.printCurrent defined")
        check(page.evaluate("typeof App.Return.printCurrent === 'function'"), "App.Return.printCurrent defined")
        check(page.evaluate("typeof App.Client.printCurrentOrder === 'function'"), "App.Client.printCurrentOrder defined")

        if console_errors:
            print("\nConsole/page errors:")
            for e in console_errors:
                print(f"    {e}")
            failures_local = len(console_errors)
            globals()['failures'] += failures_local

        browser.close()


if __name__ == "__main__":
    run()
    print(f"\n{'ALL PASS' if failures == 0 else str(failures) + ' FAILURE(S)'}")
    sys.exit(0 if failures == 0 else 1)
