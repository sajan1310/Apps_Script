"""
Verifies the two features added this session to Issue Stock:
  1. In-modal Print/Done flow after Save (form locks read-only, Print
     button reuses App.Issue.print, Done closes and resets).
  2. "Issued To / Purpose" doubling as the Vendor field: typing a name
     that matches the Vendor Master (case-insensitively) links the issue
     to that vendor's Ledger — no separate Vendor input. Also covers the
     lazy-load fix so App.State.globalIssues loads even if the user never
     visited Production's Issued Stock sub-tab first.

Run: python .pw-test/verify_issue_vendor_and_print.py
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

        # Seed vendor + empty PO/Bill/Return caches. globalIssues starts
        # EMPTY on purpose (fresh session) to validate the lazy-load fix.
        page.evaluate("""
            App.State.globalVendors = [{ name: 'Steel Traders', contact: '8888888888', address: 'Ludhiana', gstin: '', remarks: '' }];
            App.State.globalPOs = [];
            App.State.globalBills = [];
            App.State.globalReturns = [];
            App.State.globalIssues = [];

            Api.call = async (fn, payload) => {
                if (fn === 'saveIssueStock') {
                    return { success: true, data: { issueId: 'ISS-TEST-1' }, message: 'Stock issue ISS-TEST-1 logged successfully.' };
                }
                if (fn === 'getIssueData') {
                    return {
                        success: true,
                        data: [{
                            issueId: 'ISS-TEST-1', date: '14/07/2026', dateRaw: '2026-07-14',
                            issuedTo: 'Steel Traders', reference: 'JOB-1', vendor: 'Steel Traders',
                            remarks: '', totalQty: 5, totalValue: 250,
                            items: [{ name: 'Rod', size: '12mm', qty: 5, unit: 'Pcs', rate: 50, value: 250 }]
                        }]
                    };
                }
                return { success: false, message: 'unmocked: ' + fn };
            };
        """)

        print("[1] Open Issue Stock modal, fill Issued To with a known vendor name + Rate...")
        page.evaluate("App.Issue.openIssueModal()")
        page.locator("#issueStockModal").wait_for(state="visible", timeout=8000)

        # "Issued To / Purpose" is the single merged field — typing a name
        # that matches the Vendor Master (case-insensitively) is how the
        # backend links the issue to that vendor's ledger; there's no
        # separate Vendor input anymore.
        page.fill("#issueIssuedTo", "Steel Traders")
        page.fill("#issueItemsBody tr:first-child .i-item-name", "Rod")
        page.fill("#issueItemsBody tr:first-child .i-item-qty", "5")
        page.fill("#issueItemsBody tr:first-child .i-item-rate", "50")

        check(page.evaluate("document.getElementById('issueStockPrintBtn').style.display") == 'none',
              "Print button hidden before save")
        check(page.evaluate("document.getElementById('issueStockDoneBtn').style.display") == 'none',
              "Done button hidden before save")

        print("\n[2] Submit the form (mocked saveIssueStock)...")
        page.evaluate("document.getElementById('issueStockForm').requestSubmit()")
        page.wait_for_timeout(300)

        check(page.evaluate("document.getElementById('issueStockSubmitBtn').style.display") == 'none',
              "Submit button hidden after save")
        check(page.evaluate("document.getElementById('issueStockPrintBtn').style.display") != 'none',
              "Print button shown after save")
        check(page.evaluate("document.getElementById('issueStockDoneBtn').style.display") != 'none',
              "Done button shown after save")
        check(page.evaluate("document.getElementById('issueIssuedTo').disabled") is True,
              "Issued To field disabled (read-only) after save")
        check(page.evaluate("document.getElementById('issueSavedId').value") == 'ISS-TEST-1',
              "hidden saved-id field populated with new issueId")
        check(page.evaluate("!document.getElementById('issueStockDoneBtn').disabled"),
              "Done button itself stays enabled (not swept up by form-disable)")
        check(page.evaluate("!document.querySelector('#issueStockModal button[data-bs-dismiss=\"modal\"]').disabled"),
              "Cancel button itself stays enabled (not swept up by form-disable)")

        print("\n[3] Click Print — reuses App.Issue.print(issueId)...")
        page.evaluate("App.Issue.printCurrent()")
        printedHtml = page.evaluate("document.getElementById('print-bulk-body').innerHTML")
        check('ISS-TEST-1' in printedHtml, "issue id rendered in print output")
        check('Rod' in printedHtml, "item rendered in print output")
        check('250.00' in printedHtml, "computed rate*qty value rendered in print output")

        print("\n[4] Click Done — closes modal and resets to create mode...")
        page.evaluate("App.Issue.done()")
        page.wait_for_timeout(300)
        check(page.evaluate("document.getElementById('issueSavedId').value") == '',
              "saved-id cleared after Done")
        check(page.evaluate("document.getElementById('issueStockSubmitBtn').style.display") != 'none',
              "Submit button restored after Done")
        check(page.evaluate("document.getElementById('issueIssuedTo').disabled") is False,
              "form fields re-enabled after Done")

        print("\n[5] Vendor Ledger reflects the vendor-linked issue (fresh session, lazy-load)...")
        # Reset to simulate a session where Production's Issued Stock tab was
        # never visited, so App.State.globalIssues is empty going into the
        # vendor profile — this is the actual lazy-load guard under test.
        page.evaluate("App.State.globalIssues = [];")
        check(page.evaluate("App.State.globalIssues.length") == 0,
              "globalIssues reset to empty before opening vendor profile (sanity check for lazy-load test)")
        page.evaluate("App.Vendor.openProfileModal('Steel Traders')")
        page.wait_for_timeout(300)
        ledgerHtml = page.evaluate("document.getElementById('vendorLedgerBody').innerHTML")
        check(page.evaluate("App.State.globalIssues.length") > 0,
              "globalIssues lazy-loaded by openProfileModal")
        check('Stock Issued' in ledgerHtml, "'Stock Issued' entry type rendered in vendor ledger")
        check('ISS-TEST-1' in ledgerHtml, "issue id rendered as ledger reference")
        check('250.00' in ledgerHtml, "computed value rendered in vendor ledger")

        pending = page.evaluate("App.Vendor.calculateLedgerAndPending('Steel Traders').pendingList")
        check(pending == [], "Issue entries do not pollute the PO-vs-Bill pending reconciliation")

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
