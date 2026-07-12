"""
Verifies the P1.3 (delete confirmations) and P1.4 (empty states / busy
buttons / loading rows) fixes added this session.

Run: python .pw-test/verify_p1_3_p1_4_fixes.py
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

        print("[P1.3] Silent-delete fixes now route through confirmAction")
        page.evaluate("""
            App.State.currentProcessContractorRates = { rates: [{ contractorName: 'Sanjay', processName: 'Frame Painting' }] };
            document.getElementById('processFormName') || (() => {
                const el = document.createElement('input');
                el.id = 'processFormName'; el.value = 'Frame Painting';
                document.body.appendChild(el);
            })();
            document.getElementById('processFormName').value = 'Frame Painting';

            const row = document.createElement('tr');
            row.id = 'testRateRow';
            row.innerHTML = '<td><select class="proc-rate-contractor-select"><option value="Sanjay" selected>Sanjay</option></select></td>';
            document.body.appendChild(row);

            window.__confirmCalled = false;
            const origConfirm = App.Utils.confirmAction;
            App.Utils.confirmAction = function(msg, cb) { window.__confirmCalled = true; window.__confirmMsg = msg; };
            App.Process.deleteContractorRateRow('testRateRow');
            window.__confirmAction_restore = origConfirm;
        """)
        check(page.evaluate("window.__confirmCalled === true"), "App.Process.deleteContractorRateRow now calls confirmAction before deleting")
        check("Sanjay" in page.evaluate("window.__confirmMsg") and "Frame Painting" in page.evaluate("window.__confirmMsg"), "confirm message names the contractor and process")

        print("\n[P1.3] Vague-message fixes now include specifics")
        page.evaluate("""
            App.State.globalWarehousePoolOpening = [{ rowIdx: 5, outputItemName: 'Painted Frame', color: 'Blue', qty: 12, date: '10/07/2026' }];
            window.__confirmMsg2 = null;
            App.Utils.confirmAction = function(msg, cb) { window.__confirmMsg2 = msg; };
            App.Stock.deleteWarehouseOpeningEntry(5);
        """)
        msg2 = page.evaluate("window.__confirmMsg2")
        check('Painted Frame' in msg2 and 'Blue' in msg2 and '12' in msg2, f"warehouse opening delete message names item/color/qty (got: {msg2!r})")

        page.evaluate("""
            window.__confirmMsg3 = null;
            App.Utils.confirmAction = function(msg, cb) { window.__confirmMsg3 = msg; };
            App.Contractor.deletePayment(9, 'Sanjay', 1500, '05/07/2026');
        """)
        msg3 = page.evaluate("window.__confirmMsg3")
        check('Sanjay' in msg3 and '05/07/2026' in msg3, f"payment delete message names contractor and date (got: {msg3!r})")

        print("\n[P1.4] Item/Vendor renderTable show a friendly empty state")
        page.evaluate("""
            App.State.filteredItems = []; App.State.globalItems = []; App.State.itemCurrentPage = 1; App.State.itemRowsPerPage = 25;
            App.Item.renderTable();
        """)
        itemHtml = page.evaluate("document.getElementById('itemTableBody').innerHTML")
        check('Register New Item' in itemHtml, "Item empty state references + Register New Item")

        page.evaluate("""
            App.State.filteredVendors = []; App.State.globalVendors = []; App.State.vendorCurrentPage = 1; App.State.vendorRowsPerPage = 25;
            App.Vendor.renderTable();
        """)
        vendorHtml = page.evaluate("document.getElementById('vendorTableBody').innerHTML")
        check('Register New Vendor' in vendorHtml, "Vendor empty state references + Register New Vendor")

        print("\n[P1.4] Master-data forms have submit-button ids wired to busy-state")
        for form_id, btn_id in [('unitForm', 'unitSubmitBtn'), ('colorMasterForm', 'colorMasterSubmitBtn'),
                                 ('modelMasterForm', 'modelMasterSubmitBtn'), ('processTypeMasterForm', 'processTypeMasterSubmitBtn')]:
            exists = page.evaluate(f"!!document.getElementById('{btn_id}')")
            check(exists, f"#{btn_id} exists ({form_id})")

        print("\n[P1.4] Master-data loadData() shows a loading row synchronously (before its first await)")
        # loadData() is async: everything before its first `await Api.call(...)`
        # runs synchronously the instant it's invoked, so calling it WITHOUT
        # awaiting and immediately reading the tbody proves the loading row
        # is set before the fetch — exactly the "within 100ms of clicking"
        # requirement, without needing to mock Api.call's network round-trip.
        results = page.evaluate("""
            (() => {
                const out = {};
                App.Unit.loadData(); out.unit = document.getElementById('unitTableBody').innerHTML;
                App.Color.loadData(); out.color = document.getElementById('colorMasterTableBody').innerHTML;
                App.Model.loadData(); out.model = document.getElementById('modelMasterTableBody').innerHTML;
                App.ProcessType.loadData(); out.processType = document.getElementById('processTypeMasterTableBody').innerHTML;
                return out;
            })()
        """)
        check('Loading' in results.get('unit', ''), "Unit loadData shows a loading row before fetch resolves")
        check('Loading' in results.get('color', ''), "Color loadData shows a loading row before fetch resolves")
        check('Loading' in results.get('model', ''), "Model loadData shows a loading row before fetch resolves")
        check('Loading' in results.get('processType', ''), "ProcessType loadData shows a loading row before fetch resolves")
        page.wait_for_timeout(300)

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
