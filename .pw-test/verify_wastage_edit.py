"""
Verifies the new Edit capability for Wastage Log records (Return Ledger >
Wastage review section):
  1. Table row now shows an Edit button.
  2. Clicking it opens the Log Wastage modal pre-filled with the record's
     data, in edit mode (title/submit label change, hidden editing-id set).
  3. Submitting calls updateWastage(wastageId, formData) instead of
     saveWastage.
  4. A successful submit resets the modal back to create mode (title,
     submit label, editing-id cleared) — mirrors verify_issue_edit.py.

Run: python .pw-test/verify_wastage_edit.py
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
            App.State.globalWastage = [{
                wastageId: 'WST-TEST-1', date: '14/07/2026', dateRaw: '2026-07-14',
                vendor: 'Steel Traders', remarks: 'Old remark', totalQty: 3,
                items: [{ name: 'Rod', size: '12mm', qty: 3, unit: 'Pcs', reason: 'Broken during cutting' }]
            }];
            App.State.filteredWastage = [...App.State.globalWastage];

            // Only tracks Wastage-specific calls — Dashboard's background
            // auto-refresh fires unrelated Api.call()s concurrently and would
            // otherwise race with and overwrite a single shared "last call"
            // (see verify_issue_edit.py for the same pattern/reasoning).
            Api.call = async (fn, ...args) => {
                if (fn === 'updateWastage') {
                    window.__lastWastageApiCall = { fn, args };
                    return { success: true, data: { wastageId: args[0] }, message: `Wastage ${args[0]} updated successfully.` };
                }
                if (fn === 'saveWastage') {
                    window.__lastWastageApiCall = { fn, args };
                    return { success: true, data: { wastageId: 'WST-TEST-1' }, message: 'Wastage WST-TEST-1 logged successfully.' };
                }
                return { success: false, message: 'unmocked: ' + fn };
            };
        """)
        page.evaluate("App.Wastage.renderTable()")

        print("[1] Wastage row shows an Edit button...")
        editBtnHtml = page.evaluate("document.querySelector(\"#wastageTableBody button[onclick^='App.Wastage.editWastage']\")?.outerHTML || ''")
        check("Edit" in editBtnHtml, "Edit button present in row actions")

        print("\n[2] Clicking Edit opens the modal pre-filled, in edit mode...")
        page.evaluate("App.Wastage.editWastage('WST-TEST-1')")

        check(page.evaluate("document.getElementById('wastageEditingId').value") == 'WST-TEST-1',
              "hidden editing-id set to the record being edited")
        check(page.evaluate("document.getElementById('wastageVendor').value") == 'Steel Traders',
              "Vendor pre-filled")
        check(page.evaluate("document.querySelector('#wastageForm [name=remarks]').value") == 'Old remark',
              "Remarks pre-filled")
        check(page.evaluate("document.querySelector('#wastageItemsBody .w-item-name').value") == 'Rod',
              "Item row pre-filled")
        check(page.evaluate("document.querySelector('#wastageItemsBody .w-item-reason').value") == 'Broken during cutting',
              "Item reason pre-filled")
        check('Edit Wastage' in page.evaluate("document.getElementById('wastageModalTitle').innerHTML"),
              "Modal title switched to Edit mode")
        check(page.evaluate("document.getElementById('wastageSubmitBtnLabel').textContent") == 'Update Wastage',
              "Submit button label switched to 'Update Wastage'")

        print("\n[3] Submitting calls updateWastage(wastageId, formData)...")
        page.evaluate("document.getElementById('wastageForm').requestSubmit()")
        page.wait_for_timeout(300)

        lastCall = page.evaluate("window.__lastWastageApiCall")
        check(lastCall["fn"] == "updateWastage", "updateWastage was called, not saveWastage")
        check(lastCall["args"][0] == "WST-TEST-1", "called with the correct wastageId")
        check(lastCall["args"][1]["vendor"] == "Steel Traders", "formData carried through correctly")

        print("\n[4] Modal resets back to create mode after a successful update...")
        check(page.evaluate("document.getElementById('wastageEditingId').value") == '',
              "editing-id cleared after successful update")
        check('Edit Wastage' not in page.evaluate("document.getElementById('wastageModalTitle').innerHTML"),
              "Modal title reset to create mode")
        check(page.evaluate("document.getElementById('wastageSubmitBtnLabel').textContent") == 'Log Wastage',
              "Submit button label reset to 'Log Wastage'")

        print("\n[5] Opening for a fresh create after an edit doesn't leak edit state...")
        page.evaluate("App.Wastage.openWastageModal()")
        page.wait_for_timeout(200)
        check(page.evaluate("document.getElementById('wastageEditingId').value") == '',
              "editing-id stays empty when opening for create")
        check(page.evaluate("document.getElementById('wastageVendor').value") == '',
              "Vendor empty in create mode")

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
