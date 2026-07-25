"""
Verifies the new Edit capability for Issued Stock records (Production >
Issued Stock sub-tab):
  1. Table row now shows an Edit button.
  2. Clicking it opens the Issue Stock modal pre-filled with the record's
     data, in edit mode (title/submit label change, hidden editing-id set).
  3. Submitting calls updateIssueStock(issueId, formData) instead of
     saveIssueStock, and the post-save Print/Done flow still works.
  4. Cancelling out / re-opening for a fresh create resets the modal back
     to create mode (title, submit label, editing-id cleared).

Run: python .pw-test/verify_issue_edit.py
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
            App.State.globalIssues = [{
                issueId: 'ISS-TEST-1', date: '14/07/2026', dateRaw: '2026-07-14',
                issuedTo: 'Steel Traders', reference: 'JOB-1', vendor: 'Steel Traders',
                remarks: 'Old remark', totalQty: 5, totalValue: 250,
                items: [{ name: 'Rod', size: '12mm', qty: 5, unit: 'Pcs', rate: 50, value: 250 }]
            }];
            App.State.filteredIssues = [...App.State.globalIssues];

            // Only tracks Issue-specific calls — the Dashboard tab's own
            // background auto-refresh (Script_Dashboard.html) fires unrelated
            // Api.call()s concurrently with page load/this test, and would
            // otherwise race with and overwrite a single shared "last call".
            Api.call = async (fn, ...args) => {
                if (fn === 'updateIssueStock') {
                    window.__lastIssueApiCall = { fn, args };
                    return { success: true, data: { issueId: args[0] }, message: `Stock issue ${args[0]} updated successfully.` };
                }
                if (fn === 'saveIssueStock') {
                    window.__lastIssueApiCall = { fn, args };
                    return { success: true, data: { issueId: 'ISS-TEST-1' }, message: 'Stock issue ISS-TEST-1 logged successfully.' };
                }
                return { success: false, message: 'unmocked: ' + fn };
            };
        """)
        page.evaluate("App.Issue.renderTable()")

        print("[1] Issued Stock row shows an Edit button...")
        editBtnHtml = page.evaluate("document.querySelector(\"#issueTableBody button[onclick^='App.Issue.editIssue']\")?.outerHTML || ''")
        check("Edit" in editBtnHtml, "Edit button present in row actions")

        print("\n[2] Clicking Edit opens the modal pre-filled, in edit mode...")
        page.evaluate("App.Issue.editIssue('ISS-TEST-1')")
        page.locator("#issueStockModal").wait_for(state="visible", timeout=8000)

        check(page.evaluate("document.getElementById('issueEditingId').value") == 'ISS-TEST-1',
              "hidden editing-id set to the record being edited")
        check(page.evaluate("document.getElementById('issueIssuedTo').value") == 'Steel Traders',
              "Issued To pre-filled")
        check(page.evaluate("document.getElementById('issueReferenceInput').value") == 'JOB-1',
              "Reference pre-filled")
        check(page.evaluate("document.querySelector('#issueStockForm [name=remarks]').value") == 'Old remark',
              "Remarks pre-filled")
        check(page.evaluate("document.querySelector('#issueItemsBody .i-item-name').value") == 'Rod',
              "Item row pre-filled")
        check(page.evaluate("document.querySelector('#issueItemsBody .i-item-qty').value") == '5',
              "Item qty pre-filled")
        check('Edit Issue' in page.evaluate("document.getElementById('issueStockModalTitle').innerHTML"),
              "Modal title switched to Edit mode")
        check(page.evaluate("document.getElementById('issueStockSubmitBtnLabel').textContent") == 'Update Issue',
              "Submit button label switched to 'Update Issue'")

        print("\n[3] Submitting calls updateIssueStock(issueId, formData)...")
        page.evaluate("document.getElementById('issueStockForm').requestSubmit()")
        page.wait_for_timeout(300)

        lastCall = page.evaluate("window.__lastIssueApiCall")
        check(lastCall["fn"] == "updateIssueStock", "updateIssueStock was called, not saveIssueStock")
        check(lastCall["args"][0] == "ISS-TEST-1", "called with the correct issueId")
        check(lastCall["args"][1]["issuedTo"] == "Steel Traders", "formData carried through correctly")

        check(page.evaluate("document.getElementById('issueStockPrintBtn').style.display") != 'none',
              "Print button shown after update (same post-save flow as create)")
        check(page.evaluate("document.getElementById('issueStockDoneBtn').style.display") != 'none',
              "Done button shown after update")
        check(page.evaluate("document.getElementById('issueSavedId').value") == 'ISS-TEST-1',
              "saved-id set to the edited record's id")

        print("\n[4] Done resets modal back to create mode...")
        page.evaluate("App.Issue.done()")
        page.wait_for_timeout(300)
        check(page.evaluate("document.getElementById('issueEditingId').value") == '',
              "editing-id cleared after Done")
        check('Edit Issue' not in page.evaluate("document.getElementById('issueStockModalTitle').innerHTML"),
              "Modal title reset to create mode")
        check(page.evaluate("document.getElementById('issueStockSubmitBtnLabel').textContent") == 'Issue Stock',
              "Submit button label reset to 'Issue Stock'")

        print("\n[5] Opening for a fresh create after an edit doesn't leak edit state...")
        page.evaluate("App.Issue.openIssueModal()")
        page.wait_for_timeout(200)
        check(page.evaluate("document.getElementById('issueEditingId').value") == '',
              "editing-id stays empty when opening for create")
        check(page.evaluate("document.getElementById('issueIssuedTo').value") == '',
              "Issued To empty in create mode")

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
