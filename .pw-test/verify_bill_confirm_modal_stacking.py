"""
Verification script: reported bug — "sometimes, when I record a received
bill, the confirmation dialog appears behind the main form."

Root cause traced to safeModalShow()'s "is another modal already open?"
check (`document.querySelector('.modal.show')`), which doesn't exclude the
modal being (re)shown itself. After successfully recording a NEW bill,
App.Bill.openReceiveModal() re-opens #receiveBillModal for the next entry
while it's technically still marked '.show' from the first round — so
safeModalShow sees ITS OWN modal as "another modal open", treats itself as
nested, and bumps its own z-index to 1070 — the exact same static z-index
#confirmModal uses. That ties the two modals' z-index, and the tie is
broken by DOM order: #confirmModal sits earlier in Modals.html than
#receiveBillModal, so the LATER element wins the tie and renders on top —
meaning every confirmation dialog for the rest of that "record another
bill" sitting appears BEHIND the bill form instead of on top of it.

This reproduces the full flow: record one bill (confirm dialog must be on
top — true even before the fix), then record a SECOND bill in the same
sitting (confirm dialog must still be on top — this is the case that was
broken before the fix).

Run: python .pw-test/verify_bill_confirm_modal_stacking.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

MOCK_API_RESPONSES = {
    "saveBill": {"success": True, "data": {"billNumber": "BILL-1"}, "message": "Bill recorded successfully."},
    "getBillData": {"success": True, "data": []},
}


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_context(viewport={"width": 1400, "height": 900}).new_page()
        page.goto(DIST_HTML.resolve().as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        page.evaluate(f"""
            window.__mockResponses = {json.dumps(MOCK_API_RESPONSES)};
            window.google = {{
                script: {{
                    run: {{
                        withSuccessHandler(cb) {{
                            const runner = {{ withFailureHandler() {{ return runner; }} }};
                            Object.keys(window.__mockResponses).forEach(method => {{
                                runner[method] = (...args) => setTimeout(() => cb(window.__mockResponses[method]), 20);
                            }});
                            return runner;
                        }}
                    }}
                }}
            }};
        """)

        failures = []
        def check(cond, msg):
            print(("PASS: " if cond else "FAIL: ") + msg)
            if not cond:
                failures.append(msg)

        def fill_bill_form(bill_number):
            # Vendor is a Select2-enhanced <select> — bypass the widget and
            # just give the underlying native <select> a valid option/value,
            # which is all form submission/required-field validation cares about.
            page.evaluate("""
                const sel = document.getElementById('billVendor');
                if (![...sel.options].some(o => o.value === 'Acme Vendor')) {
                    sel.add(new Option('Acme Vendor', 'Acme Vendor'));
                }
                sel.value = 'Acme Vendor';
            """)
            page.fill("#receiveBillModal input[name='billNumber']", bill_number)
            page.fill("#billItemsBody .b-item-name", "Steel Rod")
            page.fill("#billItemsBody .b-item-qty", "10")
            page.fill("#billItemsBody .b-item-price", "50")

        def stacking_info():
            return page.evaluate("""
                (() => {
                    const form = document.getElementById('receiveBillModal');
                    const confirm = document.getElementById('confirmModal');
                    const formZ = getComputedStyle(form).zIndex;
                    const confirmZ = getComputedStyle(confirm).zIndex;
                    const dialog = confirm.querySelector('.modal-dialog');
                    const rect = dialog.getBoundingClientRect();
                    const cx = rect.left + rect.width / 2;
                    const cy = rect.top + 30;
                    const topEl = document.elementFromPoint(cx, cy);
                    const topElIsInsideConfirm = confirm.contains(topEl);
                    return { formZ, confirmZ, topElIsInsideConfirm };
                })()
            """)

        print("\n[Round 1] Record the first bill of this sitting...")
        page.evaluate("App.Bill.openReceiveModal()")
        page.locator("#receiveBillModal").wait_for(state="visible", timeout=TIMEOUT)
        fill_bill_form("INV-1001")
        page.locator("#billSubmitBtn").click()
        page.locator("#confirmModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(200)

        info1 = stacking_info()
        print(f"  Bill form z-index: {info1['formZ']}, confirm modal z-index: {info1['confirmZ']}")
        check(info1["topElIsInsideConfirm"], "Round 1: confirm dialog renders on top of the bill form")

        page.locator("#confirmActionBtn").click()
        # saveBill (mocked) resolves, success + !isEdit -> openReceiveModal() reopens the form
        page.wait_for_timeout(400)
        page.locator("#receiveBillModal").wait_for(state="visible", timeout=TIMEOUT)
        print("  Form re-opened for the next entry (as designed after a successful new-bill save)")

        print("\n[Round 2] Record a SECOND bill in the same sitting (this is the reported bug)...")
        fill_bill_form("INV-1002")
        page.locator("#billSubmitBtn").click()
        page.locator("#confirmModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(200)

        info2 = stacking_info()
        print(f"  Bill form z-index: {info2['formZ']}, confirm modal z-index: {info2['confirmZ']}")
        check(info2["topElIsInsideConfirm"], "Round 2: confirm dialog STILL renders on top of the bill form (the reported bug)")
        check(info2["formZ"] != info2["confirmZ"] or info2["formZ"] == "", "Round 2: bill form's z-index was not bumped to tie with the confirm modal's")

        browser.close()

        if failures:
            print(f"\n{len(failures)} CHECK(S) FAILED")
        else:
            print("\nALL CHECKS PASSED")
        return not failures


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
