"""
Verification script — Mobile "New PO" write flow (More tab > PO Ledger >
FAB > New PO).

The first write action added to the mobile PO Ledger screen (previously
read-only: list + status filter + print). Mirrors desktop's savePO
(module_po.js, unchanged) but -- like MApp.Returns.save() -- logs exactly
ONE item per PO instead of desktop's multi-line form, a deliberate scope
narrowing for fast field entry.

Covers: opening the sheet loads vendor/item reference data, picking a
vendor auto-fills its known contact, validation (vendor required, item
required, qty > 0) blocks the save, a successful save closes the sheet +
toasts the new PO number + refreshes the ledger list behind it, and a
rejected save keeps the sheet open with an error toast.

Run: python .pw-test/verify_mobile_new_po.py
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
    window.__mockRoutes = {};
    window.__callCounts = {};
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
                            window.__callCounts[prop] = (window.__callCounts[prop] || 0) + 1;
                            window.__lastArgs = window.__lastArgs || {};
                            window.__lastArgs[prop] = args;
                            setTimeout(() => {
                                const resp = (prop in window.__mockRoutes) ? window.__mockRoutes[prop] : { success: true, data: [] };
                                if (resp && resp.__throwError) {
                                    if (state.failureCb) state.failureCb(new Error(resp.__throwError));
                                } else {
                                    successCb(resp);
                                }
                            }, 60);
                        };
                    }});
                    return proxy;
                }
            }
        }
    };
"""

SAMPLE_POS = """
    ({ success: true, data: [
        { poNumber: 'PO-100', vendor: 'Acme Vendor', poDate: '01/01/2026', status: 'Completed', totalQty: 10, grandTotal: 1000, items: [] }
    ] })
"""
SAMPLE_VENDORS = """
    ({ success: true, data: [
        { name: 'Acme Vendor', contact: '9876543210', address: '', gstin: '', remarks: '' },
        { name: 'No-Contact Vendor', contact: '', address: '', gstin: '', remarks: '' }
    ] })
"""
SAMPLE_ITEMS = """
    ({ success: true, data: [
        { name: 'Brake Pads', size: 'Standard', narration: '', baseUnit: 'Pcs', purchaseUnit: 'Pcs', weightPerBaseUnit: 0, vendors: [] }
    ] })
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
        page.evaluate(f"window.__mockRoutes.getPOData = {SAMPLE_POS};")
        page.evaluate(f"window.__mockRoutes.getVendorsData = {SAMPLE_VENDORS};")
        page.evaluate(f"window.__mockRoutes.getItemsData = {SAMPLE_ITEMS};")

        print("\n[Navigate] More tab -> open PO Ledger -> tap the New PO FAB")
        page.evaluate("MApp.Shell.showTab('more')")
        page.wait_for_timeout(50)
        page.evaluate("MApp.PO.openLedgerSheet()")
        page.wait_for_timeout(300)
        check(page.locator("#po-ledger-list .mb-card").count() == 1, "PO Ledger shows the 1 seeded PO")

        page.click("#sheet-po-ledger .mb-fab")
        page.wait_for_timeout(250)
        sheet_open = page.evaluate("document.getElementById('sheet-new-po').classList.contains('open')")
        check(sheet_open, "sheet-new-po opens on FAB tap")
        vendor_field = page.locator("#new-po-vendor-field")
        check(vendor_field.count() == 1, "vendor picker field rendered (reference data loaded)")

        print("\n[Validation] saving with nothing chosen is blocked, no API call fires")
        page.click("#new-po-save-btn")
        page.wait_for_timeout(50)
        calls_before = page.evaluate("window.__callCounts.savePO || 0")
        check(calls_before == 0, f"savePO NOT called yet (nothing chosen) (got {calls_before} call(s))")
        toast = page.locator("#mapp-toast-stack .mb-toast-error").last
        check("vendor" in (toast.text_content() or "").lower(), f"toast asks to choose a vendor first (got {toast.text_content()!r})")

        print("\n[Pick vendor] choosing a vendor with a known contact auto-fills the contact field")
        page.click("#new-po-vendor-field")
        page.wait_for_timeout(100)
        page.click("text=Acme Vendor")
        page.wait_for_timeout(100)
        vendor_label = page.evaluate("document.getElementById('new-po-vendor-field').textContent")
        check(vendor_label == 'Acme Vendor', f"vendor field shows the chosen vendor (got {vendor_label!r})")
        contact_val = page.evaluate("document.getElementById('new-po-contact').value")
        check(contact_val == '9876543210', f"contact auto-filled from vendor record (got {contact_val!r})")

        print("\n[Validation] vendor chosen but no item yet still blocks the save")
        page.click("#new-po-save-btn")
        page.wait_for_timeout(50)
        calls_after_vendor = page.evaluate("window.__callCounts.savePO || 0")
        check(calls_after_vendor == 0, f"savePO still not called (no item chosen) (got {calls_after_vendor} call(s))")

        print("\n[Pick item] choosing an item shows its name + size")
        page.click("#new-po-item-field")
        page.wait_for_timeout(100)
        page.click("text=Brake Pads")
        page.wait_for_timeout(100)
        item_label = page.evaluate("document.getElementById('new-po-item-field').textContent")
        check("Brake Pads" in item_label and "Standard" in item_label, f"item field shows name + size (got {item_label!r})")

        print("\n[Validation] zero quantity still blocks the save")
        page.fill("#new-po-qty", "0")
        page.click("#new-po-save-btn")
        page.wait_for_timeout(50)
        calls_after_zero_qty = page.evaluate("window.__callCounts.savePO || 0")
        check(calls_after_zero_qty == 0, f"savePO still not called (zero qty) (got {calls_after_zero_qty} call(s))")

        print("\n[Success] a valid submission saves, toasts the PO number, closes, and refreshes the ledger behind it")
        page.evaluate("window.__mockRoutes.savePO = { success: true, message: 'PO #PO-101 created successfully.', data: { poNumber: 'PO-101' } };")
        page.evaluate(f"window.__mockRoutes.getPOData = {{ success: true, data: [...{SAMPLE_POS}.data, {{ poNumber: 'PO-101', vendor: 'Acme Vendor', poDate: '13/07/2026', status: 'PO Issued', totalQty: 5, grandTotal: 500, items: [] }}] }};")
        page.fill("#new-po-qty", "5")
        page.fill("#new-po-price", "100")
        page.click("#new-po-save-btn")
        page.wait_for_timeout(300)

        args = page.evaluate("window.__lastArgs.savePO")
        check(args[0]['vendor'] == 'Acme Vendor', f"savePO called with the chosen vendor (got {args[0].get('vendor')!r})")
        check(args[0]['contact'] == '9876543210', f"savePO called with the auto-filled contact (got {args[0].get('contact')!r})")
        import json as _json
        items_sent = _json.loads(args[0]['items'])
        check(items_sent == [{'name': 'Brake Pads', 'size': 'Standard', 'narration': '', 'unit': 'Pcs', 'qty': 5, 'price': 100}],
              f"savePO called with exactly 1 item, correctly shaped (got {items_sent!r})")

        sheet_closed = page.evaluate("!document.getElementById('sheet-new-po').classList.contains('open')")
        check(sheet_closed, "sheet-new-po closes on success")
        success_toast = page.locator("#mapp-toast-stack .mb-toast-success").last
        check("PO-101" in (success_toast.text_content() or ""), f"success toast includes the new PO number (got {success_toast.text_content()!r})")
        check(page.locator("#po-ledger-list .mb-card").count() == 2, "PO Ledger behind the sheet now shows both POs (auto-refreshed)")

        print("\n[Failure] a rejected save keeps the sheet open with an error toast")
        page.click("#sheet-po-ledger .mb-fab")
        page.wait_for_timeout(250)
        page.click("#new-po-vendor-field")
        page.wait_for_timeout(100)
        page.click("text=Acme Vendor")
        page.wait_for_timeout(100)
        page.click("#new-po-item-field")
        page.wait_for_timeout(100)
        page.click("text=Brake Pads")
        page.wait_for_timeout(100)
        page.fill("#new-po-qty", "3")
        page.evaluate("window.__mockRoutes.savePO = { success: false, message: 'Failed to save PO: Vendor is required.' };")
        page.click("#new-po-save-btn")
        page.wait_for_timeout(250)
        sheet_still_open = page.evaluate("document.getElementById('sheet-new-po').classList.contains('open')")
        check(sheet_still_open, "sheet-new-po stays open after a failed save (nothing lost)")
        fail_toast = page.locator("#mapp-toast-stack .mb-toast-error").last
        check("vendor is required" in (fail_toast.text_content() or "").lower(), f"server's rejection message shown verbatim (got {fail_toast.text_content()!r})")
        save_btn_reenabled = page.evaluate("!document.getElementById('new-po-save-btn').disabled")
        check(save_btn_reenabled, "Save button re-enabled after a failed save")

        print("\n[Cancel] closing the sheet manually discards the in-progress form")
        page.evaluate("MApp.PO.closeNewSheet()")
        page.wait_for_timeout(50)
        check(page.evaluate("!document.getElementById('sheet-new-po').classList.contains('open')"), "sheet-new-po closes cleanly via Cancel/close")

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
