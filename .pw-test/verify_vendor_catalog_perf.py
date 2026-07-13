"""
Verification script — App.PO.openVendorCatalog() performance refactor
(Script_PO.html). Last-purchase-date lookup is now a Map<vendor|name|size|
narration, date> built once from all POs, instead of a triple-nested scan
(items x vendors x POs, with a per-PO .some() over that PO's items) run per
item x vendor row. Confirms: the most RECENT PO's date wins when an item was
bought from the same vendor on multiple POs, and an item x vendor pair with
no matching PO history shows a blank last-purchase date.

Run: python .pw-test/verify_vendor_catalog_perf.py
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


SEED = """
    App.State.globalItems = [
      { name: 'Widget', size: '', narration: '', vendors: [ { vendor: 'Acme Vendor', rate: 50 } ] },
      { name: 'Gadget', size: 'L', narration: '', vendors: [ { vendor: 'Acme Vendor', rate: 20 } ] }
    ];
    App.State.globalPOs = [
      { poNumber: 'PO-1', poDate: '01/01/2026', poDateRaw: '2026-01-01T00:00:00.000Z', vendor: 'Acme Vendor',
        items: [ { name: 'Widget', size: '', narration: '' } ] },
      { poNumber: 'PO-2', poDate: '15/03/2026', poDateRaw: '2026-03-15T00:00:00.000Z', vendor: 'Acme Vendor',
        items: [ { name: 'Widget', size: '', narration: '' } ] }
    ];
"""


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context()
        page = ctx.new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(500)

        page.evaluate("""
            window.google = {
                script: {
                    run: {
                        withSuccessHandler(cb) {
                            const runner = { withFailureHandler() { return runner; } };
                            return new Proxy(runner, { get(target, prop) {
                                if (prop in target) return target[prop];
                                return (...args) => {};
                            }});
                        }
                    }
                }
            };
        """)

        print("\n[Seed] Widget bought from Acme on 2 POs (Jan + Mar); Gadget never bought from Acme")
        page.evaluate(SEED)

        page.evaluate("App.PO.openVendorCatalog();")
        page.wait_for_timeout(50)

        rows = page.locator("#vendorCatalogBody tr")
        row_texts = [rows.nth(i).inner_text().replace("\n", " | ") for i in range(rows.count())]
        joined = " || ".join(row_texts)

        widget_row = next((t for t in row_texts if "Widget" in t), None)
        gadget_row = next((t for t in row_texts if "Gadget" in t), None)

        check(widget_row is not None, f"Widget row present (rows: {joined})")
        check(widget_row is not None and "15/03/2026" in widget_row,
              f"Widget shows the MOST RECENT PO date (15/03/2026, not the older 01/01/2026) (row: {widget_row})")

        check(gadget_row is not None, f"Gadget row present (rows: {joined})")
        if gadget_row is not None:
            cells = rows.nth(row_texts.index(gadget_row)).locator("td")
            last_date_cell = cells.nth(5).inner_text().strip()
            check(last_date_cell == "", f"Gadget (no PO history) has a blank last-purchase date (got '{last_date_cell}')")

        check(len(console_errors) == 0, f"no page errors ({console_errors})")

        browser.close()


if __name__ == "__main__":
    run()
    print("\nALL TESTS PASSED" if failures == 0 else f"\n{failures} TEST(S) FAILED")
    sys.exit(1 if failures else 0)
