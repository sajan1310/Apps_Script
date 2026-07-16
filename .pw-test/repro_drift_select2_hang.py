"""
Repro script: does clicking the "Repoint To" Select2 dropdown in the
Item Reference Integrity modal hang the page?

Mocks getItemIdentityDriftReport with 2 distinct stale groups and a
realistic-sized App.State.globalItems list, opens the modal via
App.Item.checkReferenceIntegrity(), then clicks the Select2 widget for the
first group's dropdown and checks whether the page is still responsive
afterward (a hang would mean the post-click evaluate() call itself times out
or the dropdown never renders).
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

MOCK_ITEMS = [{"name": f"Item {i:04d}", "size": "General"} for i in range(3000)]

MOCK_DRIFT = []
for g in range(30):
    for sheet, ctx in [("Process Components", f"Process PRC-{g}"), ("Wastage Log", f"Wastage WST-{g}"), ("Bill Ledger", f"Bill #B-{g}")]:
        MOCK_DRIFT.append({"sheet": sheet, "context": ctx, "itemName": f"Stale Item {g}", "size": "500ml"})


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False, slow_mo=100)
        ctx = browser.new_context()
        page = ctx.new_page()

        console_errors = []
        page.on("console", lambda msg: console_errors.append(msg.text) if msg.type == "error" else None)
        page.on("pageerror", lambda exc: console_errors.append(f"pageerror: {exc}"))

        url = DIST_HTML.as_uri()
        page.goto(url, wait_until="domcontentloaded")
        page.wait_for_timeout(1500)

        page.evaluate(f"""
            App.State.globalItems = {json.dumps(MOCK_ITEMS)};
            Api.call = async (fn) => {{
                if (fn === 'getItemIdentityDriftReport') {{
                    return {{ success: true, data: {json.dumps(MOCK_DRIFT)}, message: '3 stale reference(s) found' }};
                }}
                return {{ success: false, message: 'unmocked: ' + fn }};
            }};
        """)

        print(f"[1] Opening Check Reference Integrity modal ({len(MOCK_ITEMS)} items, {len(MOCK_DRIFT)} findings)...")
        import time
        t_open = time.time()
        page.evaluate("App.Item.checkReferenceIntegrity()")
        modal = page.locator("#itemDriftModal")
        modal.wait_for(state="visible", timeout=TIMEOUT)
        print(f"  Modal visible (took {time.time()-t_open:.2f}s)")

        page.wait_for_timeout(500)
        rows = page.locator("#itemDriftModalBody tbody tr")
        print(f"  Row count: {rows.count()}")

        select2_container = page.locator("#itemDriftModalBody .select2-container").first
        print(f"  Select2 containers rendered: {page.locator('#itemDriftModalBody .select2-container').count()}")

        print("[2] Clicking the first Select2 dropdown...")
        try:
            select2_container.click(timeout=5000)
            print("  Click returned (no hang on click itself)")
        except Exception as e:
            print(f"  CLICK TIMED OUT / FAILED: {e}")

        print("[3] Checking page responsiveness after click...")
        import time
        t0 = time.time()
        try:
            result = page.evaluate("1 + 1")
            print(f"  Page still responsive, evaluate() returned: {result} (took {time.time()-t0:.2f}s)")
        except Exception as e:
            print(f"  PAGE UNRESPONSIVE: {e}")

        dropdown_open = page.locator(".select2-dropdown").count()
        print(f"  .select2-dropdown elements present: {dropdown_open}")
        if dropdown_open:
            search_visible = page.locator(".select2-search__field").first.is_visible()
            print(f"  Search field visible: {search_visible}")

        page.screenshot(path=str(Path(__file__).parent / "repro_drift_select2_hang.png"))
        print("  Screenshot saved")

        if console_errors:
            print("\n[Console errors/pageerrors captured]")
            for e in console_errors:
                print("  -", e)
        else:
            print("\n  No console errors captured")

        page.wait_for_timeout(1000)
        browser.close()


if __name__ == "__main__":
    run()
