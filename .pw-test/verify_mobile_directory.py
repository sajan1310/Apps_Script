"""
Verification script — Mobile Directory (More tab > Vendors/Clients/
Contractors), the one shared sheet-directory + MApp.Directory module reused
across all 3 entity types (see Mobile_Script.html for why: the 3 desktop
list APIs are structurally identical and none returns a pre-computed
outstanding figure, so this stays a simple read-only directory).

Covers: opening each of the 3 types routes to the right API and title,
search, tap-to-call tel: links, an entry with no contact on file, and the
empty/failure states.

Run: python .pw-test/verify_mobile_directory.py
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
                            setTimeout(() => {
                                const resp = window.__mockRoutes[prop];
                                if (resp && resp.__throwError) {
                                    if (state.failureCb) state.failureCb(new Error(resp.__throwError));
                                } else {
                                    successCb(resp);
                                }
                            }, 100);
                        };
                    }});
                    return proxy;
                }
            }
        }
    };
"""

SAMPLE_VENDORS = """
    ({ success: true, data: [
        { name: 'Acme Vendor', contact: '+91 98765 43210', address: 'Ludhiana', gstin: '', remarks: '' },
        { name: 'No-Contact Traders', contact: '', address: '', gstin: '', remarks: '' }
    ] })
"""
SAMPLE_CLIENTS = """
    ({ success: true, data: [ { name: 'Client A', contact: '9876500000', address: 'Delhi', gstin: '', remarks: '' } ] })
"""
SAMPLE_CONTRACTORS = """
    ({ success: true, data: [ { name: 'Contractor X', contact: '9998887770', address: '', gstPan: '', remarks: '' } ] })
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
        page.evaluate("MApp.Shell.showTab('more')")
        page.wait_for_timeout(50)

        print("\n[Vendors] opens with the right title, routes to getVendorsData")
        page.evaluate(f"window.__mockRoutes.getVendorsData = {SAMPLE_VENDORS};")
        page.evaluate("MApp.Directory.open('vendor')")
        page.wait_for_timeout(250)
        title = page.evaluate("document.getElementById('directory-title').textContent")
        check(title == 'Vendors', f"sheet title is 'Vendors' (got {title!r})")
        check(page.locator("#directory-list .mb-card").count() == 2, "2 vendor cards rendered")

        print("\n[Tap-to-call] a vendor with a contact renders a tel: link; one without shows a fallback")
        links = page.evaluate("""
            Object.fromEntries([...document.querySelectorAll('#directory-list .mb-card')].map(card => {
                const name = card.querySelector('.mb-card-title').textContent.trim();
                const a = card.querySelector('a[href^="tel:"]');
                return [name, a ? a.getAttribute('href') : null];
            }))
        """)
        check(links.get('Acme Vendor') == 'tel:+91 98765 43210', f"Acme Vendor's contact is a tel: link (got {links.get('Acme Vendor')!r})")
        check(links.get('No-Contact Traders') is None, "No-Contact Traders (empty contact) has no tel: link")
        no_contact_text = page.evaluate("""
            [...document.querySelectorAll('#directory-list .mb-card')]
                .find(c => c.querySelector('.mb-card-title').textContent.includes('No-Contact'))
                .querySelector('.mb-card-sub').textContent
        """)
        check('No contact on file' in no_contact_text, f"No-Contact Traders shows a fallback message (got {no_contact_text!r})")

        print("\n[Search] narrows by name")
        page.fill("#directory-search", "Acme")
        page.evaluate("MApp.Directory.onSearch(document.getElementById('directory-search').value)")
        page.wait_for_timeout(50)
        check(page.locator("#directory-list .mb-card").count() == 1, "searching 'Acme' narrows to 1 card")
        page.evaluate("MApp.Directory.close()")

        print("\n[Clients] opens with the right title, routes to getClientsData")
        page.evaluate(f"window.__mockRoutes.getClientsData = {SAMPLE_CLIENTS};")
        page.evaluate("MApp.Directory.open('client')")
        page.wait_for_timeout(250)
        check(page.evaluate("document.getElementById('directory-title').textContent") == 'Clients', "sheet title is 'Clients'")
        check(page.locator("#directory-list .mb-card").count() == 1, "1 client card rendered")
        page.evaluate("MApp.Directory.close()")

        print("\n[Contractors] opens with the right title, routes to getContractorsData")
        page.evaluate(f"window.__mockRoutes.getContractorsData = {SAMPLE_CONTRACTORS};")
        page.evaluate("MApp.Directory.open('contractor')")
        page.wait_for_timeout(250)
        check(page.evaluate("document.getElementById('directory-title').textContent") == 'Contractors', "sheet title is 'Contractors'")
        check(page.locator("#directory-list .mb-card").count() == 1, "1 contractor card rendered")

        print("\n[Empty state] zero entries shows the empty message")
        page.evaluate("window.__mockRoutes.getContractorsData = { success: true, data: [] };")
        page.evaluate("MApp.Directory.open('contractor')")
        page.wait_for_timeout(250)
        check(page.locator("#directory-list .mb-state").count() > 0, "empty state renders for zero contractors")

        print("\n[Failure state] a rejected call shows the retry error state")
        page.evaluate("window.__mockRoutes.getContractorsData = { __throwError: 'Simulated failure' };")
        page.evaluate("MApp.Directory.open('contractor')")
        page.wait_for_timeout(250)
        check(page.locator("#directory-list .mb-state-error").count() > 0, "error state renders on a rejected call")
        check(page.locator("#directory-list .mb-state-retry").count() > 0, "a Retry button is present")

        page.evaluate("MApp.Directory.close()")
        page.wait_for_timeout(50)
        check(page.evaluate("!document.getElementById('sheet-directory').classList.contains('open')"), "sheet-directory closes cleanly")

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
