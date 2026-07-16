"""
Verification script: Item Reference Integrity dialog (App.Item.checkReferenceIntegrity
/ fixDriftReference / _removeDriftRow, module_items.js#getItemIdentityDriftReport /
fixItemIdentityDriftReference).

Tests:
  1. Modal opens with one row per distinct stale identity (3 findings -> 2 groups),
     each Fix button starts disabled.
  2. Picking a target in a row's Select2 enables that row's Fix button (and only
     that row's).
  3. Clicking Fix removes just that row (not a full re-fetch/re-render) and leaves
     the other row's dropdown untouched/still usable.
  4. Fixing the last remaining row closes the modal.
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

MOCK_ITEMS = [
    {"name": "Paint Thinner", "size": "1L"},
    {"name": "Rusty Bolt Set", "size": "M8"},
]

MOCK_DRIFT = [
    {"sheet": "Wastage Log", "context": "Wastage WST-1", "itemName": "Old Paint Name", "size": "500ml"},
    {"sheet": "Process Components", "context": "Process PRC-1", "itemName": "Old Paint Name", "size": "500ml"},
    {"sheet": "Bill Ledger", "context": "Bill #B-9", "itemName": "Old Bolt Name", "size": "M8"},
]


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context()
        page = ctx.new_page()

        console_errors = []
        page.on("pageerror", lambda exc: console_errors.append(f"pageerror: {exc}"))

        url = DIST_HTML.as_uri()
        page.goto(url, wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        fix_calls = []
        page.expose_function("recordFixCall", lambda args: fix_calls.append(args))

        page.evaluate(f"""
            App.State.globalItems = {json.dumps(MOCK_ITEMS)};
            let driftCallCount = 0;
            Api.call = async (fn, ...args) => {{
                if (fn === 'getItemIdentityDriftReport') {{
                    driftCallCount++;
                    window.__driftCallCount = driftCallCount;
                    return {{ success: true, data: {json.dumps(MOCK_DRIFT)}, message: '3 stale reference(s) found' }};
                }}
                if (fn === 'fixItemIdentityDriftReference') {{
                    await recordFixCall(args);
                    return {{ success: true, message: 'Repointed all references from "' + args[0] + '" to "' + args[2] + '".' }};
                }}
                return {{ success: false, message: 'unmocked: ' + fn }};
            }};
        """)

        print("[1] Opening Check Reference Integrity modal...")
        page.evaluate("App.Item.checkReferenceIntegrity()")
        modal = page.locator("#itemDriftModal")
        modal.wait_for(state="visible", timeout=TIMEOUT)
        rows = page.locator("#itemDriftModalBody tbody tr")
        assert rows.count() == 2, f"expected 2 grouped rows (3 findings -> 2 distinct identities), got {rows.count()}"
        print(f"  ✅ 2 grouped rows rendered (3 raw findings deduped correctly)")

        fix_buttons = page.locator(".drift-fix-btn")
        assert fix_buttons.nth(0).is_disabled(), "Fix button should start disabled before a target is picked"
        assert fix_buttons.nth(1).is_disabled(), "Second row's Fix button should also start disabled"
        print("  ✅ Both Fix buttons start disabled")

        print("[2] Picking a target in row 0's Select2...")
        page.locator("#itemDriftModalBody .select2-container").nth(0).click()
        page.wait_for_timeout(300)
        page.locator(".select2-results__option").first.click()
        page.wait_for_timeout(200)

        assert not fix_buttons.nth(0).is_disabled(), "row 0's Fix button should enable after picking a target"
        assert fix_buttons.nth(1).is_disabled(), "row 1's Fix button must stay disabled (independent rows)"
        print("  ✅ Only row 0's Fix button enabled after its own pick")

        print("[3] Clicking Fix on row 0...")
        fix_buttons.nth(0).click()
        page.wait_for_timeout(300)

        assert rows.count() == 1, f"expected exactly 1 row left after fixing row 0, got {rows.count()}"
        print("  ✅ Row 0 removed locally, row 1 still present")

        drift_call_count = page.evaluate("window.__driftCallCount")
        assert drift_call_count == 1, f"getItemIdentityDriftReport should have been called exactly once (no re-fetch on fix), got {drift_call_count}"
        print("  ✅ No full re-fetch happened on fix (optimistic local removal confirmed)")

        summary_text = page.locator("#itemDriftSummary").inner_text()
        assert "1 distinct item identity remaining" in summary_text, f"summary should reflect 1 remaining, got: {summary_text}"
        print(f"  ✅ Summary updated: \"{summary_text}\"")

        print("[4] Fixing the last remaining row...")
        page.locator("#itemDriftModalBody .select2-container").first.click()
        page.wait_for_timeout(300)
        page.locator(".select2-results__option").first.click()
        page.wait_for_timeout(200)
        page.locator(".drift-fix-btn").first.click()
        page.wait_for_timeout(500)

        assert not modal.is_visible(), "modal should auto-close once the last row is fixed"
        print("  ✅ Modal closed after fixing the last row")

        assert len(fix_calls) == 2, f"expected exactly 2 fixItemIdentityDriftReference calls, got {len(fix_calls)}"
        print(f"  ✅ fixItemIdentityDriftReference called twice, args: {fix_calls}")

        if console_errors:
            print("\n[Console/page errors]")
            for e in console_errors:
                print("  -", e)
            raise AssertionError("page errors occurred during test")
        else:
            print("\n  ✅ No console/page errors")

        print("\nALL CHECKS PASSED")
        browser.close()


if __name__ == "__main__":
    run()
