"""
Verification script — Fix #7 (Bill auto-match race condition), one of the 9
architectural gaps verified+fixed on 2026-07-13 (see
verification_2026_07_13_architectural_gaps in project memory).

Covers two distinct bugs in App.Bill.runAutoMatch/applyAutoMatch
(Script_Bill.html) + their trigger wiring (Script_Core.html):

1. Permanent-exclusion bug: once a row auto-matched to a real PO (dataset.po
   no longer 'DIRECT'), it was silently excluded from ever being
   re-evaluated again on further edits — the debounce-arming code and
   runAutoMatch's own candidate filter both gated on dataset.po === 'DIRECT'.
   A later qty/name edit should still refresh the suggestion.

2. Stale-response-reverts-manual-override bug: applyAutoMatch used to
   unconditionally overwrite every row in its captured snapshot, without
   re-checking whether the user committed a manual PO choice for that row
   WHILE the request was in flight. It now re-checks dataset.autoMatched at
   apply time and skips any row that's since become 'manual'.

Run: python .pw-test/verify_bill_automatch_race.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

failures = []
def check(cond, msg):
    print(("PASS: " if cond else "FAIL: ") + msg)
    if not cond:
        failures.append(msg)


# window.__matchResponses[callIndex] (0-based, in call order) controls what
# suggestPoAllocations resolves with on that Nth call; window.__matchDelayMs
# controls its artificial latency, both settable mid-test.
MOCK_JS = """
    window.__suggestCallCount = 0;
    window.__suggestCallArgs = [];
    window.__matchResponses = [];
    window.__matchDelayMs = 50;
    window.google = {
        script: {
            run: {
                withSuccessHandler(successCb) {
                    const state = { successCb, failureCb: null };
                    let proxy;
                    const runner = { withFailureHandler(failureCb) { state.failureCb = failureCb; return proxy; } };
                    proxy = new Proxy(runner, { get(target, prop) {
                        if (prop in target) return target[prop];
                        return (...args) => {
                            if (prop === 'suggestPoAllocations') {
                                const idx = window.__suggestCallCount;
                                window.__suggestCallCount++;
                                window.__suggestCallArgs.push(args);
                                const resp = window.__matchResponses[idx] || { success: true, data: [] };
                                setTimeout(() => successCb(resp), window.__matchDelayMs);
                                return;
                            }
                            // Everything else (getVendorsData, getItemsData, etc. used
                            // by page bootstrap) resolves immediately, empty.
                            setTimeout(() => successCb({ success: true, data: [] }), 5);
                        };
                    }});
                    return proxy;
                }
            }
        }
    };
"""


def suggest_response(row_uid, po_number, qty, unmatched_qty=0):
    allocs = [{"poNumber": po_number, "qty": qty, "rateConflict": None}]
    return {
        "success": True,
        "data": [{"rowIndex": row_uid, "allocations": allocs, "unmatchedQty": unmatched_qty}]
    }


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_context(viewport={"width": 1400, "height": 900}).new_page()
        page.goto(DIST_HTML.resolve().as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)
        page.evaluate(MOCK_JS)

        def open_bill_form():
            page.evaluate("App.Bill.openReceiveModal()")
            page.locator("#receiveBillModal").wait_for(state="visible", timeout=TIMEOUT)
            page.evaluate("""
                const sel = document.getElementById('billVendor');
                if (![...sel.options].some(o => o.value === 'Acme Vendor')) {
                    sel.add(new Option('Acme Vendor', 'Acme Vendor'));
                }
                sel.value = 'Acme Vendor';
            """)
            page.fill("#billItemsBody .b-item-name", "Steel Rod")
            page.fill("#billItemsBody .b-item-qty", "10")
            page.fill("#billItemsBody .b-item-price", "50")

        def row_uid():
            return page.evaluate("document.querySelector('#billItemsBody tr[data-row-uid]').dataset.rowUid")

        print("\n=== Test 1: an already-auto-matched row is NOT permanently excluded from further re-matching ===")
        open_bill_form()
        uid = row_uid()
        page.evaluate(f"window.__matchResponses = [{json.dumps(suggest_response(uid, 'PO-100', 10))}];")
        page.evaluate("App.Bill.runAutoMatch()")
        page.wait_for_timeout(250)

        po_after_first = page.evaluate("document.querySelector('#billItemsBody tr[data-row-uid]').dataset.po")
        check(po_after_first == "PO-100", f"first auto-match applies PO-100 (got {po_after_first!r})")

        # open_bill_form() fills name/qty/price, and each of those arms the
        # 400ms debounce. Those pending timers have to be drained BEFORE the
        # count below means anything: on a loaded machine one could still land
        # inside this step's own wait window, making the total 3 and failing a
        # `== 2` assertion that had nothing to do with re-matching. (Observed
        # intermittently, roughly 1 run in 9.) Wait for the count to go stable,
        # then assert on the DELTA — the same shape Test 3 below already uses.
        page.wait_for_function(
            """
            () => {
                const n = window.__suggestCallCount;
                const stable = window.__amPrev === n ? (window.__amStable || 0) + 1 : 0;
                window.__amPrev = n;
                window.__amStable = stable;
                return stable >= 3;
            }
            """,
            timeout=10000,
        )
        calls_before_edit = page.evaluate("window.__suggestCallCount")

        # Edit qty again -- under the OLD code this row (dataset.po no longer
        # 'DIRECT') would never re-arm the debounce timer at all.
        # The response is keyed off the live call count, not a hardcoded index,
        # so it still lines up whatever the settled baseline turned out to be.
        page.evaluate(f"window.__matchResponses[window.__suggestCallCount] = {json.dumps(suggest_response(uid, 'PO-200', 7))}")
        page.fill("#billItemsBody .b-item-qty", "7")
        page.wait_for_timeout(500)  # 400ms debounce + 50ms mock latency

        calls_after_edit = page.evaluate("window.__suggestCallCount")
        check(calls_after_edit == calls_before_edit + 1,
              f"editing qty on an already-matched row re-triggers runAutoMatch exactly once "
              f"(before={calls_before_edit}, after={calls_after_edit})")
        po_after_second = page.evaluate("document.querySelector('#billItemsBody tr[data-row-uid]').dataset.po")
        check(po_after_second == "PO-200", f"the refreshed suggestion (PO-200) was applied, not stuck on the stale PO-100 (got {po_after_second!r})")

        print("\n=== Test 2: a manual override committed WHILE a suggestion is in-flight is never reverted ===")
        page.evaluate("safeModalHide && safeModalHide('receiveBillModal')")
        page.wait_for_timeout(200)
        open_bill_form()
        uid2 = row_uid()

        # Slow response (300ms) so there's a real window to commit a manual
        # override before it resolves.
        page.evaluate("window.__matchDelayMs = 300")
        page.evaluate(f"window.__matchResponses = [{json.dumps(suggest_response(uid2, 'PO-300', 10))}];")
        page.evaluate("App.Bill.runAutoMatch(); void 0;")  # fire-and-forget, don't await

        page.wait_for_timeout(50)  # well before the mock's 300ms resolves
        page.evaluate("""
            const row = document.querySelector('#billItemsBody tr[data-row-uid]');
            row.dataset.po = 'PO-MANUAL';
            row.dataset.autoMatched = 'manual';
        """)

        page.wait_for_timeout(400)  # let the in-flight (now-stale) response land

        po_final = page.evaluate("document.querySelector('#billItemsBody tr[data-row-uid]').dataset.po")
        matched_final = page.evaluate("document.querySelector('#billItemsBody tr[data-row-uid]').dataset.autoMatched")
        check(po_final == "PO-MANUAL", f"the manual override survives the late-arriving stale suggestion (got dataset.po={po_final!r})")
        check(matched_final == "manual", f"dataset.autoMatched stays 'manual', not silently flipped back to 'auto' (got {matched_final!r})")

        print("\n=== Test 3: submitting the bill flushes a pending debounced auto-match first ===")
        page.evaluate("safeModalHide && safeModalHide('receiveBillModal')")
        page.wait_for_timeout(200)
        page.evaluate("window.__matchDelayMs = 50")
        open_bill_form()
        uid3 = row_uid()
        call_count_before_edit = page.evaluate("window.__suggestCallCount")
        resp_json = json.dumps(suggest_response(uid3, "PO-400", 10))
        page.evaluate(f"window.__matchResponses[window.__suggestCallCount] = {resp_json}")

        page.fill("#billItemsBody .b-item-name", "Steel Rod 2")  # arms the 400ms debounce, does NOT fire yet
        page.fill("#receiveBillModal input[name='billNumber']", "INV-RACE")

        # Submit immediately -- well before the 400ms debounce would have
        # fired on its own.
        page.click("#billSubmitBtn")
        page.wait_for_timeout(600)

        calls_after_submit = page.evaluate("window.__suggestCallCount")
        check(calls_after_submit > call_count_before_edit, f"submitting flushed the pending debounced auto-match instead of leaving it to fire later (calls before={call_count_before_edit}, after={calls_after_submit})")

        confirm_visible = page.locator("#confirmModal").is_visible()
        check(confirm_visible, "the confirm dialog still appears after the flush resolves (submit flow isn't broken)")

        browser.close()

        if failures:
            print(f"\n{len(failures)} CHECK(S) FAILED")
        else:
            print("\nALL CHECKS PASSED")
        return not failures


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
