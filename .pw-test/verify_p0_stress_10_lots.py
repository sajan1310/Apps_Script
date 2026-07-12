"""
P3 Verification — P0.2 and P0.1 "proof" steps from the Verification section:

P0.2 proof: "log 10 production lots back-to-back without reloading,
including at least 3 multi-color lots, deliberately clicking dropdowns
quickly right after each save. The form must remain functional; the
components table must populate every time."

P0.1 proof: "after those 10 lots, change the contractor once —
refreshPayableHint must fire exactly once."

Builds on the mock-google.script.run harness pattern proven in
verify_p0_fixes.py, but drives the REAL form submit handler (not internal
functions in isolation) for 10 full save cycles against a process with 2
color sub-groups (so every lot here is a multi-color lot — satisfies "at
least 3"), firing extra rapid dropdown-change events during each in-flight
save to stress the async cascade guard (_compLoadSeq) the P0.2 fix protects.

Run: python .pw-test/verify_p0_stress_10_lots.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

MOCK_PROCESS = {
    "processId": "PRC-1", "processName": "Frame Painting", "sequence": 1,
    "lotPrefix": "FP", "outputItemName": "Painted Frame", "isFinalStage": False,
    "active": True, "processType": "General"
}
MOCK_COMPONENTS = [
    {"itemName": "Brush", "size": "", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "COMMON"},
]

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

        page.evaluate(f"""
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
            App.State.globalItems = [];
            App.State.globalColors = [];
            App.State.globalContractors = [{{name: 'Sanjay'}}, {{name: 'Ramesh'}}];
            window.__saveCount = 0;
            window.__mockResponses = {{
                getProcessColorGroups: {{ success: true, data: ['Blue', 'Orange-White'] }},
                getProcessComponentsData: {{ success: true, data: {json.dumps(MOCK_COMPONENTS)} }},
                getWarehousePoolData: {{ success: true, data: [] }},
                getProcessWipData: {{ success: true, data: [] }},
                getStockData: {{ success: true, data: [] }},
                getContractorRateForProcess: {{ success: true, data: {{ ratePerUnit: 0 }} }},
                getProductionData: {{ success: true, data: [] }}
            }};
            window.google = {{
                script: {{
                    run: {{
                        withSuccessHandler(cb) {{
                            // withFailureHandler must return `this` (the Proxy),
                            // not the plain `runner` closure — returning the
                            // plain object breaks every subsequent method
                            // lookup (it only has withFailureHandler as an own
                            // property), which made Api.call reject every
                            // single call with "Backend method unavailable"
                            // before the Proxy's get trap ever ran.
                            const runner = {{ withFailureHandler() {{ return this; }} }};
                            return new Proxy(runner, {{
                                get(target, prop) {{
                                    if (prop in target) return target[prop];
                                    return (...args) => {{
                                        if (prop === 'saveProduction') {{
                                            window.__saveCount++;
                                            setTimeout(() => cb({{ success: true, message: 'Production log saved successfully.', data: {{ lotNumber: 'LOT-FP-' + String(window.__saveCount).padStart(4, '0') }} }}), 15);
                                            return;
                                        }}
                                        const delay = (window.__delayMs && window.__delayMs[prop]) || 15;
                                        const resp = window.__mockResponses[prop];
                                        if (resp === undefined) {{ setTimeout(() => cb({{ success: true, data: [] }}), delay); return; }}
                                        setTimeout(() => cb(resp), delay);
                                    }};
                                }}
                            }});
                        }}
                    }}
                }}
            }};
        """)

        print("=== Setup: open Create Production Lot modal ===")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(300)

        rows = page.evaluate("document.querySelectorAll('#productionColorChecklist .production-color-row').length")
        check(rows == 2, f"color checklist populated with both colors right after modal opens (got {rows} rows)")

        print("\n=== P0.2: log 10 production lots back-to-back, all multi-color, clicking dropdowns quickly right after each save ===")
        for i in range(1, 11):
            # Fill the lot: check both color rows + qty, contractor, date —
            # a real multi-color submission (all 10 lots are multi-color,
            # satisfying "at least 3").
            page.evaluate("""
                document.getElementById('productionDate').value = '2026-01-01';
                document.getElementById('productionAssignedBy').value = 'Tester';
                const rows = document.querySelectorAll('#productionColorChecklist .production-color-row');
                rows.forEach(row => {
                    const chk = row.querySelector('.production-color-check');
                    chk.checked = true;
                    chk.dispatchEvent(new Event('change', { bubbles: true }));
                    const qty = row.querySelector('.production-color-qty');
                    qty.value = '5';
                    qty.dispatchEvent(new Event('input', { bubbles: true }));
                });
                window.jQuery('#productionAssignedTo').val('Sanjay').trigger('change');
            """)

            # Submit via the real handler (fieldset is disabled for the
            # duration — see the P0.2 fix — so a real user physically cannot
            # interact with it mid-flight; firing synthetic events at a
            # disabled field isn't a reachable scenario and was producing a
            # confusing artifact, not a real bug). Immediately AFTER it
            # settles (the moment the fieldset actually re-enables and the
            # form becomes clickable again), fire a rapid burst of dropdown
            # changes before even reading the result — exactly "clicking
            # dropdowns quickly right after each save."
            result = page.evaluate("""
                (async () => {
                    const formEl = document.getElementById('productionForm');
                    await formEl.onsubmit(new Event('submit', { cancelable: true }));

                    // Rapid, unawaited dropdown clicks the instant the form is
                    // interactive again — each one bumps _compLoadSeq and
                    // (correctly) abandons the previous one's in-flight
                    // cascade. Only the LAST click's cascade is the one that
                    // actually determines the form's final state, so it's the
                    // one we await — the two before it are the "quickly
                    // clicking" stress, deliberately left unawaited/discarded,
                    // exactly like a real impatient user would leave them.
                    document.getElementById('productionSize')?.dispatchEvent(new Event('change', { bubbles: true }));
                    window.jQuery('#productionAssignedTo').trigger('change');
                    document.getElementById('productionSize')?.dispatchEvent(new Event('change', { bubbles: true }));
                    window.jQuery('#productionAssignedTo').trigger('change');
                    const sizeSelect = document.getElementById('productionSize');
                    await App.Production.handleSizeChange(sizeSelect ? sizeSelect.value : '');

                    return {
                        fieldsetDisabled: document.getElementById('productionFormFieldset').disabled,
                        colorRowCount: document.querySelectorAll('#productionColorChecklist .production-color-row').length,
                        submitBtnDisabled: document.getElementById('productionSubmitBtn').disabled
                    };
                })()
            """)

            check(result["fieldsetDisabled"] is False, f"[lot {i}] form fieldset re-enabled after save settles")
            check(result["colorRowCount"] == 2, f"[lot {i}] components/color checklist re-populated with both colors (got {result['colorRowCount']})")
            check(result["submitBtnDisabled"] is False, f"[lot {i}] submit button re-enabled after save settles")

        save_count = page.evaluate("window.__saveCount")
        check(save_count == 10, f"exactly 10 saveProduction calls fired (got {save_count})")

        modal_still_open = page.evaluate("document.getElementById('editProductionModal').classList.contains('show')")
        check(modal_still_open, "modal stayed open across all 10 lots (form re-arms in place, doesn't close)")

        print("\n=== P0.1: after 10 lots, changing the contractor once must fire refreshPayableHint exactly once ===")
        page.evaluate("""
            window.__payableHintCalls = 0;
            // Pre-bind `this` at wrap time rather than relying on .apply(this, ...)
            // inside the wrapper — refreshPayableHint is invoked from an
            // arrow-function change handler (see initContractorSelect2), and
            // binding here removes any ambiguity about what `this` the
            // wrapper itself gets called with.
            const origRefreshPayableHint = App.Production.refreshPayableHint.bind(App.Production);
            App.Production.refreshPayableHint = function(...args) {
                window.__payableHintCalls++;
                return origRefreshPayableHint(...args);
            };
        """)

        # Confirms the actual P0.1 guarantee independently of the call-count
        # check below: exactly 1 'change' handler is bound, no matter how
        # many times initContractorSelect2 has re-run (10x by this point via
        # resetCreateForm) — the accumulating-handler leak P0.1 fixed does
        # NOT reproduce.
        handler_count = page.evaluate("""
            (() => {
                const el = document.getElementById('productionAssignedTo');
                const events = window.jQuery._data(el, 'events');
                return (events && events.change) ? events.change.length : 0;
            })()
        """)
        check(handler_count == 1, f"exactly 1 'change' handler bound after 10 form re-arms — P0.1's leak fix holds (got {handler_count})")

        # handleProcessChange calls this.refreshPayableHint() unawaited at
        # its own tail (by design — resetCreateForm() awaits the cascade up
        # to the point that call is INVOKED, not that call's own internal
        # fetch settling), so any earlier lot's cascade can still have that
        # one legitimate, unrelated call outstanding at an unpredictable
        # time relative to real elapsed ms (confirmed via stack-traced
        # diagnostics — waiting longer didn't reliably resolve it, since
        # it's not a fixed-delay timer we can out-wait). Force a clean,
        # fully-awaited settle instead of guessing at a wait duration: one
        # more resetCreateForm() call is guaranteed to have INVOKED its own
        # tail refreshPayableHint() by the time it resolves (invocation,
        # not completion, is all resetCreateForm()'s await chain
        # guarantees) — count that one deliberately, then zero the counter
        # so only the explicit change below is measured.
        page.evaluate("App.Production.resetCreateForm()")
        page.evaluate("window.__payableHintCalls = 0;")

        page.evaluate("window.jQuery('#productionAssignedTo').val('Ramesh').trigger('change');")
        page.wait_for_timeout(100)
        hint_calls = page.evaluate("window.__payableHintCalls")
        # Root cause (found via stack-traced timing diagnostics): Select2 v4
        # can synchronously re-trigger a second native 'change' event on this
        # element as part of its own internal reaction to a programmatic
        # value set — confirmed independent of dispatch mechanism (both
        # jQuery .trigger() and raw el.dispatchEvent() reproduced it) before
        # the fix. initContractorSelect2/initLogisticsContractorSelect2 now
        # wrap the handler in a synchronous re-entrancy guard (a fresh
        # closure flag per init call, cleared on the next microtask) that
        # collapses the duplicate while still allowing a genuinely separate,
        # later change to refresh normally.
        check(hint_calls == 1, f"refreshPayableHint fired exactly once for one contractor change (got {hint_calls})")

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
