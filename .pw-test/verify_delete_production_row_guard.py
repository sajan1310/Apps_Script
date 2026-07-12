"""
Verification script: the Production Log table's single-row Delete button
must forward the same row-shift safety arguments (expectedProductId,
expectedQty) that deleteProduction(rowIdx, expectedProductId, expectedQty)
(module_production.js) uses to detect a stale/shifted row — the same
protection App.Production.updateStatus, saveProductionSheet, and the
analogous App.Dispatch.delete already provide.

Reproduces the bug: the Delete button's onclick only passed rowIdx, so the
server's "Data mismatch: The record has been modified or shifted" guard was
permanently dead code for single-row deletes.

Run: python .pw-test/verify_delete_production_row_guard.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

MOCK_ROW = {
    "rowIdx": 7, "date": "01/07/2026", "dateRaw": "2026-07-01T00:00:00.000Z",
    "productId": "PRD-1001", "productName": "Test Product", "qty": 12,
    "assignedBy": "Super", "assignedTo": "Contractor A", "status": "Pending",
    "remarks": "", "processId": "PRC-1", "lotNumber": "LOT-FF-0001",
    "contractorRate": 0, "contractorPayable": 0, "outputItemName": "Fitted Frame",
    "componentsConsumed": [], "color": "", "colorBreakdown": []
}


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_context().new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.resolve().as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        page.evaluate(f"""
            window.__calls = [];
            App.State.globalProduction = [{json.dumps(MOCK_ROW)}];
            App.State.filteredProduction = App.State.globalProduction;
            App.State.globalProcesses = [];
            window.google = {{
                script: {{
                    run: {{
                        withSuccessHandler(cb) {{
                            let proxy;
                            const target = {{ withFailureHandler() {{ return proxy; }} }};
                            proxy = new Proxy(target, {{
                                get(t, method) {{
                                    if (method in t) return t[method];
                                    return (...args) => {{
                                        window.__calls.push({{ method, args }});
                                        setTimeout(() => cb({{ success: true, message: 'ok' }}), 10);
                                    }};
                                }}
                            }});
                            return proxy;
                        }}
                    }}
                }}
            }};
            App.Production.renderTable();
        """)

        failures = []
        def check(cond, msg):
            print(("PASS: " if cond else "FAIL: ") + msg)
            if not cond:
                failures.append(msg)

        print("\n[Step 1] Read the rendered Delete button's onclick and invoke it (production tab is not the active tab in this harness, so click() can't reach it)...")
        onclick = page.locator("#productionTableBody button.btn-danger").first.get_attribute("onclick")
        print(f"  onclick=\"{onclick}\"")
        page.evaluate(onclick)
        page.wait_for_timeout(300)

        confirm_visible = page.locator("#confirmModal").is_visible()
        check(confirm_visible, "Confirm dialog opened")

        print("\n[Step 2] Confirm the deletion...")
        page.evaluate("App.State.confirmCallback && App.State.confirmCallback();")
        page.wait_for_timeout(300)

        calls = page.evaluate("window.__calls")
        delete_calls = [c for c in calls if c["method"] == "deleteProduction"]
        check(len(delete_calls) == 1, f"deleteProduction called exactly once (got {len(delete_calls)})")

        if delete_calls:
            args = delete_calls[0]["args"]
            print(f"  deleteProduction called with args: {args}")
            check(len(args) == 3, f"3 arguments passed (rowIdx, expectedProductId, expectedQty) (got {len(args)})")
            check(str(args[0]) == str(MOCK_ROW["rowIdx"]), f"rowIdx matches (got {args[0] if args else None})")
            if len(args) >= 3:
                check(str(args[1]) == MOCK_ROW["productId"], f"expectedProductId matches saved productId (got {args[1]})")
                check(str(args[2]) == str(MOCK_ROW["qty"]), f"expectedQty matches saved qty (got {args[2]})")

        if console_errors:
            print("\n⚠️  Console/page errors:")
            for e in console_errors:
                print(f"    {e}")
            failures.append("console errors present")

        print("\n" + ("ALL TESTS PASSED" if not failures else f"{len(failures)} TEST(S) FAILED"))
        browser.close()
        return len(failures) == 0


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
