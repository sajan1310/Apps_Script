"""
Verifies the Ready to Dispatch table shows a "Colors" button when a Product
Tag's colorBreakdown has real color info, and clicking it opens the modal
with the correct per-color rows (produced/dispatched/ready).

Run: python .pw-test/verify_dispatch_color_breakdown_ui.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(r"c:\Users\erkar\my-app-script-project\dist\index.html")
TIMEOUT = 8000

MOCK_READY = [
    {
        "productId": "PROD-001", "productName": "Bicycle 16 inch Crysta", "producedQty": 25,
        "dispatchedQty": 0, "readyQty": 25,
        "colorBreakdown": [
            {"color": "Blue-White / BCP", "producedQty": 15, "dispatchedQty": 0, "readyQty": 15},
            {"color": "Red-White / Black", "producedQty": 10, "dispatchedQty": 0, "readyQty": 10},
        ]
    },
    {
        "productId": "PROD-002", "productName": "Bicycle 20 inch Ford", "producedQty": 8,
        "dispatchedQty": 0, "readyQty": 8,
        "colorBreakdown": [
            {"color": "", "producedQty": 8, "dispatchedQty": 0, "readyQty": 8},
        ]
    }
]


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context()
        page = ctx.new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        page.evaluate(f"""
            App.State.globalReadyToDispatch = {json.dumps(MOCK_READY)};
            App.State.filteredReadyToDispatch = {json.dumps(MOCK_READY)};
            App.Dispatch.renderReadyTable();
        """)
        page.wait_for_timeout(300)

        ok = True

        print("[Check] PROD-001 (2 real colors) shows a Colors button...")
        prod1_btn_text = page.evaluate("""
            (() => {
                const rows = Array.from(document.querySelectorAll('#readyToDispatchTableBody tr'));
                const row = rows.find(r => r.textContent.includes('PROD-001'));
                const btn = row ? row.querySelector('button.btn-outline-dark') : null;
                return btn ? btn.textContent.trim() : null;
            })()
        """)
        print(f"  button text: {prod1_btn_text!r}")
        if not prod1_btn_text or '2' not in prod1_btn_text:
            print("  FAIL: expected a Colors button showing count 2")
            ok = False
        else:
            print("  PASS: Colors button shows count 2")

        print("\n[Check] PROD-002 (only 1 entry, blank color) shows a dash, no button...")
        prod2_has_btn = page.evaluate("""
            (() => {
                const rows = Array.from(document.querySelectorAll('#readyToDispatchTableBody tr'));
                const row = rows.find(r => r.textContent.includes('PROD-002'));
                return row ? !!row.querySelector('button.btn-outline-dark') : null;
            })()
        """)
        print(f"  has Colors button: {prod2_has_btn}")
        if prod2_has_btn:
            print("  FAIL: expected NO Colors button for a single blank-color entry")
            ok = False
        else:
            print("  PASS: no Colors button shown for a color-less product")

        print("\n[Check] Clicking PROD-001's Colors button opens the modal with correct rows...")
        page.evaluate("App.Dispatch.openColorBreakdown('PROD-001')")
        page.wait_for_timeout(500)

        modal_visible = page.locator("#dispatchColorBreakdownModal").is_visible()
        print(f"  modal visible: {modal_visible}")
        ok = ok and modal_visible

        body_text = page.evaluate("document.getElementById('dispatchColorBreakdownBody').innerText")
        print(f"  modal body text:\n{body_text}")
        if "Blue-White / BCP" not in body_text or "15" not in body_text:
            print("  FAIL: expected 'Blue-White / BCP' row with qty 15")
            ok = False
        else:
            print("  PASS: Blue-White / BCP row with qty 15 present")
        if "Red-White / Black" not in body_text or "10" not in body_text:
            print("  FAIL: expected 'Red-White / Black' row with qty 10")
            ok = False
        else:
            print("  PASS: Red-White / Black row with qty 10 present")

        if console_errors:
            print("\n  Console/page errors:")
            for e in console_errors:
                print(f"    {e}")
            ok = False

        browser.close()
        return ok


if __name__ == "__main__":
    ok = run()
    print("\n" + ("ALL PASS" if ok else "SOME FAILED"))
    sys.exit(0 if ok else 1)
