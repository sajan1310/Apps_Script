"""
Verifies the live "Now producing: X / Y / Z" combination preview in the
Colors to Produce checklist, for a 3-axis process (Frame=primary,
Mudguard, Rim), matching the exact server-side combining rule.

Run: python .pw-test/verify_color_combination_preview.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(r"c:\Users\erkar\my-app-script-project\dist\index.html")
TIMEOUT = 8000

MOCK_PROCESS = {
    "processId": "PRC-FIT", "processName": "Fitting Frame 16 inch Crysta S/Rim", "sequence": 5,
    "lotPrefix": "FIT", "outputItemName": "Fitted Frame 16 inch Crysta S/Rim", "isFinalStage": False,
    "active": True, "processType": "General", "primaryColorAxis": ""
}
MOCK_COMPONENTS = [
    {"itemName": "Painted Frame", "size": "", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
    {"itemName": "Mudguard", "size": "", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
    {"itemName": "Rim", "size": "", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
]
MOCK_POOL = (
    [{"outputItemName": "Painted Frame", "color": c, "processId": "PRC-P", "qty": 10} for c in ["Blue-White", "Red-White"]]
    + [{"outputItemName": "Mudguard", "color": c, "processId": "PRC-M", "qty": 10} for c in ["Black", "White"]]
    + [{"outputItemName": "Rim", "color": c, "processId": "PRC-R", "qty": 10} for c in ["BCP", "Chrome"]]
)
MOCK_AXES = {
    "success": True,
    "data": {
        "axes": [
            {"key": "pool:painted frame", "label": "Painted Frame", "colors": ["Blue-White", "Red-White"], "source": "pool"},
            {"key": "pool:mudguard", "label": "Mudguard", "colors": ["Black", "White"], "source": "pool"},
            {"key": "pool:rim", "label": "Rim", "colors": ["BCP", "Chrome"], "source": "pool"}
        ],
        "primaryColorAxis": "Painted Frame",
        "primaryAxisKey": "pool:painted frame"
    }
}

MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": ["Blue-White", "Red-White", "Black", "White", "BCP", "Chrome"]},
    "getProcessColorAxes": MOCK_AXES,
    "getProcessComponentsData": {"success": True, "data": MOCK_COMPONENTS},
    "getWarehousePoolData": {"success": True, "data": MOCK_POOL},
    "getProcessWipData": {"success": True, "data": MOCK_POOL},
    "getStockData": {"success": True, "data": []},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
}


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
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
            App.State.globalItems = [];
            App.State.globalColors = [];
            App.State.globalContractors = [{{ contractorId: 'C1', name: 'Sanjay', active: true }}];
            App.State.globalProduction = [];
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

        ok = True
        print("[Step 1] Open Create modal, select the 3-axis process...")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(500)
        page.evaluate("""
            (() => {
                const sel = document.getElementById('productionProcessId');
                sel.value = 'PRC-FIT';
                sel.dispatchEvent(new Event('change', { bubbles: true }));
            })()
        """)
        page.wait_for_timeout(1000)

        def check_and_fill(color, qty):
            page.evaluate(f"""
                (() => {{
                    const row = Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row'))
                        .find(r => r.dataset.color === {json.dumps(color)});
                    const chk = row.querySelector('.production-color-check');
                    chk.checked = true;
                    chk.dispatchEvent(new Event('change', {{ bubbles: true }}));
                }})()
            """)
            page.wait_for_timeout(150)
            page.evaluate(f"""
                (() => {{
                    const row = Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row'))
                        .find(r => r.dataset.color === {json.dumps(color)});
                    const qtyInput = row.querySelector('.production-color-qty');
                    qtyInput.value = '{qty}';
                    qtyInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
                }})()
            """)
            page.wait_for_timeout(150)

        print("[Step 2] Check ONLY the primary (Blue-White) with qty 12 -- preview should show just that...")
        check_and_fill('Blue-White', 12)
        preview1 = page.evaluate("document.getElementById('productionColorCombinationPreview').innerText")
        print(f"  preview: {preview1!r}")
        if preview1.strip() != 'Now producing: Blue-White':
            print("  FAIL: expected 'Now producing: Blue-White'")
            ok = False
        else:
            print("  PASS")

        print("\n[Step 3] Check Mudguard=Black qty 12 too -- preview should now show 2-way combo...")
        check_and_fill('Black', 12)
        preview2 = page.evaluate("document.getElementById('productionColorCombinationPreview').innerText")
        print(f"  preview: {preview2!r}")
        if preview2.strip() != 'Now producing: Blue-White / Black':
            print("  FAIL: expected 'Now producing: Blue-White / Black'")
            ok = False
        else:
            print("  PASS")

        print("\n[Step 4] Check Rim=BCP qty 12 too -- preview should now show 3-way combo...")
        check_and_fill('BCP', 12)
        preview3 = page.evaluate("document.getElementById('productionColorCombinationPreview').innerText")
        print(f"  preview: {preview3!r}")
        if preview3.strip() != 'Now producing: Blue-White / Black / BCP':
            print("  FAIL: expected 'Now producing: Blue-White / Black / BCP'")
            ok = False
        else:
            print("  PASS")

        print("\n[Step 5] Also check Rim=Chrome (2nd Rim entry) -- now ambiguous, must NOT guess a combined string...")
        check_and_fill('Chrome', 5)
        preview4 = page.evaluate("document.getElementById('productionColorCombinationPreview').innerText")
        print(f"  preview: {preview4!r}")
        if 'Blue-White / Black / BCP' in preview4 and 'separately' not in preview4:
            print("  FAIL: should not show a fully-combined string once Rim has 2 checked entries (ambiguous)")
            ok = False
        elif 'Now producing: Blue-White' in preview4 and 'separately' in preview4:
            print("  PASS: correctly shows primary + ambiguity note instead of guessing")
        else:
            print(f"  FAIL: unexpected preview text: {preview4!r}")
            ok = False

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
