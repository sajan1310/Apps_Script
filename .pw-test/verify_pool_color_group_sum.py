"""
Verification script: Per-Process Pool Components tables' suggested quantities
when a process draws from 2+ independently-colored Warehouse Pool inputs at
once (a composite output color, e.g. "BCP / Blue-White" — see
COLOR_COMBO_DELIMITER in config.js).

Reproduces the reported bug: with colors like BCP/Blue-White, BCP/Orange-White,
..., Black/Sea-Green-White all checked at qty 25 each, every column in both
Per-Process Pool Components tables (one for the "Blue-White"-style axis, one
for the "BCP"/"Black" axis) showed a flat 25 instead of the sum of every
composite color sharing that column — e.g. the "Blue-White" column should be
25 (BCP/Blue-White) + 25 (Black/Blue-White) = 50, and the "BCP" column should
be 25 x 6 White-shades = 150.

Run: python .pw-test/verify_pool_color_group_sum.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

WHITE_SHADES = ["Blue-White", "Orange-White", "Pink-White", "Purple-White", "Red-White", "Sea Green-White"]
RIM_COLORS = ["BCP", "Black"]
COMPOSITE_COLORS = [f"{rim} / {shade}" for rim in RIM_COLORS for shade in WHITE_SHADES]

MOCK_PROCESS = {
    "processId": "PRC-1", "processName": "Fitting Frame 16 inch Crysta S/Rim", "sequence": 2,
    "lotPrefix": "FF", "outputItemName": "Fitted Frame 16 inch Crysta S/Rim", "isFinalStage": False,
    "active": True, "processType": "General"
}
MOCK_COMPONENTS = [
    {"itemName": "Painted Frame Crysta 16 inch D/Gaddi", "size": "16 inch", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
    {"itemName": "Fitted Rim 16 inch", "size": "16 inch", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
]
MOCK_POOL_ROWS = (
    [{"outputItemName": "Painted Frame Crysta 16 inch D/Gaddi", "color": c, "qty": 1000} for c in WHITE_SHADES] +
    [{"outputItemName": "Fitted Rim 16 inch", "color": c, "qty": 1000} for c in RIM_COLORS]
)
MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": COMPOSITE_COLORS},
    "getProcessComponentsData": {"success": True, "data": MOCK_COMPONENTS},
    "getWarehousePoolData": {"success": True, "data": MOCK_POOL_ROWS},
    "getProcessWipData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
}


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_context(viewport={"width": 1600, "height": 1000}).new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.resolve().as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        page.evaluate(f"""
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
            App.State.globalItems = [];
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

        failures = []
        def check(cond, msg):
            print(("PASS: " if cond else "FAIL: ") + msg)
            if not cond:
                failures.append(msg)

        print("\n[Step 1] Open Create modal, select the process...")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.evaluate("document.getElementById('productionProcessId').value='PRC-1'; App.Production.handleProcessChange('PRC-1');")
        page.wait_for_timeout(800)

        rows = page.locator("#productionColorChecklist .production-color-row")
        check(rows.count() == 12, f"12 composite color rows rendered (got {rows.count()})")

        print("\n[Step 2] Check all 12 composite colors at qty=25 each...")
        for color in COMPOSITE_COLORS:
            row = page.locator(f"#productionColorChecklist .production-color-row[data-color='{color}']")
            row.locator(".production-color-check").check()
            page.wait_for_timeout(80)
            row.locator(".production-color-qty").fill("25")
            page.wait_for_timeout(80)
        page.wait_for_timeout(500)

        print("\n[Step 3] Read Per-Process Pool Components table cells...")
        pool_values = page.evaluate("""
            (() => {
                const out = {};
                document.querySelectorAll('#productionPoolColorGroupsContainer .pool-group-qty').forEach(input => {
                    const color = input.dataset.color;
                    out[color] = out[color] || [];
                    out[color].push(input.value);
                });
                return out;
            })()
        """)
        print(f"  Raw pool cell values by column: {pool_values}")

        for shade in WHITE_SHADES:
            vals = pool_values.get(shade, [])
            check(len(vals) == 1 and vals[0] == "50", f"'{shade}' column sums both rim colors to 50 (got {vals})")

        for rim in RIM_COLORS:
            vals = pool_values.get(rim, [])
            check(len(vals) == 1 and vals[0] == "150", f"'{rim}' column sums all 6 White shades to 150 (got {vals})")

        if console_errors:
            print("\n⚠️  Console/page errors:")
            for e in console_errors:
                print(f"    {e}")

        print("\n" + ("ALL TESTS PASSED" if not failures else f"{len(failures)} TEST(S) FAILED"))
        browser.close()
        return len(failures) == 0


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
