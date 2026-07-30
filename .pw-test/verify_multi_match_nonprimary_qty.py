"""
Verification script: a non-primary Color Axis row whose color segment-matches
SEVERAL checked primary colors at once must auto-fill with the SUM of those
matching primaries, not just the first one's quantity.

Reproduces the reported bug exactly as screenshotted on the live app
("Fitted Frame 20 inch Valcano Shocker Tubeless"):

  Primary axis "Painted Frame Valcano 20 inch IBC Shocker"
      Copper-Black         5
      Metallic Green-Black 4
      SeaGreen-Red         5
      Sky Blue-Black       4      -> 18 units of frame, 13 of them "-Black"

  Non-primary axis "Fitted Rim 20 inch"
      BCP                  (unchecked)
      Black                5      <-- WRONG, should be 13

_matchingPrimaryColorQty walked the checked primary rows and returned the
FIRST one that segment-matched ("Copper-Black" -> 5), so the rim quantity
silently described only one of the three black-framed batches. Every
consumer of that number inherited the error: the auto-fill on check, the
live re-sync when a primary qty is edited (onColorQtyChanged), and the rim
axis's own Per-Process Pool Components consumption — 5 rims debited for 13
black-rimmed units.

Run: python .pw-test/verify_multi_match_nonprimary_qty.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

FRAME_COLORS = ["Copper-Black", "Metallic Green-Black", "SeaGreen-Red", "Sky Blue-Black"]
RIM_COLORS = ["BCP", "Black"]

MOCK_PROCESS = {
    "processId": "PRC-1", "processName": "Fitted Frame 20 inch Valcano Shocker Tubeless",
    "sequence": 2, "lotPrefix": "FF", "outputItemName": "Fitted Frame 20 inch",
    "isFinalStage": False, "active": True, "processType": "General"
}

MOCK_COMPONENTS = [
    {"itemName": "Assembly Screws", "size": "", "sourceType": "ITEM", "qtyPerUnit": 4, "colorGroup": "COMMON"},
]
MOCK_ITEMS = [{"name": c["itemName"], "size": c["size"]} for c in MOCK_COMPONENTS]

MOCK_AXES = {
    "axes": [
        {"key": "tag:frame color", "label": "Painted Frame Valcano 20 inch IBC Shocker",
         "colors": FRAME_COLORS, "source": "tag"},
        {"key": "tag:rim color", "label": "Fitted Rim 20 inch", "colors": RIM_COLORS, "source": "tag"},
    ],
    "primaryColorAxis": "Painted Frame Valcano 20 inch IBC Shocker",
    "primaryAxisKey": "tag:frame color",
}

MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": sorted(FRAME_COLORS + RIM_COLORS)},
    "getProcessColorAxes": {"success": True, "data": MOCK_AXES},
    "getProcessComponentsData": {"success": True, "data": MOCK_COMPONENTS},
    "getWarehousePoolData": {"success": True, "data": []},
    "getProcessWipData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
}


def qty_of(page, color):
    return page.evaluate(
        """(c) => {
            const row = document.querySelector(`#productionColorChecklist .production-color-row[data-color="${c}"]`);
            return row ? row.querySelector('.production-color-qty').value : null;
        }""", color)


def checked_of(page, color):
    return page.evaluate(
        """(c) => {
            const row = document.querySelector(`#productionColorChecklist .production-color-row[data-color="${c}"]`);
            return row ? row.querySelector('.production-color-check').checked : null;
        }""", color)


def check_row(page, color):
    page.evaluate(
        """(c) => document.querySelector(
            `#productionColorChecklist .production-color-row[data-color="${c}"] .production-color-check`).click()""",
        color)
    page.wait_for_timeout(250)


def fill_qty(page, color, value):
    page.locator(
        f"#productionColorChecklist .production-color-row[data-color='{color}'] .production-color-qty"
    ).fill(value)
    page.wait_for_timeout(250)


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
            App.State.globalItems = {json.dumps(MOCK_ITEMS)};
            window.__mockResponses = {json.dumps(MOCK_API_RESPONSES)};
            window.google = {{
                script: {{
                    run: {{
                        withSuccessHandler(cb) {{
                            const runner = {{ withFailureHandler() {{ return runner; }} }};
                            Object.keys(window.__mockResponses).forEach(method => {{
                                runner[method] = (...args) => setTimeout(() => cb(window.__mockResponses[method]), 30);
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

        print("[Step 1] Open Create modal, select the 2-axis process...")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.evaluate("""
            document.getElementById('productionProcessId').value = 'PRC-1';
            App.Production.handleProcessChange('PRC-1');
        """)
        page.wait_for_timeout(900)

        rows_info = page.evaluate("""
            Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row')).map(r => ({
                color: r.dataset.color, primary: r.dataset.primary
            }))
        """)
        check(len(rows_info) == 6, f"all 6 axis rows rendered (got {rows_info})")

        print("\n[Step 2] Check the 4 primary frame colors with the screenshot's quantities (5/4/5/4 = 18)...")
        for color, qty in zip(FRAME_COLORS, ["5", "4", "5", "4"]):
            check_row(page, color)
            fill_qty(page, color, qty)

        primary_qtys = {c: qty_of(page, c) for c in FRAME_COLORS}
        check(list(primary_qtys.values()) == ["5", "4", "5", "4"],
              f"primary quantities entered as screenshotted (got {primary_qtys})")

        print("\n[Step 3] The 'Black' rim row auto-checked itself from the -Black frame cascade...")
        check(checked_of(page, "Black") is True,
              "Fitted Rim 'Black' auto-CHECKED itself off the '-Black' frame colors")
        check(checked_of(page, "BCP") is False,
              "Fitted Rim 'BCP' stays unchecked - no frame color matches it")

        print("\n[Step 4] THE BUG: 'Black' must total EVERY matching frame color (5+4+4=13), not just the first (5)...")
        black_qty = qty_of(page, "Black")
        check(black_qty == "13",
              f"Fitted Rim 'Black' = 13 = Copper-Black 5 + Metallic Green-Black 4 + Sky Blue-Black 4, "
              f"not the first match's 5 alone and not the 18 grand total (got '{black_qty}')")

        print("\n[Step 5] Editing ONE matching primary re-totals the rim row live...")
        fill_qty(page, "Sky Blue-Black", "10")
        black_qty2 = qty_of(page, "Black")
        check(black_qty2 == "19",
              f"'Black' follows the new matching total (5+4+10=19) (got '{black_qty2}')")

        print("\n[Step 6] Editing the NON-matching primary (SeaGreen-Red) leaves the rim row alone...")
        fill_qty(page, "SeaGreen-Red", "50")
        black_qty3 = qty_of(page, "Black")
        check(black_qty3 == "19",
              f"'Black' ignores SeaGreen-Red's change - no '-Black' segment in it (got '{black_qty3}')")

        print("\n[Step 7] Unchecking one matching primary drops its share...")
        check_row(page, "Copper-Black")
        black_qty4 = qty_of(page, "Black")
        check(checked_of(page, "Black") is True,
              "'Black' stays checked - two other '-Black' frame colors are still checked")
        check(black_qty4 == "14",
              f"'Black' drops to 4+10=14 once Copper-Black is unchecked (got '{black_qty4}')")

        print("\n[Step 8] A single-match non-primary row is unchanged (regression guard)...")
        # Only Sky Blue-Black remains as a "-Black" match once Metallic
        # Green-Black is unchecked too: the sum collapses back to exactly the
        # old single-match behavior.
        check_row(page, "Metallic Green-Black")
        black_qty5 = qty_of(page, "Black")
        check(black_qty5 == "10",
              f"with one matching primary left, 'Black' equals that row's own qty (10) (got '{black_qty5}')")

        print("\n[Step 9] Manual override still stops auto-tracking...")
        fill_qty(page, "Black", "77")
        fill_qty(page, "Sky Blue-Black", "12")
        black_final = qty_of(page, "Black")
        check(black_final == "77",
              f"operator's manual 77 survives a later primary edit (got '{black_final}')")

        if console_errors:
            print("\nConsole/page errors:")
            for e in console_errors:
                print(f"    {e}")
            failures.append("console errors present")

        print("\n" + ("ALL TESTS PASSED" if not failures else f"{len(failures)} TEST(S) FAILED"))
        browser.close()
        return len(failures) == 0


if __name__ == "__main__":
    sys.exit(0 if run() else 1)
