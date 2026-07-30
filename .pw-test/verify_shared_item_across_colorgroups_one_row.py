"""
Verification script: a single physical item that is genuinely shared between
two colors (e.g. one sticker sheet printed with both colors, consumed by
both a "Red-White" frame lot and a "SeaGreen" frame lot) must render as ONE
row in the Log Production Lot's Per-Color matrix, with one quantity column
per color — not two separate rows.

Reported bug: the recipe tags this shared item under two different
colorGroups ("Red-White" and "SeaGreen"), each pointing at the identical
itemName. The existing merge logic (_stripColorSubstring/_matchedColorToken
in Script_Production.html) strips THIS color's token out of the item name
before using it as the row key — built for the "same generic component,
genuinely different item per color" case (e.g. "Frame Sticker---Red-White" /
"Frame Sticker---SeaGreen"). That stripping actively corrupts a shared item's
key: the "SeaGreen" token is found and stripped INSIDE the item's own name
("...Pink/SeaGreen" -> "...Pink"), while the "Red-White" token is never found
in it at all (the item is named after "Pink", not "Red") -- so that pass
keeps the full raw name unchanged. Two different leftover strings -> two
rows for what is really one item.

Fix: _sharedItemSlotKeys detects an itemName+size referenced under 2+ distinct
colorGroups and skips stripping for those -- they're matched on their raw,
already-identical name instead. Covered at all three call sites that do this
stripping: populateColorMatrixForColors (live Create/Edit matrix),
populateComponentsConsumedDirect (Edit Lot restore), and groupComponentsForSheet
(printed sheet) -- this script covers the first, the live path the bug was
originally reported against.

A second, genuinely-different-item-per-color pair ("Frame Sticker---Red-White"
/ "Frame Sticker---SeaGreen") is included as a regression check: that case
MUST still collapse via stripping into one "Frame Sticker" row, same as before
this fix.

Run: python .pw-test/verify_shared_item_across_colorgroups_one_row.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

SHARED_ITEM = "Curvy Sticker Backrest Sticker Pink/SeaGreen"

MOCK_PROCESS = {
    "processId": "PRC-1", "processName": "Packing", "sequence": 3,
    "lotPrefix": "PACK", "outputItemName": "Packed Cycle", "isFinalStage": True,
    "active": True, "processType": "General"
}

MOCK_COMPONENTS = [
    {"itemName": "Carton Box", "size": "", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "COMMON"},
    # The shared item -- genuinely ONE physical component, tagged under both colors.
    {"itemName": SHARED_ITEM, "size": "GENERAL", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "Red-White"},
    {"itemName": SHARED_ITEM, "size": "GENERAL", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "SeaGreen"},
    # Regression control -- genuinely different literal item per color, same generic slot.
    {"itemName": "Frame Sticker---Red-White", "size": "16 inch", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "Red-White"},
    {"itemName": "Frame Sticker---SeaGreen", "size": "16 inch", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "SeaGreen"},
]
MOCK_ITEMS = [{"name": SHARED_ITEM, "size": "GENERAL"}, {"name": "Carton Box", "size": ""},
              {"name": "Frame Sticker---Red-White", "size": "16 inch"}, {"name": "Frame Sticker---SeaGreen", "size": "16 inch"}]

MOCK_AXES = {
    "axes": [
        {"key": "tag:frame color", "label": "Frame Color", "colors": ["Red-White", "SeaGreen"], "source": "tag"},
    ],
    "primaryColorAxis": "Frame Color",
    "primaryAxisKey": "tag:frame color",
}

MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": ["Red-White", "SeaGreen"]},
    "getProcessColorAxes": {"success": True, "data": MOCK_AXES},
    "getProcessComponentsData": {"success": True, "data": MOCK_COMPONENTS},
    "getWarehousePoolData": {"success": True, "data": []},
    "getProcessWipData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
}


def matrix_rows(page):
    return page.evaluate("""
        Array.from(document.querySelectorAll('#productionColorMatrixBody tr[data-merged="true"]')).map(row => {
            const headerCells = Array.from(document.querySelectorAll('#productionColorMatrixHeaderRow th[data-color]'));
            const cells = {};
            headerCells.forEach((th, i) => {
                const cell = row.children[i + 4]; // Name/Size/Narration/Source precede color columns
                const qty = cell ? cell.querySelector('.matrix-qty')?.value || '' : '';
                cells[th.dataset.color] = qty;
            });
            return {
                displayName: row.querySelector('.prod-comp-display-name')?.value || '',
                size: row.querySelector('.prod-comp-size')?.value || '',
                cells
            };
        })
    """)


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
            window.google = {{ script: {{ run: {{
                withSuccessHandler(cb) {{
                    const runner = {{ withFailureHandler() {{ return runner; }} }};
                    Object.keys(window.__mockResponses).forEach(method => {{
                        runner[method] = () => setTimeout(() => cb(window.__mockResponses[method]), 30);
                    }});
                    return runner;
                }}
            }} }} }};
        """)

        failures = []

        def check(cond, msg):
            print(("PASS: " if cond else "FAIL: ") + msg)
            if not cond:
                failures.append(msg)

        print("[Step 1] Open Log Production Lot on the Packing process...")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.evaluate(f"""
            document.getElementById('productionProcessId').value = 'PRC-1';
            App.Production.handleProcessChange('PRC-1');
        """)
        page.wait_for_timeout(900)

        print("\n[Step 2] Check both Red-White and SeaGreen at 10 and 8 respectively...")
        page.locator('#productionColorChecklist .production-color-row[data-color="Red-White"] .production-color-check').click()
        page.wait_for_timeout(200)
        page.locator('#productionColorChecklist .production-color-row[data-color="Red-White"] .production-color-qty').fill("10")
        page.wait_for_timeout(200)
        page.locator('#productionColorChecklist .production-color-row[data-color="SeaGreen"] .production-color-check').click()
        page.wait_for_timeout(200)
        page.locator('#productionColorChecklist .production-color-row[data-color="SeaGreen"] .production-color-qty').fill("8")
        page.wait_for_timeout(400)

        rows = matrix_rows(page)
        print("  matrix rows:", json.dumps(rows, indent=None))

        shared_rows = [r for r in rows if SHARED_ITEM in r["displayName"]]
        check(len(shared_rows) == 1,
              f"Exactly ONE row exists for the shared item (got {len(shared_rows)})")
        if shared_rows:
            check(shared_rows[0]["displayName"] == SHARED_ITEM,
                  f"Row keeps the full, unmangled item name (got '{shared_rows[0]['displayName']}')")
            check(shared_rows[0]["cells"].get("Red-White") == "10",
                  f"Red-White column shows 10 (got {shared_rows[0]['cells'].get('Red-White')})")
            check(shared_rows[0]["cells"].get("SeaGreen") == "8",
                  f"SeaGreen column shows 8 (got {shared_rows[0]['cells'].get('SeaGreen')})")

        print("\n[Regression check] Frame Sticker (genuinely different item per color) still merges via stripping...")
        sticker_rows = [r for r in rows if "Frame Sticker" in r["displayName"]]
        check(len(sticker_rows) == 1,
              f"Exactly ONE row exists for Frame Sticker, color-stripped (got {len(sticker_rows)}: {[r['displayName'] for r in sticker_rows]})")
        if sticker_rows:
            check(sticker_rows[0]["cells"].get("Red-White") == "10",
                  f"Frame Sticker Red-White column shows 10 (got {sticker_rows[0]['cells'].get('Red-White')})")
            check(sticker_rows[0]["cells"].get("SeaGreen") == "8",
                  f"Frame Sticker SeaGreen column shows 8 (got {sticker_rows[0]['cells'].get('SeaGreen')})")

        check(len(console_errors) == 0, f"No uncaught page errors (got {console_errors})")

        print("\n" + "=" * 60)
        if failures:
            print(f"FAILURES: {len(failures)}")
            for f in failures:
                print("  -", f)
        else:
            print("ALL CHECKS PASSED")
        print("=" * 60)

        browser.close()
        return len(failures) == 0


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
