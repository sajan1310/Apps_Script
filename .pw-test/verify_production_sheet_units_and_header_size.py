"""
Verifies the Production Sheet print/PDF export now shows the unit alongside
each Qty (e.g. "10 Pcs", "5 Kg" for an Items Master item; "10 Pcs" fallback
for a Warehouse Pool WIP item with no Items Master entry) in both the
Common table and the Per-Color matrix table, and that header cells are now
sized >= body font (not smaller).

Run: python .pw-test/verify_production_sheet_units_and_header_size.py
"""
import sys
import io
import json
import re
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(r"c:\Users\erkar\my-app-script-project\dist\index.html")
TIMEOUT = 8000

MOCK_PROCESS = {
    "processId": "PRC-1151", "processName": "Packing Crysta 16 inch D/Gaddi Steel Rim", "sequence": 6,
    "lotPrefix": "PKG011", "outputItemName": "16 inch Crysta D/Gaddi Steel Rim", "isFinalStage": True,
    "active": True, "processType": "Packing", "primaryColorAxis": ""
}

MOCK_ITEMS = [
    {"name": "Carton Box 16 inch", "size": "", "baseUnit": "Pcs"},
    {"name": "Adhesive Tape", "size": "", "baseUnit": "Kg"},
]

SAVED_LOT = {
    "rowIdx": 73, "lotNumber": "LOT-PKG011-0003", "processId": "PRC-1151", "status": "Completed",
    "qty": 40, "date": "29/07/2026", "color": "Blue-White, Pink-White",
    "colorBreakdown": [{"color": "Blue-White", "qty": 20}, {"color": "Pink-White", "qty": 20}],
    "componentsConsumed": [
        {"itemName": "Fitted Frame 16 inch Crysta S/Rim", "size": "", "narration": "", "sourceType": "POOL", "qty": 20, "colorGroup": "Blue-White"},
        {"itemName": "Fitted Frame 16 inch Crysta S/Rim", "size": "", "narration": "", "sourceType": "POOL", "qty": 20, "colorGroup": "Pink-White"},
        {"itemName": "Carton Box 16 inch", "size": "", "narration": "Packing material", "sourceType": "ITEM", "qty": 40, "colorGroup": "COMMON"},
        {"itemName": "Adhesive Tape", "size": "", "narration": "", "sourceType": "ITEM", "qty": 2, "colorGroup": "COMMON"},
    ],
    "productId": "", "productName": "", "outputItemName": "16 inch Crysta D/Gaddi Steel Rim",
    "assignedTo": "Sanjay", "sheetRemarks": ""
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
            App.State.globalItems = {json.dumps(MOCK_ITEMS)};
            App.State.globalProduction = [{json.dumps(SAVED_LOT)}];
        """)

        print("[Step 1] Open the Production Sheet popup and build the export...")
        page.evaluate("App.Production.viewProductionSheet(0)")
        page.locator("#productionSheetModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(500)

        page.evaluate("App.Production._buildProductionSheetForExport()")
        page.wait_for_timeout(300)

        common_html = page.evaluate("document.getElementById('print-production-sheet-common-tables')?.innerHTML || ''")
        matrix_html = page.evaluate("document.getElementById('print-production-sheet-matrix-tables')?.innerHTML || ''")

        ok = True

        print("\n[Check] Common table shows units next to qty...")
        if "40 Pcs" not in common_html:
            print("  FAIL: expected 'Carton Box 16 inch' row to show '40 Pcs' (Items Master baseUnit=Pcs)")
            ok = False
        else:
            print("  PASS: 'Carton Box 16 inch' shows '40 Pcs'")

        if "2 Kg" not in common_html:
            print("  FAIL: expected 'Adhesive Tape' row to show '2 Kg' (Items Master baseUnit=Kg)")
            ok = False
        else:
            print("  PASS: 'Adhesive Tape' shows '2 Kg'")

        print("\n[Check] Matrix table shows units next to each per-color qty (Pool item, no Items Master entry -> 'Pcs' fallback)...")
        if "20 Pcs" not in matrix_html:
            print("  FAIL: expected 'Fitted Frame...' matrix cells to show '20 Pcs' (fallback unit)")
            ok = False
        else:
            print("  PASS: matrix cells show '20 Pcs'")

        # Both color columns (Blue-White, Pink-White) should each carry the
        # unit -- count occurrences of "20 Pcs" in the matrix row.
        occurrences = len(re.findall(r'20 Pcs', matrix_html))
        print(f"  '20 Pcs' occurrence count in matrix: {occurrences}")
        if occurrences != 2:
            print(f"  FAIL: expected exactly 2 occurrences (one per color column), got {occurrences}")
            ok = False
        else:
            print("  PASS: unit shown once per color column, as expected")

        print("\n[Check] Header font-size is now >= the PLAIN (non-emphasis) body font-size...")
        head_font_match = re.search(r'<th[^>]*font-size:([\d.]+)px', common_html)
        # td cells in order: Item Name (emphasis), Size/Narration (plain),
        # ..., Qty (emphasis) -- the 2nd <td> is a plain-tier one.
        body_font_matches = re.findall(r'<td[^>]*font-size:([\d.]+)px', common_html)
        if head_font_match and len(body_font_matches) >= 2:
            head_font = float(head_font_match.group(1))
            plain_body_font = float(body_font_matches[1])
            print(f"  header font={head_font}px, plain body cell font={plain_body_font}px")
            if head_font < plain_body_font:
                print("  FAIL: header font is still smaller than the plain body font")
                ok = False
            else:
                print("  PASS: header font >= plain body font")
        else:
            print("  FAIL: could not find header/body font-size in rendered HTML")
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
