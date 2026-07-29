"""
Verifies App.Production.bulkDownloadPDF() now downloads ONE SEPARATE PDF
per selected lot (each built from the same rich single-lot Production
Sheet template), instead of one combined multi-page PDF or the flatter
buildProductionSheetPrintPageHtml row list. Mocks App.Print.downloadElementAsPDF
so we can inspect what container/filename it was called with per lot,
without needing a live html2pdf/network fetch.

Run: python .pw-test/verify_production_bulk_download_separate_pdfs.py
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
    "processId": "PRC-1151", "processName": "Packing Crysta 16 inch D/Gaddi Steel Rim", "sequence": 6,
    "lotPrefix": "PKG011", "outputItemName": "16 inch Crysta D/Gaddi Steel Rim", "isFinalStage": True,
    "active": True, "processType": "Packing", "primaryColorAxis": ""
}

def make_lot(row_idx, lot_number, date, qty, color, colorGroup):
    return {
        "rowIdx": row_idx, "lotNumber": lot_number, "processId": "PRC-1151", "status": "Completed",
        "qty": qty, "date": date, "color": color,
        "colorBreakdown": [{"color": color, "qty": qty}],
        "componentsConsumed": [
            {"itemName": "Fitted Frame 16 inch Crysta S/Rim", "size": "", "narration": "", "sourceType": "POOL", "qty": qty, "colorGroup": colorGroup},
        ],
        "productId": "", "productName": "", "outputItemName": "16 inch Crysta D/Gaddi Steel Rim",
        "assignedTo": "Sanjay", "sheetRemarks": ""
    }

LOTS = [
    make_lot(73, "LOT-PKG011-0003", "29/07/2026", 10, "Blue-White", "Blue-White"),
    make_lot(74, "LOT-PKG011-0004", "29/07/2026", 20, "Black", "Black"),
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
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
            App.State.globalProduction = {json.dumps(LOTS)};
            App.State.selectedProduction = ['73', '74'];

            window.__pdfCalls = [];
            App.Print.downloadElementAsPDF = async (elementId, filename) => {{
                const el = document.getElementById(elementId);
                window.__pdfCalls.push({{
                    elementId, filename,
                    qtyText: el ? el.querySelector('#print-prod-qty')?.innerText : null,
                    matrixHtml: el ? (el.querySelector('#print-production-sheet-matrix-tables')?.innerHTML || '') : null,
                }});
                return true;
            }};
            void 0;
        """)

        print("[Step 1] Call App.Production.bulkDownloadPDF() with 2 selected lots...")
        page.evaluate("App.Production.bulkDownloadPDF()")
        page.wait_for_timeout(1000)

        calls = page.evaluate("window.__pdfCalls")
        print(f"  downloadElementAsPDF calls: {json.dumps(calls, indent=2)}")

        ok = True
        if len(calls) != 2:
            print(f"  FAIL: expected 2 separate downloadElementAsPDF calls (one per lot), got {len(calls)}")
            ok = False
        else:
            print("  PASS: exactly 2 separate PDF downloads triggered (not 1 combined file)")

        if calls and any(c['elementId'] != 'print-production-sheet-container' for c in calls):
            print("  FAIL: expected every call to target 'print-production-sheet-container' (the rich single-lot template), not a bulk-page container")
            ok = False
        else:
            print("  PASS: every call uses the rich single-lot Production Sheet container")

        filenames = [c['filename'] for c in calls]
        if len(set(filenames)) != len(filenames):
            print(f"  FAIL: filenames are not unique: {filenames}")
            ok = False
        else:
            print(f"  PASS: filenames are distinct: {filenames}")

        if len(calls) == 2:
            qty_texts = [c['qtyText'] for c in calls]
            print(f"  Lot Qty shown at each download: {qty_texts}")
            if qty_texts != ['10', '20']:
                print(f"  FAIL: expected qty '10' then '20' (matching each lot), got {qty_texts}")
                ok = False
            else:
                print("  PASS: each download's Lot Qty matches its own lot, no stale data from the previous iteration")

            has_blue_white_0 = 'Blue-White' in (calls[0]['matrixHtml'] or '')
            has_black_1 = 'Black' in (calls[1]['matrixHtml'] or '')
            has_black_0 = 'Black' in (calls[0]['matrixHtml'] or '')
            has_blue_white_1 = 'Blue-White' in (calls[1]['matrixHtml'] or '')
            print(f"  Call 0 matrix mentions Blue-White={has_blue_white_0} Black={has_black_0}; Call 1 mentions Blue-White={has_blue_white_1} Black={has_black_1}")
            if not (has_blue_white_0 and not has_black_0 and has_black_1 and not has_blue_white_1):
                print("  FAIL: matrix content leaked between lots or didn't match the expected per-lot color")
                ok = False
            else:
                print("  PASS: matrix content is correctly scoped to each lot's own color, no cross-lot leakage")

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
