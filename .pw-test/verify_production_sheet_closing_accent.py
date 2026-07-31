"""
Verifies the Production Sheet's "closing accent" fix (2026-07-31): the
container's plain border-bottom was replaced with a real block element
(.print-sheet-closing-accent, View_Print.html), and App.Print's shared
downloadElementAsPDF (Script_Print.html) protects it in html2pdf's
pagebreak.avoid list the same way it already protects table rows -- so a
row that happens to land near a page-height slice boundary can no longer
squeeze the closing bar flush against it with no visible gap.

Mocks window.html2pdf itself (not App.Print.downloadElementAsPDF) so the
real function runs and we can inspect the actual options object it builds.

Run: python .pw-test/verify_production_sheet_closing_accent.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(r"c:\Users\erkar\my-app-script-project\dist\index.html")

PROCESS = {
    "processId": "PRC-1", "processName": "Packing", "sequence": 1, "lotPrefix": "PK",
    "outputItemName": "Item", "isFinalStage": True, "active": True, "processType": "Packing"
}
LOT = {
    "rowIdx": 1, "lotNumber": "LOT-1", "processId": "PRC-1", "status": "Completed", "qty": 20,
    "date": "31/07/2026", "color": "", "colorBreakdown": [],
    "componentsConsumed": [{"itemName": "Item A", "size": "GENERAL", "narration": "", "sourceType": "STOCK", "qty": 5}],
    "productId": "", "productName": "", "outputItemName": "Item", "assignedTo": "", "sheetRemarks": ""
}


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_context().new_page()
        errors = []
        page.on("pageerror", lambda e: errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        page.evaluate("""(d) => {
            App.State.globalColors = [];
            App.State.globalProcesses = [d.process];
            App.State.globalProduction = [d.lot];
            App.State.globalItems = [];
            App.Production._populateProductionSheetData(d.lot, 0);
        }""", {"process": PROCESS, "lot": LOT})

        ok = True

        # ---- DOM structure -------------------------------------------------
        dom = page.evaluate("""() => {
            const container = document.getElementById('print-production-sheet-container');
            const accent = container.querySelector('.print-sheet-closing-accent');
            return {
                lastChildIsAccent: container.lastElementChild === accent,
                accentExists: !!accent,
                containerBorderBottom: container.style.borderBottom || getComputedStyle(container).borderBottomWidth,
            };
        }""")
        print("[DOM]", json.dumps(dom, indent=2))

        if not dom["accentExists"]:
            print("  FAIL: .print-sheet-closing-accent is missing from the container")
            ok = False
        else:
            print("  PASS: .print-sheet-closing-accent is present")

        if not dom["lastChildIsAccent"]:
            print("  FAIL: closing accent is not the container's last child")
            ok = False
        else:
            print("  PASS: closing accent is the last element in the container")

        # ---- html2pdf options actually built by downloadElementAsPDF -------
        page.evaluate("""() => {
            window.__capturedOpts = null;
            window.html2pdf = function() {
                const chain = {
                    set(opts) { window.__capturedOpts = opts; return chain; },
                    from() { return chain; },
                    save() { return Promise.resolve(); }
                };
                return chain;
            };
        }""")

        ok_call = page.evaluate("""async () => {
            return await App.Print.downloadElementAsPDF('print-production-sheet-container', 'test.pdf');
        }""")
        opts = page.evaluate("() => window.__capturedOpts")
        print("\n[html2pdf options]", json.dumps(opts.get("pagebreak") if opts else None, indent=2))

        if not ok_call:
            print("  FAIL: downloadElementAsPDF reported failure")
            ok = False

        avoid = (opts or {}).get("pagebreak", {}).get("avoid", [])
        if "tr" not in avoid:
            print(f"  FAIL: 'tr' missing from pagebreak.avoid: {avoid}")
            ok = False
        else:
            print("  PASS: 'tr' still protected (rows never sliced mid-content)")

        if ".print-sheet-closing-accent" not in avoid:
            print(f"  FAIL: '.print-sheet-closing-accent' missing from pagebreak.avoid: {avoid}")
            ok = False
        else:
            print("  PASS: '.print-sheet-closing-accent' is protected in pagebreak.avoid")

        if errors:
            print("\n  Console/page errors:", errors)
            ok = False

        browser.close()
        return ok


if __name__ == "__main__":
    ok = run()
    print("\n" + ("ALL PASS" if ok else "SOME FAILED"))
    sys.exit(0 if ok else 1)
