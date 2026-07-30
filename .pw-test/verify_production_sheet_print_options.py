"""
Verifies the "Print options" panel on the Production Sheet dialog (2026-07-30):
per-column include/exclude and portrait/landscape, feeding both Print Sheet and
Download PDF.

Key invariant: excluding a column is a PRINT-ONLY choice. state.colors (which
serializeProductionSheet reads) must be untouched, so unticking a column can
never delete a quantity from the lot.

Run: python .pw-test/verify_production_sheet_print_options.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(r"c:\Users\erkar\my-app-script-project\dist\index.html")

MASTER = ["BLACK", "RED", "BLUE", "WHITE", "GREY", "YELLOW", "ORANGE", "GREEN", "B/T GREEN"]

PROCESS = {
    "processId": "PRC-RR024", "processName": "Packing Runway 26 inch", "sequence": 6,
    "lotPrefix": "RR024", "outputItemName": "26 inch Runway IBC", "isFinalStage": True,
    "active": True, "processType": "Packing", "primaryColorAxis": ""
}

LOT = {
    "rowIdx": 91, "lotNumber": "LOT-RR024-0001", "processId": "PRC-RR024",
    "status": "Completed", "qty": 20, "date": "30/07/2026", "color": "BLUE-WHITE",
    "colorBreakdown": [
        {"color": "BLUE-WHITE", "qty": 4}, {"color": "BLACK-RED", "qty": 2},
        {"color": "GREY-YELLOW", "qty": 5}, {"color": "SMALL KIT 20\"", "qty": 0},
    ],
    "componentsConsumed": [
        {"itemName": "Painted Frame", "size": "26 inch", "narration": "", "sourceType": "POOL", "qty": 4, "colorGroup": "BLUE-WHITE"},
        {"itemName": "Painted Frame", "size": "26 inch", "narration": "", "sourceType": "POOL", "qty": 2, "colorGroup": "BLACK-RED"},
        {"itemName": "Painted Frame", "size": "26 inch", "narration": "", "sourceType": "POOL", "qty": 5, "colorGroup": "GREY-YELLOW"},
        {"itemName": "RIM TAPE", "size": "26 inch", "narration": "", "sourceType": "STOCK", "qty": 40, "colorGroup": "SMALL KIT 20\""},
    ],
    "productId": "", "productName": "", "outputItemName": "26 inch Runway IBC",
    "assignedTo": "", "sheetRemarks": ""
}

HEADERS_JS = """() => {
  const el = document.getElementById('print-production-sheet-container');
  const mt = el.querySelector('#print-production-sheet-matrix-tables table');
  const st = el.querySelector('#print-production-sheet-subgroup-tables table');
  const heads = t => t ? [...t.querySelectorAll('thead th')].map(th => th.textContent.trim()) : [];
  return { matrix: heads(mt), sub: heads(st) };
}"""


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context(viewport={"width": 1400, "height": 900})
        page = ctx.new_page()
        errors = []
        page.on("pageerror", lambda e: errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1200)
        page.evaluate("""(d) => {
            App.State.globalColors = d.master.map(n => ({ name: n }));
            App.State.globalProcesses = [d.process];
            App.State.globalProduction = [d.lot];
            App.State.globalItems = [];
            App.Production._populateProductionSheetData(d.lot, 0);
        }""", {"master": MASTER, "process": PROCESS, "lot": LOT})

        ok = True

        # ---- the panel is built from the lot's own columns -----------------
        panel = page.evaluate("""() => $$('#productionSheetPrintColumns input[type=checkbox]')
            .map(b => ({ group: b.dataset.group, checked: b.checked }))""")
        print("[panel]", json.dumps(panel, indent=2))
        if not panel:
            print("  FAIL: print-options panel rendered no column checkboxes")
            ok = False
        elif not all(p["checked"] for p in panel):
            print("  FAIL: columns should default to all-included")
            ok = False
        else:
            print(f"  PASS: panel lists all {len(panel)} columns, all ticked by default")

        # ---- default export includes everything ---------------------------
        page.evaluate("App.Production._buildProductionSheetForExport()")
        base = page.evaluate(HEADERS_JS)
        print("\n[default export]", json.dumps(base, indent=2))
        if "GREY-YELLOW" not in base["matrix"]:
            print("  FAIL: a colour column is missing from the default export")
            ok = False
        else:
            print("  PASS: default export includes every colour column")

        # ---- untick one colour -> gone from print, kept in data ------------
        page.evaluate("""() => {
            const b = $$('#productionSheetPrintColumns input[type=checkbox]')
              .find(x => x.dataset.group === 'GREY-YELLOW');
            b.checked = false;
        }""")
        page.evaluate("App.Production._buildProductionSheetForExport()")
        after = page.evaluate(HEADERS_JS)
        print("\n[after unticking GREY-YELLOW]", json.dumps(after, indent=2))
        if "GREY-YELLOW" in after["matrix"]:
            print("  FAIL: unticked colour still printed")
            ok = False
        else:
            print("  PASS: unticked colour is left out of the printed sheet")

        still = page.evaluate("() => App.State.currentProductionSheet.colors")
        if "GREY-YELLOW" not in still:
            print(f"  FAIL: unticking MUTATED the lot's data: {still}")
            print("        serializeProductionSheet reads state.colors; this would delete a real quantity.")
            ok = False
        else:
            print("  PASS: the lot's own colour list is untouched (print-only choice)")

        ser = page.evaluate("() => App.Production.serializeProductionSheet()")
        comps = ser["components"] if isinstance(ser, dict) and "components" in ser else ser
        if not any(str(c.get("colorGroup")) == "GREY-YELLOW" for c in comps):
            print("  FAIL: the excluded colour's quantity was dropped on save")
            ok = False
        else:
            print("  PASS: the excluded colour's quantity still saves")

        # ---- untick a sub-group ------------------------------------------
        page.evaluate("""() => {
            $$('#productionSheetPrintColumns input[type=checkbox]').forEach(b => { b.checked = true; });
            const b = $$('#productionSheetPrintColumns input[type=checkbox]')
              .find(x => x.dataset.group === 'SMALL KIT 20"');
            if (b) b.checked = false;
        }""")
        page.evaluate("App.Production._buildProductionSheetForExport()")
        nosub = page.evaluate(HEADERS_JS)
        shown = page.evaluate("() => getComputedStyle(document.getElementById('print-prod-subgroup-section')).display !== 'none'")
        print(f"\n[after unticking the sub-group] sub headers={nosub['sub']} sectionShown={shown}")
        if shown:
            print("  FAIL: sub-group section still shown after its only column was excluded")
            ok = False
        else:
            print("  PASS: excluding the only sub-group hides that whole section")

        # ---- landscape changes the geometry the fit loop measures against --
        page.evaluate("""() => {
            $$('#productionSheetPrintColumns input[type=checkbox]').forEach(b => { b.checked = true; });
            document.getElementById('prodSheetOrientLandscape').checked = true;
        }""")
        opts = page.evaluate("() => App.Production._printOptions()")
        over = page.evaluate("() => App.Production._pdfOverridesForOptions()")
        print("\n[landscape]", json.dumps({"opts": opts, "overrides": over}, indent=2))
        if not opts["landscape"]:
            print("  FAIL: landscape not detected from the panel")
            ok = False
        elif over.get("jsPDF", {}).get("orientation") != "landscape":
            print("  FAIL: jsPDF overrides do not request landscape")
            ok = False
        elif over.get("captureWidthPx") != page.evaluate("() => App.Print.PAGE_HEIGHT_PX"):
            print("  FAIL: capture width was not switched to the rotated page width")
            ok = False
        else:
            print("  PASS: landscape sets both jsPDF orientation and the rotated capture width")

        if errors:
            print("\n  Console/page errors:")
            for e in errors:
                print(f"    {e}")
            ok = False

        browser.close()
        return ok


if __name__ == "__main__":
    ok = run()
    print("\n" + ("ALL PASS" if ok else "SOME FAILED"))
    sys.exit(0 if ok else 1)
