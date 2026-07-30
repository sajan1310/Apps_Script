"""
Verifies the colour-vs-sub-group split added on 2026-07-30.

A lot's colour list is whatever its recipe tagged components with, and that
legitimately includes packing/variant buckets that are NOT colours — the
reported case was 'KIT BAG 20"' / 'SMALL KIT 20"' sitting alongside Blue-White
and Black. Treating those as colour-matrix columns made the print matrix 10
columns / 901px wide against a 748px page, and html2canvas silently dropped
everything past the edge, including the one column carrying nearly every
quantity on the sheet.

A group counts as a colour only if its whole name decomposes into Color Master
colours (so "Blue-White" and "B/T Green-Orange" do, 'SMALL KIT 20"' does not).
Colours that merely segment-match each other ("Blue" / "Blue-White" /
"Blue-White/Black") stay clubbed in ONE matrix.

Asserted here:
  1. Classification itself, including the composite and non-colour cases.
  2. Colours and sub-groups get separate print tables, and the colour matrix
     no longer runs off the captured box.
  3. Rows only appear in a table where they actually carry a quantity, so the
     colour matrix is not padded out with rows of dashes.
  4. state.colors stays the FULL list. The split is presentation-only —
     serializeProductionSheet reads state.colors, so dropping sub-groups from
     it would silently drop their quantities on save.
  5. The derived split is refreshed when a column is added, not left stale.

Run: python .pw-test/verify_production_sheet_subgroup_split.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(r"c:\Users\erkar\my-app-script-project\dist\index.html")

MASTER = ["BLACK", "RED", "SEAGREEN", "BLUE", "WHITE", "GREY", "YELLOW",
          "ORANGE", "GREEN", "B/T GREEN", "COPPER", "METALLIC GREEN"]

PROCESS = {
    "processId": "PRC-RR024", "processName": "Packing Runway 26 inch",
    "sequence": 6, "lotPrefix": "RR024", "outputItemName": "26 inch Runway IBC T/Tube 2.00 FSDD",
    "isFinalStage": True, "active": True, "processType": "Packing", "primaryColorAxis": ""
}

# One lot mixing real colours with two packing buckets.
LOT = {
    "rowIdx": 91, "lotNumber": "LOT-RR024-0001", "processId": "PRC-RR024",
    "status": "Completed", "qty": 20, "date": "30/07/2026",
    "color": "BLUE-WHITE",
    "colorBreakdown": [
        {"color": "BLUE-WHITE", "qty": 4}, {"color": "BLACK-RED", "qty": 2},
        {"color": "GREY-YELLOW", "qty": 5}, {"color": "B/T GREEN-ORANGE", "qty": 5},
        {"color": "KIT BAG 20\"", "qty": 0}, {"color": "SMALL KIT 20\"", "qty": 0},
    ],
    "componentsConsumed": [
        {"itemName": "Painted Frame Runway 26 inch IBC", "size": "26 inch", "narration": "",
         "sourceType": "POOL", "qty": 4, "colorGroup": "BLUE-WHITE"},
        {"itemName": "Painted Frame Runway 26 inch IBC", "size": "26 inch", "narration": "",
         "sourceType": "POOL", "qty": 2, "colorGroup": "BLACK-RED"},
        {"itemName": "RIM TAPE", "size": "26 inch", "narration": "",
         "sourceType": "STOCK", "qty": 40, "colorGroup": "SMALL KIT 20\""},
        {"itemName": "FREE WHEEL", "size": "26 inch", "narration": "18T",
         "sourceType": "STOCK", "qty": 20, "colorGroup": "SMALL KIT 20\""},
        {"itemName": "HANDLE GRIP BIG", "size": "26 inch", "narration": "",
         "sourceType": "STOCK", "qty": 20, "colorGroup": "KIT BAG 20\""},
        {"itemName": "FORK---BRACE", "size": "GENERAL", "narration": "95-100 MM",
         "sourceType": "STOCK", "qty": 20, "colorGroup": "COMMON"},
    ],
    "productId": "", "productName": "", "outputItemName": "26 inch Runway IBC T/Tube 2.00 FSDD",
    "assignedTo": "", "sheetRemarks": ""
}

CLASSIFY_CASES = [
    ("BLUE-WHITE", True), ("B/T GREEN-ORANGE", True), ("BLACK-RED", True),
    ("Blue", True), ("Blue-White/Black", True), ("METALLIC GREEN", True),
    ("KIT BAG 20\"", False), ("SMALL KIT 20\"", False),
    ("COMMON", False), ("", False),
    # a name that merely STARTS with a colour is not a colour group
    ("COPPER KIT", False),
]


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
        }""", {"master": MASTER, "process": PROCESS, "lot": LOT})

        ok = True

        # ---- 1. classification -------------------------------------------
        got = page.evaluate("(cases) => cases.map(c => [c[0], App.Production._isColorGroupName(c[0])])",
                            CLASSIFY_CASES)
        print("[classification]")
        for (name, expected), (_, actual) in zip(CLASSIFY_CASES, got):
            mark = "PASS" if actual == expected else "FAIL"
            if actual != expected:
                ok = False
            print(f"  {mark}: {name!r:24} -> isColor={actual} (expected {expected})")

        # ---- 2/4. dialog render + stored split ----------------------------
        page.evaluate("App.Production._populateProductionSheetData(App.State.globalProduction[0], 0)")
        state = page.evaluate("""() => {
            const s = App.State.currentProductionSheet;
            return { colors: s.colors, colorGroups: s.colorGroups, subGroups: s.subGroups };
        }""")
        print("\n[stored split]", json.dumps(state, indent=2))

        if sorted(state["colorGroups"]) != sorted(["BLUE-WHITE", "BLACK-RED", "GREY-YELLOW", "B/T GREEN-ORANGE"]):
            print(f"  FAIL: unexpected colorGroups: {state['colorGroups']}")
            ok = False
        else:
            print("  PASS: colorGroups holds only the real colours")

        if sorted(state["subGroups"]) != sorted(['KIT BAG 20"', 'SMALL KIT 20"']):
            print(f"  FAIL: unexpected subGroups: {state['subGroups']}")
            ok = False
        else:
            print("  PASS: subGroups holds only the packing buckets")

        # state.colors must remain the FULL list -- the save path reads it.
        if sorted(state["colors"]) != sorted(state["colorGroups"] + state["subGroups"]):
            print(f"  FAIL: state.colors is not the union of both groups: {state['colors']}")
            print("        serializeProductionSheet reads state.colors; a short list drops quantities on save.")
            ok = False
        else:
            print("  PASS: state.colors is still the full list (save path cannot lose sub-group qty)")

        # ---- save round-trip: sub-group quantities survive ----------------
        ser = page.evaluate("() => App.Production.serializeProductionSheet()")
        comps = ser["components"] if isinstance(ser, dict) and "components" in ser else ser
        groups_saved = sorted({str(c.get("colorGroup", "")) for c in comps})
        print(f"\n[serialize] colorGroups present after save: {groups_saved}")
        for needed in ['SMALL KIT 20"', 'KIT BAG 20"']:
            if needed not in groups_saved:
                print(f"  FAIL: {needed!r} was dropped by serializeProductionSheet")
                ok = False
        if all(n in groups_saved for n in ['SMALL KIT 20"', 'KIT BAG 20"']):
            print("  PASS: sub-group quantities survive a save round-trip")

        # ---- 2/3. print tables -------------------------------------------
        page.evaluate("App.Production._buildProductionSheetForExport()")
        geo = page.evaluate("""() => {
            const el = document.getElementById('print-production-sheet-container');
            const prev = el.getAttribute('style');
            el.style.position='absolute'; el.style.left='0'; el.style.top='0';
            el.style.visibility='hidden'; el.style.display='block';
            el.style.width = App.Print.PAGE_WIDTH_PX + 'px';
            const mt = el.querySelector('#print-production-sheet-matrix-tables table');
            const st = el.querySelector('#print-production-sheet-subgroup-tables table');
            const heads = t => t ? [...t.querySelectorAll('thead th')].map(th => th.textContent.trim()) : null;
            const r = {
              matrixHeaders: heads(mt), subHeaders: heads(st),
              matrixRowCount: mt ? mt.querySelectorAll('tbody tr').length : 0,
              subRowCount: st ? st.querySelectorAll('tbody tr').length : 0,
              subSectionShown: getComputedStyle(document.getElementById('print-prod-subgroup-section')).display !== 'none',
              overflowPx: Math.round(el.scrollWidth - el.clientWidth),
            };
            el.setAttribute('style', prev);
            return r;
        }""")
        print("\n[print tables]", json.dumps(geo, indent=2))

        if not geo["subSectionShown"]:
            print("  FAIL: Sub-Group Components section stayed hidden")
            ok = False
        else:
            print("  PASS: Sub-Group Components section is rendered")

        for bad in ['KIT BAG 20"', 'SMALL KIT 20"']:
            if bad in (geo["matrixHeaders"] or []):
                print(f"  FAIL: {bad!r} is still a colour-matrix column")
                ok = False
        if not any(b in (geo["matrixHeaders"] or []) for b in ['KIT BAG 20"', 'SMALL KIT 20"']):
            print("  PASS: no packing bucket appears as a colour column")

        if sorted(h for h in (geo["subHeaders"] or []) if 'KIT' in h) != sorted(['KIT BAG 20"', 'SMALL KIT 20"']):
            print(f"  FAIL: sub-group table headers wrong: {geo['subHeaders']}")
            ok = False
        else:
            print("  PASS: both packing buckets are columns of the sub-group table")

        # The colour matrix must not be padded with all-dash rows: only the
        # Painted Frame carries colour quantities here.
        if geo["matrixRowCount"] != 1:
            print(f"  FAIL: colour matrix has {geo['matrixRowCount']} rows, expected 1 (only Painted Frame has colour qty)")
            ok = False
        else:
            print("  PASS: colour matrix carries only rows that have a colour quantity")

        if geo["overflowPx"] > 2:
            print(f"  FAIL: sheet still overflows the captured box by {geo['overflowPx']}px")
            ok = False
        else:
            print(f"  PASS: nothing runs past the captured box ({geo['overflowPx']}px)")

        # ---- 5. split is refreshed, not stale -----------------------------
        refreshed = page.evaluate("""() => {
            const s = App.State.currentProductionSheet;
            s.colors.push('PURPLE');
            App.Production._refreshSheetGroupSplit(s);
            return { colorGroups: s.colorGroups, subGroups: s.subGroups };
        }""")
        # PURPLE is absent from this test's Color Master, so it must land in subGroups.
        if 'PURPLE' not in refreshed["subGroups"]:
            print(f"\n  FAIL: _refreshSheetGroupSplit did not pick up the new column: {refreshed}")
            ok = False
        else:
            print("\n  PASS: _refreshSheetGroupSplit re-derives the split after a column is added")

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
