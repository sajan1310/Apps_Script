"""
Verifies two Production Sheet behaviours:

  1. Narration is resolved LIVE from Items Master (by item name + size) rather
     than read off the snapshot the lot stored in componentsConsumed. Editing
     an item's narration in Items Master must show up on an already-saved lot's
     sheet -- on screen AND in the print/PDF export -- with no re-save of the
     lot. The stored value survives only where Items Master has nothing to say
     (no entry at all, e.g. a Warehouse Pool WIP item; or an entry whose own
     narration is blank).

  2. The item's Base Unit is displayed with the quantity in the on-screen
     dialog: as an input-group addon on the Common table's single Required Qty,
     and as a per-row "Unit" cell on the Per-Color matrix (which has one qty
     column per colour).

Also checks that retyping a row onto a different item re-resolves both fields.

Run: python .pw-test/verify_production_sheet_narration_and_unit.py
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
    "processId": "PRC-1151", "processName": "Packing Crysta 16 inch", "sequence": 6,
    "lotPrefix": "PKG011", "outputItemName": "16 inch Crysta D/Gaddi Steel Rim", "isFinalStage": True,
    "active": True, "processType": "Packing", "primaryColorAxis": ""
}

# Items Master is the source of truth for narration + base unit.
#  - Carton Box:    narration CHANGED in Items Master since the lot was logged.
#  - Adhesive Tape: Items Master narration is BLANK -> the lot's own stored
#                   narration must survive.
#  - Frame Sticker: per-colour item, narration comes from Items Master.
#  - Rework Charge: not in Items Master at all -> stored narration survives,
#                   and no unit can be resolved.
MOCK_ITEMS = [
    {"name": "Carton Box 16 inch", "size": "", "baseUnit": "Pcs", "narration": "Corrugated 5-ply (revised)"},
    {"name": "Adhesive Tape", "size": "", "baseUnit": "Kg", "narration": ""},
    {"name": "Frame Sticker---Blue", "size": "", "baseUnit": "Set", "narration": "Sticker kit A"},
    {"name": "Frame Sticker---Pink", "size": "", "baseUnit": "Set", "narration": "Sticker kit A"},
]

SAVED_LOT = {
    "rowIdx": 73, "lotNumber": "LOT-PKG011-0003", "processId": "PRC-1151", "status": "Completed",
    "qty": 40, "date": "29/07/2026", "color": "Blue, Pink",
    "colorBreakdown": [{"color": "Blue", "qty": 20}, {"color": "Pink", "qty": 20}],
    "componentsConsumed": [
        # Stale snapshot narration -- Items Master must win.
        {"itemName": "Carton Box 16 inch", "size": "", "narration": "OLD 3-ply note",
         "sourceType": "ITEM", "qty": 40, "colorGroup": "COMMON"},
        # Items Master narration is blank -> keep the stored note.
        {"itemName": "Adhesive Tape", "size": "", "narration": "Hand-typed tape note",
         "sourceType": "ITEM", "qty": 2, "colorGroup": "COMMON"},
        # Not in Items Master -> keep the stored note, no unit.
        {"itemName": "Rework Charge", "size": "", "narration": "Ad-hoc labour",
         "sourceType": "ITEM", "qty": 1, "colorGroup": "COMMON"},
        # Per-colour rows: both strip to "Frame Sticker", so they merge into one
        # matrix row -- narration resolved from each literal item name.
        {"itemName": "Frame Sticker---Blue", "size": "", "narration": "stale kit note",
         "sourceType": "ITEM", "qty": 20, "colorGroup": "Blue"},
        {"itemName": "Frame Sticker---Pink", "size": "", "narration": "stale kit note",
         "sourceType": "ITEM", "qty": 20, "colorGroup": "Pink"},
    ],
    "productId": "", "productName": "", "outputItemName": "16 inch Crysta D/Gaddi Steel Rim",
    "assignedTo": "Sanjay", "sheetRemarks": ""
}

ROW_DUMP_JS = """
() => {
  const read = (row) => ({
    name: row.querySelector('.prod-sheet-item-name')?.value ?? null,
    size: row.querySelector('.prod-sheet-size')?.value ?? null,
    narration: row.querySelector('.prod-sheet-narration')?.value ?? null,
    unit: row.querySelector('.prod-sheet-unit')?.textContent.trim() ?? null,
    unitIsAddon: !!row.querySelector('.input-group .prod-sheet-unit'),
    qty: row.querySelector('.prod-sheet-qty')?.value ?? null,
    colorQty: Array.from(row.querySelectorAll('.prod-sheet-color-qty'))
      .map(i => [i.dataset.color, i.value])
  });
  return {
    common: Array.from(document.querySelectorAll('#productionSheetCommonBody tr'))
      .filter(r => r.querySelector('.prod-sheet-item-name')).map(read),
    matrix: Array.from(document.querySelectorAll('#productionSheetMatrixTables .prod-sheet-matrix-tbody tr'))
      .filter(r => r.querySelector('.prod-sheet-item-name')).map(read)
  };
}
"""


def run():
    failures = []

    def check(label, condition, detail=''):
        if condition:
            print(f"  PASS: {label}")
        else:
            print(f"  FAIL: {label}" + (f" -- {detail}" if detail else ''))
            failures.append(label)

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_context().new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        page.evaluate(f"""
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
            App.State.globalItems = {json.dumps(MOCK_ITEMS)};
            App.State.globalProduction = [{json.dumps(SAVED_LOT)}];
        """)

        print("[Step 1] Open the Production Sheet for the saved lot...")
        page.evaluate("App.Production.viewProductionSheet(0)")
        page.locator("#productionSheetModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(400)

        dump = page.evaluate(ROW_DUMP_JS)
        common = {r['name']: r for r in dump['common']}
        matrix = {r['name']: r for r in dump['matrix']}
        print(f"  common rows: {list(common)}")
        print(f"  matrix rows: {list(matrix)}")

        print("\n[Check 1] Narration comes live from Items Master, not the lot's snapshot...")
        carton = common.get('Carton Box 16 inch')
        check("Carton Box narration replaced by the current Items Master value",
              carton and carton['narration'] == 'Corrugated 5-ply (revised)',
              f"got {carton and carton['narration']!r}, wanted 'Corrugated 5-ply (revised)'")

        print("\n[Check 2] Blank Items Master narration falls back to the stored value...")
        tape = common.get('Adhesive Tape')
        check("Adhesive Tape keeps its stored 'Hand-typed tape note'",
              tape and tape['narration'] == 'Hand-typed tape note',
              f"got {tape and tape['narration']!r}")

        print("\n[Check 3] An item Items Master doesn't know keeps its stored narration...")
        rework = common.get('Rework Charge')
        check("Rework Charge keeps 'Ad-hoc labour'",
              rework and rework['narration'] == 'Ad-hoc labour',
              f"got {rework and rework['narration']!r}")

        print("\n[Check 4] Per-colour rows resolve narration from their literal item name...")
        sticker = matrix.get('Frame Sticker')
        check("merged 'Frame Sticker' matrix row exists (both colours merged)",
              sticker is not None, f"matrix rows were {list(matrix)}")
        if sticker:
            check("its narration is the Items Master value 'Sticker kit A', not 'stale kit note'",
                  sticker['narration'] == 'Sticker kit A', f"got {sticker['narration']!r}")
            check("it still carries a qty for both colour columns",
                  sorted(c for c, v in sticker['colorQty'] if v) == ['Blue', 'Pink'],
                  f"got {sticker['colorQty']}")

        print("\n[Check 5] Base Unit is shown with the quantity in the Common table (input-group addon)...")
        check("Carton Box shows 'Pcs' as an addon next to its qty",
              carton and carton['unit'] == 'Pcs' and carton['unitIsAddon'],
              f"unit={carton and carton['unit']!r} isAddon={carton and carton['unitIsAddon']}")
        check("Adhesive Tape shows its own 'Kg' (not a blanket 'Pcs')",
              tape and tape['unit'] == 'Kg', f"got {tape and tape['unit']!r}")
        # A NAMED item Items Master doesn't know keeps the long-standing 'Pcs'
        # fallback (see _resolveDisplayUnit) -- the dialog must agree with the
        # printed sheet, which has always labelled such rows 'Pcs'. Only a row
        # with no item name at all shows the em dash.
        check("Rework Charge (named, unknown to Items Master) falls back to 'Pcs', as the print does",
              rework and rework['unit'] == 'Pcs', f"got {rework and rework['unit']!r}")
        page.evaluate("App.Production.addCommonSheetRow()")
        blank_unit = page.evaluate("""() => {
          const rows = document.querySelectorAll('#productionSheetCommonBody tr');
          return rows[rows.length - 1].querySelector('.prod-sheet-unit')?.textContent.trim();
        }""")
        check("a freshly added blank row shows an em dash rather than guessing a unit",
              blank_unit == '\u2014', f"got {blank_unit!r}")
        page.evaluate("""() => {
          const rows = document.querySelectorAll('#productionSheetCommonBody tr');
          rows[rows.length - 1].remove();
        }""")

        print("\n[Check 6] Base Unit is shown as a per-row cell on the Per-Color matrix...")
        check("'Unit' column header is present on the matrix table",
              page.evaluate("""Array.from(document.querySelectorAll('#productionSheetMatrixTables thead th'))
                               .some(th => th.textContent.trim() === 'Unit')"""))
        check("Frame Sticker matrix row shows 'Set'",
              sticker and sticker['unit'] == 'Set', f"got {sticker and sticker['unit']!r}")

        print("\n[Check 7] Editing a row's Item Name re-resolves its Narration and Unit...")
        page.evaluate("""() => {
          const row = Array.from(document.querySelectorAll('#productionSheetCommonBody tr'))
            .find(r => r.querySelector('.prod-sheet-item-name')?.value === 'Rework Charge');
          const input = row.querySelector('.prod-sheet-item-name');
          input.value = 'Adhesive Tape';
          input.dispatchEvent(new Event('change', { bubbles: true }));
        }""")
        page.wait_for_timeout(200)
        retyped = next(r for r in page.evaluate(ROW_DUMP_JS)['common'] if r['name'] == 'Adhesive Tape'
                       and r['qty'] == '1')
        check("retyped row picked up 'Kg' from the new item",
              retyped['unit'] == 'Kg', f"got {retyped['unit']!r}")
        check("its hand-typed narration survived (Items Master's is blank for Adhesive Tape)",
              retyped['narration'] == 'Ad-hoc labour', f"got {retyped['narration']!r}")

        page.evaluate("""() => {
          const row = Array.from(document.querySelectorAll('#productionSheetCommonBody tr'))
            .find(r => r.querySelector('.prod-sheet-qty')?.value === '1');
          const input = row.querySelector('.prod-sheet-item-name');
          input.value = 'Carton Box 16 inch';
          input.dispatchEvent(new Event('change', { bubbles: true }));
        }""")
        page.wait_for_timeout(200)
        retyped2 = next(r for r in page.evaluate(ROW_DUMP_JS)['common'] if r['qty'] == '1')
        check("retyping onto an item WITH an Items Master narration overwrites the row's note",
              retyped2['narration'] == 'Corrugated 5-ply (revised)', f"got {retyped2['narration']!r}")
        check("and its unit switched to 'Pcs'",
              retyped2['unit'] == 'Pcs', f"got {retyped2['unit']!r}")

        print("\n[Check 8] The print/PDF export carries the live narration through...")
        page.evaluate("App.Production._buildProductionSheetForExport()")
        page.wait_for_timeout(300)
        common_html = page.evaluate("document.getElementById('print-production-sheet-common-tables')?.innerHTML || ''")
        matrix_html = page.evaluate("document.getElementById('print-production-sheet-matrix-tables')?.innerHTML || ''")
        check("exported Common table shows 'Corrugated 5-ply (revised)'",
              'Corrugated 5-ply (revised)' in common_html)
        check("exported Common table no longer shows the stale 'OLD 3-ply note'",
              'OLD 3-ply note' not in common_html)
        check("exported matrix shows 'Sticker kit A', not 'stale kit note'",
              'Sticker kit A' in matrix_html and 'stale kit note' not in matrix_html)
        check("exported matrix still shows the unit with each qty ('20 Set')",
              '20 Set' in matrix_html)

        print("\n[Check 9] Save payload is unaffected by the new Unit cell...")
        payload = page.evaluate("App.Production.serializeProductionSheet()")
        names = sorted({c['itemName'] for c in payload['components']})
        print(f"  serialized items: {names}")
        check("both per-colour matrix quantities still serialize",
              len([c for c in payload['components'] if c['itemName'] == 'Frame Sticker']) == 2,
              f"got {[c for c in payload['components'] if c['itemName'] == 'Frame Sticker']}")
        check("no component picked up a unit string as its narration",
              all(c['narration'] not in ('Pcs', 'Kg', 'Set', '\u2014') for c in payload['components']))

        if console_errors:
            print("\n  Console/page errors:")
            for e in console_errors:
                print(f"    {e}")
            failures.append('console errors')

        browser.close()

    return failures


if __name__ == "__main__":
    fails = run()
    print("\n" + ("ALL PASS" if not fails else f"{len(fails)} FAILED: {fails}"))
    sys.exit(0 if not fails else 1)
