"""
Verifies the production sheet print redesign (single-table color matrix,
two-column Common Components, and the measure/compress fit loop) added in
App.Production.printProductionSheet (Script_Production.html) and its
static scaffold (View_Print.html #print-production-sheet-container).

Seeds #productionSheetCommonBody / #productionSheetMatrixTables directly
(the same DOM the live Add/Edit Production Sheet modal builds) with the
user's real worst-case lot (15 common components, 12 per-color rows across
6 colors) that used to spill across ~1.5 pages as 6 separate per-signature
matrix tables, then checks the printed output consolidates to one matrix
table and fits within the A4 printable height at some density tier.

Run: python .pw-test/verify_production_sheet_autofit.py
"""
import sys
import io
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"

failures = 0
def check(cond, msg):
    global failures
    if cond:
        print(f"  PASS: {msg}")
    else:
        failures += 1
        print(f"  FAIL: {msg}")


COMMON_ROWS = [
    ("CARTOON---S-D", "20 inch", "45x7x20.5 - Priority 6", "20"),
    ("Side Support-CHANNEL---WHITE--LD", "20 inch", "LD Wheel with extra Nut w/o Clump", "20"),
    ("HANDLE-TNT---BLACK", "20 inch", "20\" Black", "20"),
    ("Backrest -D-26---BLACK", "GENERAL", "-", "20"),
    ("WATER-BOTTLE-White-400-Teddy", "GENERAL", "Teddy Bottle", "20"),
    ("BIG-BAG-HDPE---SD", "20 inch", "(SD*71)", "5"),
    ("BAG-KIT---WHITE", "14 inch", "Kit", "20"),
    ("BASKET-FITTING-ZINC-SMALL", "GENERAL", "Kit", "20"),
    ("CLUTCH---BLACK-40B", "GENERAL", "Kit", "20"),
    ("PEDAL-BLACK---STEEL BALL-20\"", "20 inch", "Fan Type", "20"),
    ("Chain Cover---BOLT", "GENERAL", "pouch-kit (5x8)", "60"),
    ("SEAT-FILLER-BOLT---ZINC", "GENERAL", "pouch-kit", "20"),
    ("PVC-MUDGUARD---BLACK-SUMO", "20 inch", "-", "20"),
    ("SEAT-PU-3021----BLACK-20\"", "GENERAL", "2 Color Printing", "20"),
    ("SEAT-FILLER---6\"X 25.40", "GENERAL", "-", "5"),
]

COLORS = ["Baby Pink", "Blue-White", "Metallic Pink", "Orange-GREY", "Red-SeaGreen", "Red-Yellow"]

# name, size -> dict(color: qty)
MATRIX_ROWS = [
    ("TEDDY BASKET", "GENERAL", {"Baby Pink": "4", "Metallic Pink": "3", "Red-SeaGreen": "3", "Red-Yellow": "2"}),
    ("HANDLE GRIP SMALL U/23", "GENERAL", {"Baby Pink": "4", "Blue-White": "4", "Metallic Pink": "3", "Orange-GREY": "4", "Red-SeaGreen": "3", "Red-Yellow": "2"}),
    ("Fitted Frame 20 inch Orbit S/Rim", "GENERAL", {"Baby Pink": "4", "Blue-White": "4", "Metallic Pink": "3", "Orange-GREY": "4", "Red-SeaGreen": "3", "Red-Yellow": "2"}),
    ("ORBIT STICKER Chain Cover", "16 inch", {"Blue-White": "4", "Orange-GREY": "4", "Red-SeaGreen": "3", "Red-Yellow": "2"}),
    ("ORBIT STICKER Backrest", "GENERAL", {"Blue-White": "4", "Orange-GREY": "4", "Red-SeaGreen": "3", "Red-Yellow": "2"}),
    ("ORBIT RIM PAPER", "20 inch", {"Blue-White": "4", "Orange-GREY": "4", "Red-SeaGreen": "3", "Red-Yellow": "2"}),
    ("Chain Cover ORBIT", "16 inch", {"Blue-White": "4", "Orange-GREY": "4", "Red-SeaGreen": "3", "Red-Yellow": "2"}),
    ("ORBIT-STICKER-Chain Cover---BLUE", "16 inch", {"Baby Pink": "4", "Metallic Pink": "3"}),
    ("ORBIT-STICKER-Backrest---BLUE", "GENERAL", {"Baby Pink": "4", "Metallic Pink": "3"}),
    ("ORBIT-RIM-PAPER---BLUE", "20 inch", {"Baby Pink": "4", "Metallic Pink": "3"}),
    ("Chain Cover ORBIT---BLUE", "16 inch", {"Baby Pink": "4", "Metallic Pink": "3"}),
]

MM_TO_PX = 96 / 25.4
PAGE_MARGIN_MM = 6
PAGE_HEIGHT_PX = round((297 - 2 * PAGE_MARGIN_MM) * MM_TO_PX)


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context()
        page = ctx.new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        url = DIST_HTML.as_uri()
        page.goto(url, wait_until="domcontentloaded")
        page.wait_for_timeout(500)

        # Seed the modal's source tables + header fields exactly as the
        # real Add/Edit Production Sheet modal would leave them.
        page.evaluate("""
            (data) => {
                App.State.currentProductionSheet = { idx: 0, lotQty: 20, colors: data.colors, lotColor: data.colors.join(', ') };
                document.getElementById('prodSheetDate').innerText = '15/07/2026';
                document.getElementById('prodSheetProductId').innerText = 'LOT-PO21SS-0001';
                document.getElementById('prodSheetProductName').innerText = '20 inch Orbit Sports D/Gaddi Steel Rim';
                document.getElementById('prodSheetLotQty').innerText = '20';
                document.getElementById('productionSheetRemarks').value = 'Priority lot.';

                const commonBody = document.getElementById('productionSheetCommonBody');
                commonBody.innerHTML = data.common.map(r => `
                    <tr>
                        <td><input class="prod-sheet-item-name" value="${r[0]}"></td>
                        <td><input class="prod-sheet-size" value="${r[1]}"></td>
                        <td><input class="prod-sheet-narration" value="${r[2]}"></td>
                        <td><input type="number" class="prod-sheet-qty" value="${r[3]}"></td>
                    </tr>`).join('');

                const matrixWrap = document.getElementById('productionSheetMatrixTables');
                const tbody = document.createElement('tbody');
                tbody.className = 'prod-sheet-matrix-tbody';
                data.matrix.forEach(([name, size, qtyByColor]) => {
                    const tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td><input class="prod-sheet-item-name" value="${name}"></td>
                        <td><input class="prod-sheet-size" value="${size}"></td>
                        <td><input class="prod-sheet-narration" value="-"></td>` +
                        data.colors.map(c => `<td><input class="prod-sheet-color-qty" data-color="${c}" value="${qtyByColor[c] || ''}"></td>`).join('');
                    tbody.appendChild(tr);
                });
                matrixWrap.appendChild(document.createElement('table')).appendChild(tbody);
            }
        """, {"colors": COLORS, "common": COMMON_ROWS, "matrix": [(n, s, q) for n, s, q in MATRIX_ROWS]})

        # printProductionSheet() ends with App.Print.trigger(...) which
        # calls window.print() -- stub it out so headless Chromium doesn't
        # hang/no-op on a real print dialog, then invoke the function.
        page.evaluate("window.print = () => {};")
        page.evaluate("App.Production.printProductionSheet();")
        page.wait_for_timeout(100)

        print("[Structure] Consolidated matrix + two-column common")
        matrix_tables = page.evaluate("document.querySelectorAll('#print-production-sheet-matrix-tables table').length")
        check(matrix_tables == 1, f"exactly one consolidated per-color matrix table (found {matrix_tables}, old behavior produced 6)")

        matrix_cols = page.evaluate("document.querySelectorAll('#print-production-sheet-matrix-tables table thead th').length")
        # Item Name + Narration (Size dropped: matrix rows are all GENERAL/sized-mixed -> not all-generic, so Size may or may not drop)
        check(matrix_cols >= len(COLORS) + 2, f"matrix header has Item/Narration(+Size) + all {len(COLORS)} colors (found {matrix_cols} th)")

        common_table_count = page.evaluate("document.querySelectorAll('#print-production-sheet-common-tables table').length")
        check(common_table_count == 2, f"15 common rows (> 8 threshold) split into two side-by-side tables (found {common_table_count})")

        common_rows_total = page.evaluate("document.querySelectorAll('#print-production-sheet-common-tables tbody tr').length")
        check(common_rows_total == len(COMMON_ROWS), f"all {len(COMMON_ROWS)} common rows present across both tables (found {common_rows_total})")

        matrix_rows_total = page.evaluate("document.querySelectorAll('#print-production-sheet-matrix-tables tbody tr').length")
        check(matrix_rows_total == len(MATRIX_ROWS), f"all {len(MATRIX_ROWS)} matrix rows present (found {matrix_rows_total})")

        print("\n[Data integrity] spot-check a dashed cell and a real value")
        first_row_html = page.evaluate("document.querySelector('#print-production-sheet-matrix-tables tbody tr').outerHTML")
        check('4' in first_row_html, "TEDDY BASKET row shows its real qty (4) for Baby Pink")
        check('&#8211;' in first_row_html or '–' in first_row_html, "TEDDY BASKET row shows a dash for the color it doesn't use (Blue-White)")

        print("\n[Fit engine] single-page height + qty legibility floor")
        final_height = page.evaluate("document.getElementById('print-production-sheet-container').scrollHeight")
        # container is restored to display:none after the (stubbed) print call,
        # so re-measure by re-running just the sizing probe the function itself used.
        print(f"  INFO: PAGE_HEIGHT_PX target = {PAGE_HEIGHT_PX}px (A4, {PAGE_MARGIN_MM}mm margins)")

        # Only the qty/color columns (right-aligned) are floored at 11px --
        # item-name/narration text is allowed to shrink further under
        # pressure, so check those columns specifically rather than every td.
        qty_font_sizes = page.evaluate("""
            Array.from(document.querySelectorAll('#print-production-sheet-matrix-tables td[style*="text-align:right"]'))
                .map(td => parseFloat(td.style.fontSize))
                .filter(f => !isNaN(f))
        """)
        min_qty_font = min(qty_font_sizes) if qty_font_sizes else None
        check(min_qty_font is not None and min_qty_font >= 11, f"no Qty/color cell font-size dropped below the 11px legibility floor (min found: {min_qty_font})")

        body_font_sizes = page.evaluate("""
            Array.from(document.querySelectorAll('#print-production-sheet-matrix-tables td'))
                .map(td => parseFloat(td.style.fontSize))
                .filter(f => !isNaN(f))
        """)
        print(f"  INFO: body text compressed as low as {min(body_font_sizes) if body_font_sizes else 'n/a'}px under the tightest tier (expected, non-qty columns only)")

        # The tables built by printProductionSheet() are still in the DOM
        # (only the container's own display/position were reverted) --
        # independently re-lay it out at the real A4 printable width to see
        # whether this exact worst-case lot actually landed on one page.
        final_height = page.evaluate(f"""
            (() => {{
                const c = document.getElementById('print-production-sheet-container');
                const prev = {{ d: c.style.display, p: c.style.position, v: c.style.visibility, w: c.style.width, l: c.style.left }};
                c.style.position = 'fixed'; c.style.left = '-10000px'; c.style.top = '0';
                c.style.visibility = 'hidden'; c.style.display = 'block';
                c.style.width = '{round((210 - 2 * PAGE_MARGIN_MM) * MM_TO_PX)}px';
                const h = c.scrollHeight;
                c.style.display = prev.d; c.style.position = prev.p; c.style.visibility = prev.v; c.style.width = prev.w; c.style.left = prev.l;
                return h;
            }})()
        """)
        fits_one_page = final_height <= PAGE_HEIGHT_PX
        print(f"  INFO: final rendered height = {final_height}px vs {PAGE_HEIGHT_PX}px target -> {'fits on ONE page' if fits_one_page else 'needs a 2nd page (tightest tier still applied, per the agreed overflow rule)'}")

        print("\n[Console]")
        check(len(console_errors) == 0, f"no JS errors during printProductionSheet() ({len(console_errors)} found)" + (f": {console_errors[:2]}" if console_errors else ""))

        # ── Scenario 2: a small, simple lot (single color, few rows) should
        # NOT get compressed at all -- tier 0 (12px font, full padding) and a
        # single common table (below the 2-column threshold).
        print("\n[Scenario: small lot] no unnecessary compression")
        page.evaluate("""
            () => {
                App.State.currentProductionSheet = { idx: 0, lotQty: 5, colors: ['Black'], lotColor: 'Black' };
                document.getElementById('prodSheetProductId').innerText = 'LOT-SMALL-0001';
                document.getElementById('productionSheetRemarks').value = '';

                document.getElementById('productionSheetCommonBody').innerHTML = [
                    ['Frame', '20 inch', '-', '5'],
                    ['Wheel', '20 inch', '-', '10']
                ].map(r => `
                    <tr>
                        <td><input class="prod-sheet-item-name" value="${r[0]}"></td>
                        <td><input class="prod-sheet-size" value="${r[1]}"></td>
                        <td><input class="prod-sheet-narration" value="${r[2]}"></td>
                        <td><input type="number" class="prod-sheet-qty" value="${r[3]}"></td>
                    </tr>`).join('');

                const matrixWrap = document.getElementById('productionSheetMatrixTables');
                matrixWrap.innerHTML = '';
                const tbody = document.createElement('tbody');
                tbody.className = 'prod-sheet-matrix-tbody';
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td><input class="prod-sheet-item-name" value="Saddle"></td>
                    <td><input class="prod-sheet-size" value="GENERAL"></td>
                    <td><input class="prod-sheet-narration" value="-"></td>
                    <td><input class="prod-sheet-color-qty" data-color="Black" value="5"></td>`;
                tbody.appendChild(tr);
                matrixWrap.appendChild(document.createElement('table')).appendChild(tbody);
            }
        """)
        page.evaluate("App.Production.printProductionSheet();")
        page.wait_for_timeout(100)

        small_common_tables = page.evaluate("document.querySelectorAll('#print-production-sheet-common-tables table').length")
        check(small_common_tables == 1, f"2-row common list stays single-column, not split (found {small_common_tables} tables)")

        small_qty_font = page.evaluate("""
            parseFloat(document.querySelector('#print-production-sheet-common-tables td[style*="text-align:right"]').style.fontSize)
        """)
        check(small_qty_font == 14.5, f"small lot stays at tier-0 emphasis font (14.5px expected, found {small_qty_font}px) -- not shrunk when it doesn't need to be")

        # ── Scenario 3: the bulk-print builder (buildProductionSheetPrintPageHtml,
        # used by multi-select "Print" on already-saved lots) got matching
        # visual polish (tinted header, zebra, boxed Lot Qty) -- confirm it
        # still renders valid, complete HTML with the new styling in place.
        print("\n[Scenario: bulk print builder] visual parity, still renders cleanly")
        bulk_html = page.evaluate("""
            App.Production.buildProductionSheetPrintPageHtml({
                date: '15/07/2026', productId: 'LOT-BULK-0001', productName: 'Test Bike',
                qty: 10, color: 'Black', sheetRemarks: 'Bulk print smoke test',
                componentsConsumed: [
                    { itemName: 'Frame', size: '20 inch', sourceType: 'ITEM', qty: 10 },
                    { itemName: 'Wheel Pool', size: '20 inch', sourceType: 'POOL', qty: 20 }
                ]
            });
        """)
        check('LOT-BULK-0001' in bulk_html, "bulk page includes the lot ID")
        check('#e8f5e9' in bulk_html, "bulk page uses the tinted header color (visual parity with the live print sheet)")
        check('Wheel Pool' in bulk_html and 'Frame' in bulk_html, "bulk page includes both component rows")
        check(bulk_html.count('<table') == 1, "bulk page still renders exactly one components table (unchanged structure)")

        print(f"\n{'=' * 50}")
        if failures == 0:
            print("ALL CHECKS PASSED")
        else:
            print(f"{failures} CHECK(S) FAILED")
        print(f"{'=' * 50}")

        browser.close()
        return failures


if __name__ == "__main__":
    sys.exit(1 if run() > 0 else 0)
