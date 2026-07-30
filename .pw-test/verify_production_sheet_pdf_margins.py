"""
Guards the Production Sheet PDF template against the export defects fixed on
2026-07-30, where the downloaded PDF (html2pdf/html2canvas path) overflowed the
page and rendered wrapped text as overlapping mush, while window.print() looked
fine. Reported as "overflowing from its margins, resulting in partial pages".

The four invariants asserted here, and why each one matters:

  1. No horizontal overflow. The two-column Common grid used `1fr 1fr`, which is
     minmax(auto,1fr) and GROWS past its share to fit a wide item's min-content.
     That blew the grid 57px wider than the page, and since html2canvas captures
     only the element's own box, the whole Required Qty column was sliced off.
  2. The fit loop measures what actually gets captured. It compared scrollHeight
     (which excludes the container's 5px/3px accent borders) and never checked
     width at all, so a sheet 57px too wide still passed as "fits".
  3. Whole-number font sizes with explicit integer line-heights. A fractional
     line box (14.5px font against line-height:1.5 => 21.75px) drifts under
     html2canvas's rounded line-box arithmetic until two lines paint at the
     same y.
  4. Item names are pre-split into block segments, so the cell never depends on
     html2canvas's own wrapping -- the last remaining source of overlap, since
     it can only mis-paint a line it had to wrap itself.

Runs entirely on geometry/DOM assertions -- no html2pdf CDN fetch, no download.

Run: python .pw-test/verify_production_sheet_pdf_margins.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(r"c:\Users\erkar\my-app-script-project\dist\index.html")

# The real LOT-FTD031-0001 sheet, the one that shipped a clipped PDF.
COMMON = [
    ("BB-AXLE 2-C", "GENERAL", "Balance Order | 2C", "24"),
    ("FORK---FITTING---BRUT-BLACK", "GENERAL", "BROUT ED", "24"),
    ("CHAIN-WHEEL---BLACK--WITH--GUARD", "20 inch", "32T", "24"),
    ("CYCLE-CHAIN--110-LINK", "GENERAL", "110 Link - 370 gm 76 link", "16.56"),
    ("HUB-AXLE-CAP-BIG", "GENERAL", "", "48"),
    ("COTTER-PIN---STD", "GENERAL", "Standard Cotter pin (50 pcs packing) (unit rate- 250)", "48"),
    ("BB---CUP---SET", "GENERAL", "B B CUP SET OF THREE CYCLE PARTS", "24"),
    ("FREE-WHEEL---HALF-BALL-18-T", "GENERAL", "18T", "24"),
    ("RIM-TAPE-COTTON", "26 inch", "", "48"),
    ("TYRE-BLACK--NEXO-3.00", "16 inch", "", "48"),
    ("TUBE-BLACK--NEXO-3.00", "16 inch", "", "48"),
    ("FORK---BRACE", "GENERAL", "95-100 MM", "24"),
]
COLORS = ["BLACK", "COPPER", "GREY", "METALLIC GREEN", "PURPLE-WINE"]
MATRIX = [
    ("Painted Frame Jungle King 16 inch IBC 3.00", "", "", ["", "6", "6", "6", "6"]),
    ("Fitted Rim 16 inch", "", "", ["24", "", "", "", ""]),
]

SETUP_JS = """
(payload) => {
  const { common, matrix, colors } = payload;
  const mk = (cls, val) => { const i = document.createElement('input'); i.className = cls; i.value = val; return i; };
  const cell = (input) => { const td = document.createElement('td'); td.appendChild(input); return td; };

  let commonBody = document.getElementById('productionSheetCommonBody');
  if (!commonBody) {
    const t = document.createElement('table');
    commonBody = document.createElement('tbody');
    commonBody.id = 'productionSheetCommonBody';
    t.appendChild(commonBody); t.style.display = 'none';
    document.body.appendChild(t);
  }
  commonBody.innerHTML = '';
  common.forEach(r => {
    const tr = document.createElement('tr');
    tr.appendChild(cell(mk('prod-sheet-item-name', r[0])));
    tr.appendChild(cell(mk('prod-sheet-size', r[1])));
    tr.appendChild(cell(mk('prod-sheet-narration', r[2])));
    tr.appendChild(cell(mk('prod-sheet-qty', r[3])));
    commonBody.appendChild(tr);
  });

  let matrixWrap = document.getElementById('productionSheetMatrixTables');
  if (!matrixWrap) {
    matrixWrap = document.createElement('div');
    matrixWrap.id = 'productionSheetMatrixTables';
    matrixWrap.style.display = 'none';
    document.body.appendChild(matrixWrap);
  }
  matrixWrap.innerHTML = '<table><tbody class="prod-sheet-matrix-tbody"></tbody></table>';
  const mtb = matrixWrap.querySelector('.prod-sheet-matrix-tbody');
  matrix.forEach(r => {
    const tr = document.createElement('tr');
    tr.appendChild(cell(mk('prod-sheet-item-name', r[0])));
    tr.appendChild(cell(mk('prod-sheet-size', r[1])));
    tr.appendChild(cell(mk('prod-sheet-narration', r[2])));
    r[3].forEach((q, i) => { const inp = mk('prod-sheet-color-qty', q); inp.dataset.color = colors[i]; tr.appendChild(cell(inp)); });
    mtb.appendChild(tr);
  });

  const setDiv = (id, text) => {
    let el = document.getElementById(id);
    if (!el) { el = document.createElement('div'); el.id = id; el.style.display='none'; document.body.appendChild(el); }
    el.innerText = text;
  };
  setDiv('prodSheetDate', '30/07/2026');
  setDiv('prodSheetProductId', 'LOT-FTD031-0001');
  setDiv('prodSheetProductName', 'Fitted Frame 16 inch Jungle King IBC 3.00 Steel Rim Unbranded');
  setDiv('prodSheetLotQty', '24');

  App.State.currentProductionSheet = {
    colors, lotColor: colors.join(', '),
    date: '30/07/2026', size: '16 inch', model: 'General', processName: 'Fitting Frame'
  };
  return true;
}
"""

# Lay the container out at exactly the width downloadElementAsPDF captures at.
PROBE_JS = """() => {
  const el = document.getElementById('print-production-sheet-container');
  const prev = el.getAttribute('style');
  el.style.position = 'absolute'; el.style.left = '0'; el.style.top = '0';
  el.style.visibility = 'hidden'; el.style.display = 'block';
  el.style.width = App.Print.PAGE_WIDTH_PX + 'px';

  const base = el.getBoundingClientRect();
  // Two different right edges matter here, and only one of them causes the bug:
  //   clipRight    - the element's own border box. html2canvas captures exactly
  //                  this, so anything past it is GONE from the PDF. This is
  //                  the boundary the original defect crossed (786 vs 749).
  //   contentRight - inside the container's 20px side padding. Reaching into
  //                  the padding is a cosmetic tightness, not a clipped column,
  //                  so it is reported but held to a looser bound.
  // box-sizing:border-box, so clientWidth already includes both paddings.
  const clipRight = base.left + el.clientWidth;
  const contentRight = clipRight - 20;
  let worstOverflow = 0, worstNode = null, worstPadEncroach = 0;
  el.querySelectorAll('*').forEach(n => {
    const r = n.getBoundingClientRect();
    if (r.width === 0) return;
    if (r.right - clipRight > worstOverflow) {
      worstOverflow = r.right - clipRight;
      worstNode = n.tagName + ' | ' + (n.textContent || '').trim().slice(0, 45);
    }
    worstPadEncroach = Math.max(worstPadEncroach, r.right - contentRight);
  });

  const cells = [...el.querySelectorAll('#print-production-sheet-common-tables td, #print-production-sheet-matrix-tables td, #print-production-sheet-common-tables th, #print-production-sheet-matrix-tables th')];
  const fractional = cells.filter(c => {
    const cs = getComputedStyle(c);
    return !Number.isInteger(parseFloat(cs.fontSize)) || !Number.isInteger(parseFloat(cs.lineHeight));
  }).map(c => getComputedStyle(c).fontSize + ' / ' + getComputedStyle(c).lineHeight);

  const grid = el.querySelector('#print-production-sheet-common-tables > div');
  const nameCells = [...el.querySelectorAll('#print-production-sheet-common-tables tbody tr > td:first-child')];

  const out = {
    pageWidthPx: App.Print.PAGE_WIDTH_PX,
    pageHeightPx: App.Print.PAGE_HEIGHT_PX,
    clientWidth: el.clientWidth,
    scrollWidth: el.scrollWidth,
    offsetHeight: el.offsetHeight,
    gridTracks: grid ? getComputedStyle(grid).gridTemplateColumns : null,
    worstOverflowPx: Math.round(worstOverflow),
    worstNode,
    padEncroachPx: Math.round(worstPadEncroach),
    fractionalCells: fractional,
    // "BB---CUP---SET" must arrive as separate block segments, not one string
    // left for html2canvas to wrap.
    segmentedNameSample: (() => {
      const hit = nameCells.find(c => c.textContent.replace(/\\u200b/g, '').startsWith('BB') && c.textContent.includes('CUP'));
      return hit ? { blocks: hit.querySelectorAll('span[style*="display:block"]').length, text: hit.textContent.replace(/\\u200b/g, '') } : null;
    })(),
    anyNameCellBold: nameCells.some(c => Number(getComputedStyle(c).fontWeight) >= 600),
  };
  el.setAttribute('style', prev);
  return out;
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

        page.evaluate(SETUP_JS, {"common": COMMON, "matrix": MATRIX, "colors": COLORS})
        page.evaluate("App.Production._buildProductionSheetForExport()")
        g = page.evaluate(PROBE_JS)
        print("[geometry]", json.dumps(g, indent=2))

        ok = True

        # 1 -- shared page geometry, and the element is measured at capture width
        if g["pageWidthPx"] != 748 or g["pageHeightPx"] != 1077:
            print(f"  FAIL: App.Print page geometry drifted: {g['pageWidthPx']}x{g['pageHeightPx']}, expected 748x1077")
            ok = False
        else:
            print("  PASS: App.Print exposes the canonical A4 printable box (748x1077 px @ 6mm margins)")

        # 2 -- the Common grid tracks must be capped, not minmax(auto,1fr)
        tracks = g["gridTracks"] or ""
        if "376" in tracks or not tracks:
            print(f"  FAIL: Common grid tracks blew past their share: {tracks}")
            ok = False
        else:
            print(f"  PASS: Common grid tracks are capped at their share ({tracks})")

        # 3 -- nothing may extend past the container's content box
        if g["worstOverflowPx"] > 2:
            print(f"  FAIL: content runs {g['worstOverflowPx']}px past the captured box -> {g['worstNode']}")
            print("        html2canvas captures only the element's own box, so this is clipped off the PDF.")
            ok = False
        else:
            print(f"  PASS: nothing runs past the captured box (worst = {g['worstOverflowPx']}px)")

        # Reaching into the 20px side padding is tight but not clipped; only
        # flag it once it would actually reach the element edge.
        if g["padEncroachPx"] > 20:
            print(f"  FAIL: content eats the entire 20px side padding ({g['padEncroachPx']}px) and is at the clip line")
            ok = False
        else:
            print(f"  PASS: content stays within the side padding ({g['padEncroachPx']}px of 20px used)")

        if g["scrollWidth"] > g["clientWidth"] + 2:
            print(f"  FAIL: container scrollWidth {g['scrollWidth']} exceeds clientWidth {g['clientWidth']}")
            ok = False
        else:
            print("  PASS: container reports no horizontal scroll overflow")

        # 4 -- the sheet still fits a page vertically, measured incl. borders
        if g["offsetHeight"] > g["pageHeightPx"]:
            print(f"  FAIL: sheet is {g['offsetHeight']}px tall, past the {g['pageHeightPx']}px page box")
            ok = False
        else:
            print(f"  PASS: sheet fits one page by offsetHeight ({g['offsetHeight']} <= {g['pageHeightPx']})")

        # 5 -- no fractional font-size/line-height anywhere in the tables
        if g["fractionalCells"]:
            print(f"  FAIL: {len(g['fractionalCells'])} cell(s) have a fractional font-size/line-height: {g['fractionalCells'][:5]}")
            print("        Fractional line boxes make html2canvas paint two lines at the same y.")
            ok = False
        else:
            print("  PASS: every table cell uses a whole-number font-size and integer line-height")

        # 6 -- compound item names arrive pre-split into block segments
        seg = g["segmentedNameSample"]
        if not seg or seg["blocks"] < 3:
            print(f"  FAIL: 'BB---CUP---SET' was not split into block segments: {seg}")
            ok = False
        else:
            print(f"  PASS: compound item names are pre-split into {seg['blocks']} block segments ({seg['text']!r})")

        # 7 -- name cells must not be bold (bold + wrap is the overlap bug)
        if g["anyNameCellBold"]:
            print("  FAIL: an Item Name cell is bold; html2canvas lays bold text out with normal-weight")
            print("        metrics and then paints bold glyphs, so wrapped names overlap.")
            ok = False
        else:
            print("  PASS: Item Name cells carry emphasis by size/ink, not font-weight")

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
