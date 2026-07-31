"""
Audit: does any print/PDF template overflow the A4 printable box?

Every export path lays its container out at App.Print.PAGE_WIDTH_PX (748px =
210mm - 2x6mm margins) and captures exactly that box:
  - window.print()          -> browser clips at the @page margin box
  - Download PDF            -> html2canvas captures the element's own border box
  - Bulk print / Bulk PDF   -> same, via #print-bulk-container

So anything whose right edge passes the container's border box is GONE from the
output. This probes every bulk-page builder and every static single-record
template at that exact width with realistic worst-case data, and reports the
worst overflow per template.

Run: python .pw-test/audit_print_margins_all_templates.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"

# Real item names from this database — long compound tokens are the worst case
# for an auto-layout table's min-content width.
ITEMS = [
    ("FORK---FITTING---BRUT-BLACK", "GENERAL", "BROUT ED", 24, 145.50),
    ("CHAIN-WHEEL---BLACK--WITH--GUARD", "20 inch", "32T", 24, 320.00),
    ("COTTER-PIN---STD", "GENERAL", "Standard Cotter pin (50 pcs packing) (unit rate- 250)", 48, 5.00),
    ("FREE-WHEEL---HALF-BALL-18-T", "GENERAL", "18T", 24, 88.75),
    ("TYRE-BLACK--NEXO-3.00", "16 inch", "", 48, 210.00),
]

SETUP = """
() => {
  App.State.globalClients = [{
    name: 'SHRI BALAJI CYCLE TRADERS & DISTRIBUTORS',
    address: 'Shop 14, Gandhi Market, Near Bus Stand, Ludhiana 141008',
    gstin: '03AFIPS4089J1Z1', contact: '98765-43210'
  }];
  App.State.globalVendors = [];
  App.State.globalItems = [];
  App.companyLogo = null;
  return true;
}
"""

# Lays `containerId` out exactly as downloadElementAsPDF does, then finds the
# node whose right edge sits furthest past the container's own border box.
PROBE = """
(containerId) => {
  const el = document.getElementById(containerId);
  if (!el) return { missing: true };
  const prev = el.getAttribute('style');
  el.style.position = 'absolute';
  el.style.left = '0';
  el.style.top = '0';
  el.style.visibility = 'hidden';
  el.style.display = 'block';
  el.style.width = App.Print.PAGE_WIDTH_PX + 'px';
  el.style.maxWidth = 'none';

  const base = el.getBoundingClientRect();
  const clipRight = base.left + el.clientWidth;
  let worst = 0, worstNode = null;
  el.querySelectorAll('*').forEach(n => {
    const r = n.getBoundingClientRect();
    if (r.width === 0) return;
    const over = r.right - clipRight;
    if (over > worst) {
      worst = over;
      worstNode = n.tagName + (n.className ? '.' + String(n.className).split(' ')[0] : '')
        + ' | ' + (n.textContent || '').trim().replace(/\\s+/g, ' ').slice(0, 55);
    }
  });

  // Widest table, and which of its cells forces the width.
  let widestTable = null;
  el.querySelectorAll('table').forEach(t => {
    const w = t.getBoundingClientRect().width;
    if (!widestTable || w > widestTable.width) {
      const head = [...t.querySelectorAll('thead th')].map(th => ({
        text: (th.textContent || '').trim().slice(0, 18),
        w: Math.round(th.getBoundingClientRect().width),
        declared: (th.getAttribute('style') || '').match(/width:\\s*([\\d.]+%)/)?.[1] || '-'
      }));
      widestTable = { width: Math.round(w), cols: head.length, head };
    }
  });

  // Declared column widths must partition 100% — anything else means no
  // column gets the share it asks for, and the shortfall lands wherever the
  // renderer decides.
  const badSums = [];
  el.querySelectorAll('table').forEach(t => {
    const ths = [...t.querySelectorAll('thead th')];
    if (!ths.length) return;
    const decl = ths.map(th => (th.getAttribute('style') || '').match(/width:\\s*([\\d.]+)%/)?.[1]);
    if (decl.some(d => !d)) return;                 // not a width-partitioned table
    const sum = decl.reduce((a, b) => a + parseFloat(b), 0);
    if (Math.abs(sum - 100) > 0.5) {
      badSums.push({ cols: ths.length, sum, first: (ths[0].textContent || '').trim().slice(0, 18) });
    }
  });

  // The app's base `th` rule must not reach in here: a nowrap header cannot
  // wrap, so its min-content becomes the whole phrase and the auto-layout
  // table grows past width:100%; an overriding `color` replaces the white the
  // template sets on the header <tr>.
  // A th that declares its own color inline (the Production Sheet's headCell
  // does) is styling itself and is fine; the leak is a th with NO color of its
  // own whose computed color still differs from the row it should inherit.
  const thLeaks = [];
  el.querySelectorAll('thead th').forEach(th => {
    const cs = getComputedStyle(th);
    const inline = th.getAttribute('style') || '';
    const trColor = getComputedStyle(th.parentElement).color;
    const text = th.textContent.trim().slice(0, 18);
    if (cs.whiteSpace === 'nowrap' && !/white-space\\s*:/.test(inline)) {
      thLeaks.push({ kind: 'nowrap', text });
    } else if (!/(^|;)\\s*color\\s*:/.test(inline) && cs.color !== trColor) {
      thLeaks.push({ kind: `color ${cs.color} != row ${trColor}`, text });
    }
  });

  const out = {
    clientWidth: el.clientWidth,
    scrollWidth: el.scrollWidth,
    offsetHeight: el.offsetHeight,
    overflowPx: Math.round(worst),
    worstNode,
    widestTable,
    badSums,
    thLeaks: thLeaks.slice(0, 3),
    thLeakCount: thLeaks.length
  };
  el.setAttribute('style', prev);
  return out;
}
"""

BULK_CASES = {
    "PO (rates+total)": """() => App.Print.renderBulkPages([{
        poNumber: 'PO-2026-00147', poDate: '30/07/2026',
        vendor: 'JAGDAMBAY CYCLE INDUSTRIES PVT LTD',
        contact: '98765-43210', supplierRemarks: 'Deliver in 2 lots',
        description: 'Monthly frame parts order', remarks: 'Freight extra',
        items: %ITEMS%
      }], po => App.PO.buildPOPrintPageHtml(po, true, true))""",
    "PO (no rates)": """() => App.Print.renderBulkPages([{
        poNumber: 'PO-2026-00147', poDate: '30/07/2026',
        vendor: 'JAGDAMBAY CYCLE INDUSTRIES PVT LTD', items: %ITEMS%
      }], po => App.PO.buildPOPrintPageHtml(po, false, false))""",
    "Bill / Goods Receipt": """() => App.Print.renderBulkPages([{
        billNumber: 'INV-8842/26-27', billDate: '30/07/2026',
        vendor: 'JAGDAMBAY CYCLE INDUSTRIES PVT LTD', contact: '98765-43210',
        remarks: 'Received short by 2 pcs', poNumbers: ['2026-00147', '2026-00151'],
        items: %ITEMS%.map(i => Object.assign({}, i, { gstRatePct: 18, lineTotal: i.qty * i.price }))
      }], b => App.Bill.buildBillPrintPageHtml(b))""",
    "Dispatch challan": """() => App.Print.renderBulkPages([{
        dispatchNumber: 'DC-2026-0233', dispatchDate: '30/07/2026',
        clientName: 'SHRI BALAJI CYCLE TRADERS & DISTRIBUTORS',
        transport: 'Punjab Roadlines, Truck PB-10-EK-4471', invoiceNumber: 'INV-8842',
        grNumber: 'GR-99321', remarks: 'Handle with care',
        items: %ITEMS%.map(i => Object.assign({}, i, { productName: i.name, lineRemarks: i.narration }))
      }], b => App.Dispatch.buildDispatchPrintPageHtml(b))""",
    "Client order": """() => App.Print.renderBulkPages([{
        orderId: 'CO-2026-0088', orderDate: '30/07/2026',
        clientName: 'SHRI BALAJI CYCLE TRADERS & DISTRIBUTORS', remarks: 'Urgent',
        lines: %ITEMS%.map(i => ({ productName: i.name, productId: 'PRD-' + i.size, qty: i.qty, lineRemarks: i.narration }))
      }], o => App.Client.buildOrderPrintPageHtml(o))""",
    "Return note": """() => App.Print.renderBulkPages([{
        returnId: 'RET-2026-0031', returnDate: '30/07/2026',
        vendor: 'JAGDAMBAY CYCLE INDUSTRIES PVT LTD', remarks: 'Damaged in transit',
        items: %ITEMS%.map(i => Object.assign({}, i, { reason: 'Rusted / bent on arrival', lineTotal: i.qty * i.price }))
      }], r => App.Return.buildReturnPrintPageHtml(r))""",
    "Stock issue receipt": """() => App.Print.renderBulkPages([{
        issueId: 'ISS-2026-0412', issueDate: '30/07/2026', issuedTo: 'Sanjay Kumar',
        totalValue: 12500, remarks: 'Issued against LOT-FTD031-0001',
        items: %ITEMS%.map(i => Object.assign({}, i, { rate: i.price, value: i.qty * i.price }))
      }], i => App.Issue.buildIssuePrintPageHtml(i))""",
    "Process sheet": """() => App.Print.renderBulkPages([{
        processId: 'PRC-FTF', processName: 'Fitting Frame', outputItemName: 'Fitted Frame Assembled',
        sequence: 5, lotPrefix: 'FTF', processType: 'General',
        components: %ITEMS%.map(i => ({ itemName: i.name, size: i.size, narration: i.narration,
          colorGroup: 'METALLIC GREEN', qtyPerUnit: 1.5, remarks: 'Check torque before fitting' }))
      }], p => App.Process.buildProcessPrintPageHtml(p))""",
    "BOM cost sheet": """() => App.Print.renderBulkPages([{
        productId: 'PRD-JK16-IBC', productName: 'Jungle King 16 inch IBC 3.00 Steel Rim Unbranded',
        remarks: 'Costing as on 30/07/2026',
        additionalCosts: [{ description: 'Packing, labour and freight to depot', rate: 145 }],
        components: %ITEMS%.map(i => ({ itemName: i.name, size: i.size, narration: i.narration,
          vendor: 'JAGDAMBAY CYCLE INDUSTRIES', qtyPerUnit: 2, rate: i.price,
          lineCost: i.price * 2, processGroup: 'Fitting Frame' }))
      }], b => App.BOM.buildBOMPrintPageHtml(b))""",
    "Production sheet": """() => App.Print.renderBulkPages([{
        lotNumber: 'LOT-FTD031-0001', productionDate: '30/07/2026', qty: 24,
        processName: 'Fitting Frame', color: 'METALLIC GREEN', assignedTo: 'Sanjay Kumar',
        sheetRemarks: 'Handle painted frames with gloves',
        componentsConsumed: %ITEMS%.map(i => ({ itemName: i.name, size: i.size,
          sourceType: 'POOL', qty: i.qty, narration: i.narration }))
      }], p => App.Production.buildProductionSheetPrintPageHtml(p))""",
    "Item ledger": """() => {
        const td = 'padding:6px;border:1px solid #e5e5e5;';
        App.Item.getLedgerData = () => ({
          stockHtml: %ITEMS%.map(i => `<tr><td style="${td}">${i[1]}</td><td style="${td}">120</td><td style="${td}">96</td><td style="${td}">24 pending on PO-2026-00147</td></tr>`).join(''),
          compHtml: %ITEMS%.map(i => `<tr><td style="${td}">${i[1]}</td><td style="${td}">${i[2]}</td><td style="${td}">JAGDAMBAY CYCLE INDUSTRIES PVT LTD</td><td style="${td}">98765-43210</td><td style="${td}">${i[4]}</td><td style="${td}">${i[4]}</td></tr>`).join(''),
          histHtml: %ITEMS%.map(i => `<tr><td style="${td}">30/07/2026</td><td style="${td}">PRODUCTION</td><td style="${td}">LOT-FTD031-0001</td><td style="${td}">JAGDAMBAY CYCLE INDUSTRIES PVT LTD</td><td style="${td}">${i[1]}</td><td style="${td}">${i[2]}</td><td style="${td}">${i[3]}</td><td style="${td}">${i[3]}</td><td style="${td}">${i[3]}</td><td style="${td}">${i[4]}</td></tr>`).join('')
        });
        App.Print.renderBulkPages(['CHAIN-WHEEL---BLACK--WITH--GUARD'], n => App.Item.buildItemLedgerPrintPageHtml(n));
      }""",
    "Vendor ledger": """() => {
        App.Vendor.calculateLedgerAndPending = () => ({
          ledger: %ITEMS%.map(i => ({ dateStr: '30/07/2026', type: 'PURCHASE ORDER', ref: 'PO-2026-00147',
            items: i[0] + ', ' + i[2], orderQty: i[3], incomingQty: i[3], value: i[3] * i[4] })),
          pendingList: %ITEMS%.map(i => ({ name: i[0], size: i[1], ordered: i[3], received: 0, pending: i[3] }))
        });
        App.Print.renderBulkPages([{ name: 'JAGDAMBAY CYCLE INDUSTRIES PVT LTD',
          gstin: '03AFIPS4089J1Z1', contact: '98765-43210',
          address: '6-B Shiv Shakti Estate, Verka Chowk, Dehlon Road, Ludhiana 141114',
          remarks: 'Credit 45 days' }], v => App.Vendor.buildVendorLedgerPrintPageHtml(v));
      }""",
}

ITEMS_JS = json.dumps([{"name": n, "size": s, "narration": nr, "qty": q, "price": p,
                        "unit": "Pcs", "0": n} for (n, s, nr, q, p) in ITEMS])
ITEMS_TUPLES_JS = json.dumps([[n, s, nr, q, p] for (n, s, nr, q, p) in ITEMS])

# The static single-record templates in View_Print.html are populated by their
# own module at print time; here every <tbody> is filled generically with rows
# matching its own <thead>'s column count, so the geometry under test is the
# template's, not any one module's data path.
STATIC_CONTAINERS = [
    "print-po-container",
    "print-bill-container",
    "print-dispatch-container",
    "print-item-ledger-container",
    "print-vendor-ledger-container",
    "print-client-ledger-container",
    "print-contractor-ledger-container",
    "print-low-stock-container",
    "print-bom-container",
]

FILL_STATIC = """
(containerId) => {
  const el = document.getElementById(containerId);
  if (!el) return false;
  const CELLS = %CELLS%;
  el.querySelectorAll('table').forEach(t => {
    const cols = t.querySelectorAll('thead th').length;
    const body = t.querySelector('tbody');
    if (!cols || !body) return;
    body.innerHTML = '';
    for (let r = 0; r < 3; r++) {
      const tr = document.createElement('tr');
      for (let c = 0; c < cols; c++) {
        const td = document.createElement('td');
        td.style.cssText = 'padding:6px;border:1px solid #e5e5e5;';
        td.textContent = CELLS[(r * cols + c) % CELLS.length];
        tr.appendChild(td);
      }
      body.appendChild(tr);
    }
  });
  return true;
}
"""


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_context(viewport={"width": 1400, "height": 900}).new_page()
        errors = []
        page.on("pageerror", lambda e: errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1200)
        page.evaluate(SETUP)

        results = []
        for label, js in BULK_CASES.items():
            # Object-shaped items for builders that read .name/.qty; tuple-shaped
            # for the ledger stubs that index positionally.
            src = js.replace("%ITEMS%", ITEMS_TUPLES_JS if "[0]" in js or "[4]" in js else ITEMS_JS)
            try:
                page.evaluate(src)
            except Exception as exc:
                results.append((label, None, f"builder threw: {str(exc)[:120]}"))
                continue
            g = page.evaluate(PROBE, "print-bulk-container")
            results.append((label, g, None))

        cells_js = json.dumps([
            "30/07/2026", "PRODUCTION", "LOT-FTD031-0001",
            "JAGDAMBAY CYCLE INDUSTRIES PVT LTD", "GENERAL",
            "CHAIN-WHEEL---BLACK--WITH--GUARD", "48", "48", "24", "12,450.00",
        ])
        fill_js = FILL_STATIC.replace("%CELLS%", cells_js)
        for cid in STATIC_CONTAINERS:
            if not page.evaluate(fill_js, cid):
                results.append((cid + " (static)", None, "container not found"))
                continue
            g = page.evaluate(PROBE, cid)
            results.append((cid.replace("print-", "").replace("-container", "") + " (static)", g, None))

        print(f"{'TEMPLATE':<26} {'OVERFLOW':>9}  WIDEST TABLE / OFFENDER")
        print("-" * 100)
        bad = []
        for label, g, err in results:
            if err:
                print(f"{label:<26} {'ERROR':>9}  {err}")
                bad.append(label)
                continue
            wt = g.get("widestTable") or {}
            tw = wt.get("width", "-")
            flag = "  <== OVERFLOWS" if g["overflowPx"] > 2 else ""
            print(f"{label:<26} {g['overflowPx']:>7}px  table={tw}px cols={wt.get('cols', '-')}{flag}")
            if g["overflowPx"] > 2:
                bad.append(label)
                print(f"{'':<26} {'':>9}  offender: {g['worstNode']}")
                print(f"{'':<26} {'':>9}  headers: "
                      + ", ".join(f"{h['text']}={h['w']}px(decl {h['declared']})" for h in wt.get("head", [])))

        print("-" * 100)
        print(f"capture box = {results[0][1]['clientWidth']}px wide"
              if results and results[0][1] else "")
        if bad:
            print(f"FAIL: {len(bad)} of {len(results)} templates overflow: {bad}")
        else:
            print(f"PASS: all {len(results)} templates fit the {results[0][1]['clientWidth']}px capture box")

        # The two upstream causes, asserted directly so a regression names
        # itself instead of surfacing only as a mystery overflow.
        leaks = [(l, g["thLeaks"]) for l, g, e in results if g and g["thLeakCount"]]
        if leaks:
            print(f"\nFAIL: the app's base `th` styling reaches into {len(leaks)} template(s) "
                  "(see the PRINT/PDF TEMPLATE ISOLATION block in Styles.html)")
            for label, sample in leaks[:5]:
                print(f"  {label}: {sample}")
        else:
            print("\nPASS: no print <th> inherits the app's nowrap/color from the screen table styling")

        sums = [(l, g["badSums"]) for l, g, e in results if g and g["badSums"]]
        if sums:
            print(f"\nFAIL: {len(sums)} template(s) declare column widths that do not add to 100%")
            for label, s in sums[:5]:
                print(f"  {label}: {s}")
        else:
            print("PASS: every width-partitioned print table's columns add to 100%")

        if errors:
            print("\nPage errors:")
            for e in errors[:10]:
                print("  " + e)
            bad.append("page errors")

        browser.close()
        ok = not bad and not leaks and not sums
        print("\n" + ("ALL PASS" if ok else "SOME FAILED"))
        return ok


if __name__ == "__main__":
    sys.exit(0 if run() else 1)
