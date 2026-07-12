"""
Verification script — PO Ledger status bar (PO Issued / Partially Received /
Completed tiles + counts), status badge column, and the "Pending Orders"
modal.

Status itself (po.status, item.receivedQty/pendingQty) is computed
server-side in module_po.js's getPOData()/_attachPoStatus(), which can't be
exercised from a static compiled preview (no live Apps Script backend) — so
this test seeds App.State.globalPOs with representative PO objects in
exactly the shape getPOData() returns (each item already carrying
receivedQty/pendingQty, each PO already carrying status) and verifies the
CLIENT rendering/filtering layer built on top of that data: the status bar
tile counts, the table's status badge column, tile-click filtering, and the
Pending Orders modal.

Run: python .pw-test/verify_po_status_bar.py
"""
import sys
import io
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

failures = 0
def check(cond, msg):
    global failures
    if cond:
        print(f"  PASS: {msg}")
    else:
        failures += 1
        print(f"  FAIL: {msg}")


SEED_PO = """
    App.State.globalPOs = [
      { poNumber: 'TEST-PO-1', poDate: '01/01/2026', vendor: 'Vendor A', totalQty: 10, grandTotal: 1000,
        status: 'PO Issued',
        items: [ { name: 'Item A', size: '', unit: 'Pcs', qty: 10, price: 100, receivedQty: 0, pendingQty: 10 } ] },
      { poNumber: 'TEST-PO-2', poDate: '02/01/2026', vendor: 'Vendor B', totalQty: 20, grandTotal: 2000,
        status: 'Partially Received',
        items: [ { name: 'Item B', size: '', unit: 'Pcs', qty: 20, price: 100, receivedQty: 8, pendingQty: 12 } ] },
      { poNumber: 'TEST-PO-3', poDate: '03/01/2026', vendor: 'Vendor C', totalQty: 5, grandTotal: 500,
        status: 'Completed',
        items: [ { name: 'Item C', size: '', unit: 'Pcs', qty: 5, price: 100, receivedQty: 5, pendingQty: 0 } ] }
    ];
    App.State.filteredPOs = [...App.State.globalPOs];
    App.State.poStatusFilter = 'all';
    App.PO.sortFiltered();
    App.PO.renderTable();
    App.PO.updateStatusBar();
"""


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context()
        page = ctx.new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(500)

        # Neutralize every module's own loadData()/Api.call round-trip so our
        # directly-seeded state can't be clobbered by an async response
        # arriving after the seed (same technique as
        # verify_bulk_delete_confirmations.py).
        page.evaluate("""
            window.google = {
                script: {
                    run: {
                        withSuccessHandler(cb) {
                            const runner = { withFailureHandler() { return runner; } };
                            return new Proxy(runner, { get(target, prop) {
                                if (prop in target) return target[prop];
                                return (...args) => {};
                            }});
                        }
                    }
                }
            };
        """)

        page.evaluate("document.getElementById('btn-poLedger').click();")
        page.wait_for_timeout(50)

        print("\n[Seed] 3 POs (1 Issued, 1 Partially Received, 1 Completed)")
        page.evaluate(SEED_PO)
        page.wait_for_timeout(50)

        row_count = page.locator("#poTableBody tr").count()
        check(row_count == 3, f"table renders 3 PO rows (got {row_count})")

        # Default sort is PO# descending, so don't assume display order —
        # key each row's badge by its own PO# cell instead.
        badges_by_po = page.evaluate("""
            Object.fromEntries([...document.querySelectorAll('#poTableBody tr')].map(tr => {
                const cells = tr.querySelectorAll('td');
                const poNum = cells[1].textContent.trim();
                const span = cells[7].querySelector('span.badge');
                return [poNum, { text: span.textContent.trim(), cls: span.className }];
            }))
        """)
        b1, b2, b3 = badges_by_po['PO-TEST-PO-1'], badges_by_po['PO-TEST-PO-2'], badges_by_po['PO-TEST-PO-3']
        check(b1['text'] == 'PO Issued' and 'bg-primary' in b1['cls'],
              f"TEST-PO-1 status badge is 'PO Issued' / bg-primary (got {b1})")
        check(b2['text'] == 'Partially Received' and 'bg-info' in b2['cls'],
              f"TEST-PO-2 status badge is 'Partially Received' / bg-info (got {b2})")
        check(b3['text'] == 'Completed' and 'bg-success' in b3['cls'],
              f"TEST-PO-3 status badge is 'Completed' / bg-success (got {b3})")

        print("\n[Status bar] tile counts reflect the seeded data")
        counts = page.evaluate("""
            ({
                all: document.getElementById('poStatusCountAll').textContent,
                issued: document.getElementById('poStatusCountIssued').textContent,
                partial: document.getElementById('poStatusCountPartial').textContent,
                completed: document.getElementById('poStatusCountCompleted').textContent,
                pending: document.getElementById('poStatusCountPending').textContent
            })
        """)
        check(counts['all'] == '3', f"All tile shows 3 (got {counts['all']})")
        check(counts['issued'] == '1', f"PO Issued tile shows 1 (got {counts['issued']})")
        check(counts['partial'] == '1', f"Partially Received tile shows 1 (got {counts['partial']})")
        check(counts['completed'] == '1', f"Completed tile shows 1 (got {counts['completed']})")
        check(counts['pending'] == '2', f"Pending Orders button badge shows 2 (got {counts['pending']})")

        print("\n[Filter] clicking 'PO Issued' tile shows only that PO")
        page.evaluate("App.PO.filterByStatus('PO Issued')")
        page.wait_for_timeout(50)
        row_count = page.locator("#poTableBody tr").count()
        check(row_count == 1, f"table shows 1 row after filtering to 'PO Issued' (got {row_count})")
        po_text = page.evaluate("document.querySelector('#poTableBody tr td:nth-child(2)').textContent")
        check('TEST-PO-1' in po_text, f"the visible row is TEST-PO-1 (got {po_text!r})")
        active_status = page.evaluate("""
            document.querySelector('.po-status-filter-btn.active')?.dataset.status
        """)
        check(active_status == 'PO Issued', f"'PO Issued' tile carries the active class (got {active_status!r})")

        print("\n[Filter] clicking 'Completed' tile shows only that PO")
        page.evaluate("App.PO.filterByStatus('Completed')")
        page.wait_for_timeout(50)
        row_count = page.locator("#poTableBody tr").count()
        check(row_count == 1, f"table shows 1 row after filtering to 'Completed' (got {row_count})")

        print("\n[Filter] clicking 'All' tile resets the filter")
        page.evaluate("App.PO.filterByStatus('all')")
        page.wait_for_timeout(50)
        row_count = page.locator("#poTableBody tr").count()
        check(row_count == 3, f"table shows all 3 rows again (got {row_count})")

        print("\n[Pending Orders modal] opens with only the 2 non-completed PO lines")
        page.evaluate("App.PO.openPendingPOsModal()")
        page.locator("#pendingPOsModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(100)

        pending_rows = page.locator("#pendingPOsTableBody tr").count()
        check(pending_rows == 2, f"pending-orders table has 2 rows (got {pending_rows})")

        pending_pos = page.evaluate("""
            [...document.querySelectorAll('#pendingPOsTableBody tr td:first-child')].map(td => td.textContent.trim())
        """)
        check(any('TEST-PO-1' in t for t in pending_pos), f"TEST-PO-1 (fully unbilled) appears in pending list (got {pending_pos})")
        check(any('TEST-PO-2' in t for t in pending_pos), f"TEST-PO-2 (partially billed) appears in pending list (got {pending_pos})")
        check(not any('TEST-PO-3' in t for t in pending_pos), f"TEST-PO-3 (fully completed) is excluded from pending list (got {pending_pos})")

        pending_qty_cell = page.evaluate("""
            [...document.querySelectorAll('#pendingPOsTableBody tr')]
                .find(tr => tr.querySelector('td').textContent.includes('TEST-PO-2'))
                .querySelectorAll('td')[6].textContent.trim()
        """)
        check('12' in pending_qty_cell, f"TEST-PO-2's pending qty cell shows 12 (got {pending_qty_cell!r})")

        print("\n[Pending Orders modal] search narrows the list")
        page.fill("#searchPendingPOs", "Vendor A")
        page.evaluate("App.PO.filterPendingPOs(document.getElementById('searchPendingPOs').value)")
        page.wait_for_timeout(50)
        filtered_rows = page.locator("#pendingPOsTableBody tr").count()
        check(filtered_rows == 1, f"searching 'Vendor A' narrows pending list to 1 row (got {filtered_rows})")

        if console_errors:
            print("\nConsole/page errors:")
            for e in console_errors:
                print(f"    {e}")
            globals()['failures'] += len(console_errors)

        browser.close()


if __name__ == "__main__":
    run()
    print(f"\n{'ALL PASS' if failures == 0 else str(failures) + ' FAILURE(S)'}")
    sys.exit(0 if failures == 0 else 1)
