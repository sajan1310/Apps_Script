"""
P3 Verification — "Open every tab with its sheet temporarily empty — every
table must show an empty state, never a blank void."

Covers all 24 App.State.global<X>-driven render functions found by research
(the 3 Dashboard mini-widgets take their data as function parameters, not
from App.State, so can't be forced empty this way, and are out of scope —
same for the single per-lot component-sheet render, which isn't a
list-of-records table). For each: empties the exact backing array(s) (some
have non-obvious dependencies — Warehouse Pool is driven by globalProcesses,
not globalWarehousePool; Production's "All Activity" needs both
globalProduction AND globalIssues empty), calls the render function
directly, and asserts the friendly empty-state UI (a separate div, or an
inline "No X yet" tbody row) is what actually shows — not a blank tbody.

Driven entirely through direct DOM/state inspection (no tab navigation, no
network) since visibility of a `display:none` tab container doesn't affect
whether its child elements' style/classList are inspectable via JS.

Run: python .pw-test/verify_empty_states_all_tables.py
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


# Each entry: name, seed JS (must leave the module's own state truly empty),
# render call, tbody id, and either div=<emptyStateDivId> or
# inline=<expected substring in the tbody's own empty-state row>.
MODULES = [
    {"name": "PO", "seed": "App.State.globalPOs = []; App.State.filteredPOs = [];",
     "render": "App.PO.renderTable();", "tbody": "poTableBody", "div": "poEmptyState"},
    {"name": "Bill", "seed": "App.State.globalBills = []; App.State.filteredBills = [];",
     "render": "App.Bill.renderTable();", "tbody": "billTableBody", "div": "billEmptyState"},
    {"name": "Return", "seed": "App.State.globalReturns = []; App.State.filteredReturns = [];",
     "render": "App.Return.renderTable();", "tbody": "returnTableBody", "div": "returnEmptyState"},
    {"name": "Wastage", "seed": "App.State.globalWastage = []; App.State.filteredWastage = [];",
     "render": "App.Wastage.renderTable();", "tbody": "wastageTableBody", "div": "wastageEmptyState"},
    {"name": "Issue", "seed": "App.State.globalIssues = []; App.State.filteredIssues = [];",
     "render": "App.Issue.renderTable();", "tbody": "issueTableBody", "div": "issueEmptyState"},
    {"name": "Item", "seed": "App.State.globalItems = []; App.State.filteredItems = [];",
     "render": "App.Item.renderTable();", "tbody": "itemTableBody", "inline": "No items yet"},
    {"name": "Vendor", "seed": "App.State.globalVendors = []; App.State.filteredVendors = [];",
     "render": "App.Vendor.renderTable();", "tbody": "vendorTableBody", "inline": "No vendors yet"},
    {"name": "Unit", "seed": "App.State.globalUnits = [];",
     "render": "App.Unit.renderTable();", "tbody": "unitTableBody", "inline": "No units defined yet"},
    {"name": "Color", "seed": "App.State.globalColors = [];",
     "render": "App.Color.renderTable();", "tbody": "colorMasterTableBody", "inline": "No colors defined yet"},
    {"name": "Model", "seed": "App.State.globalModels = [];",
     "render": "App.Model.renderTable();", "tbody": "modelMasterTableBody", "inline": "No models defined yet"},
    {"name": "ProcessType", "seed": "App.State.globalProcessTypes = [];",
     "render": "App.ProcessType.renderTable();", "tbody": "processTypeMasterTableBody", "inline": "No process types"},
    # Stock intentionally does belt-and-suspenders: an inline "No stock
    # records found." row AND the separate div, shown together.
    {"name": "Stock (main)", "seed": "App.State.globalStock = []; App.State.filteredStock = [];",
     "render": "App.Stock.renderTable();", "tbody": "stockTableBody", "div": "stockEmptyState", "allowInlineRowToo": True},
    {"name": "Warehouse Pool", "seed": "App.State.globalProcesses = [];",
     "render": "App.Stock.renderWarehousePoolTable();", "tbody": "warehousePoolTableBody", "div": "warehousePoolEmptyState"},
    {"name": "Warehouse Pool Opening", "seed": "App.State.globalWarehousePoolOpening = [];",
     "render": "App.Stock.renderWarehouseOpeningTable();", "tbody": "warehouseOpeningTableBody", "inline": "No opening stock recorded yet"},
    {"name": "BOM", "seed": "App.State.globalBOMs = []; App.State.filteredBOMs = [];",
     "render": "App.BOM.renderTable();", "tbody": "bomTableBody", "div": "bomEmptyState"},
    {"name": "Process", "seed": "App.State.globalProcesses = []; App.State.filteredProcesses = [];",
     "render": "App.Process.renderTable();", "tbody": "processTableBody", "div": "processEmptyState"},
    {"name": "Contractor", "seed": "App.State.globalContractors = []; App.State.filteredContractors = [];",
     "render": "App.Contractor.renderTable();", "tbody": "contractorTableBody", "div": "contractorEmptyState"},
    {"name": "Production (log)", "seed": "App.State.globalProduction = []; App.State.filteredProduction = [];",
     "render": "App.Production.renderTable();", "tbody": "productionTableBody", "div": "productionEmptyState"},
    {"name": "Production (All Activity)", "seed": "App.State.globalProduction = []; App.State.globalIssues = [];",
     "render": "App.Production.renderAllActivity();", "tbody": "productionAllTableBody", "div": "productionAllEmptyState"},
    {"name": "Clients", "seed": "App.State.globalClients = []; App.State.filteredClients = [];",
     "render": "App.Client.renderClientsTable();", "tbody": "clientsTableBody", "div": "clientsEmptyState"},
    {"name": "PI / Estimates", "seed": "App.State.globalOrders = []; App.State.filteredOrders = [];",
     "render": "App.Client.renderOrdersTable();", "tbody": "clientOrdersTableBody", "div": "clientOrdersEmptyState"},
    {"name": "Pending Orders", "seed": "App.State.filteredPendingOrders = [];",
     "render": "App.Client.renderPendingOrdersTable();", "tbody": "pendingOrdersTableBody", "div": "pendingOrdersEmptyState"},
    {"name": "Ready to Dispatch", "seed": "App.State.globalReadyToDispatch = []; App.State.filteredReadyToDispatch = [];",
     "render": "App.Dispatch.renderReadyTable();", "tbody": "readyToDispatchTableBody", "div": "readyToDispatchEmptyState"},
    {"name": "Dispatched Goods", "seed": "App.State.globalDispatch = []; App.State.filteredDispatch = [];",
     "render": "App.Dispatch.renderDispatchTable();", "tbody": "dispatchTableBody", "div": "dispatchEmptyState"},
]


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

        for mod in MODULES:
            print(f"\n[{mod['name']}] table shows a friendly empty state when its data is empty")
            page.evaluate(mod["seed"] + " " + mod["render"])

            tbody_html = page.evaluate(f"document.getElementById('{mod['tbody']}')?.innerHTML ?? null")
            check(tbody_html is not None, f"#{mod['tbody']} exists in the DOM")
            if tbody_html is None:
                continue

            # No real per-record rows should remain — a row is only
            # acceptable if it's the empty-state row itself (inline case)
            # or the tbody is blank (div case).
            row_count = page.evaluate(f"document.getElementById('{mod['tbody']}').querySelectorAll('tr').length")

            if "div" in mod:
                max_rows = 1 if mod.get("allowInlineRowToo") else 0
                check(row_count <= max_rows, f"tbody has no real data rows (got {row_count}, max allowed {max_rows})")
                visible = page.evaluate(f"""
                    (() => {{
                        const el = document.getElementById('{mod['div']}');
                        if (!el) return null;
                        return getComputedStyle(el).display !== 'none';
                    }})()
                """)
                check(visible is True, f"#{mod['div']} empty-state div is visible (got {visible})")
            else:
                check(row_count == 1, f"tbody has exactly 1 row (the empty-state row) (got {row_count})")
                check(mod["inline"] in tbody_html, f"empty-state row contains {mod['inline']!r} (got: {tbody_html[:200]!r})")

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
