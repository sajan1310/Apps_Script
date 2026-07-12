"""
P3 Verification — "Bulk-select 2 records in every ledger and press Delete —
a confirmation must appear in all 11 modules." (Actual count found by
research: 14 distinct bulk-delete flows across 13 App.<Module> namespaces —
App.Client alone has two, Clients and PI/Estimate Orders.)

For each flow: seeds 2 mock records directly into App.State, renders the
table, checks 2 row checkboxes, clicks "Delete Selected", and asserts
App.Utils.confirmAction fired (App.State.confirmCallback set, confirm
message shown) BEFORE any Api.call/google.script.run round-trip — proving
the delete is gated behind confirmation, not fired immediately.

Run: python .pw-test/verify_bulk_delete_confirmations.py
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


MODULES = [
    {
        "name": "PO",
        "nav": "document.getElementById('btn-poLedger').click();",
        "seed": """
            App.State.globalPOs = [
              {poNumber:'SWP-PO-1', poDate:'01/01/2026', vendor:'Vendor A', items:[{name:'Item',size:'',qty:1,unit:'Pcs'}], totalQty:1, grandTotal:100},
              {poNumber:'SWP-PO-2', poDate:'02/01/2026', vendor:'Vendor B', items:[{name:'Item',size:'',qty:1,unit:'Pcs'}], totalQty:1, grandTotal:200}
            ];
            App.State.filteredPOs = App.State.globalPOs;
            App.PO.renderTable();
        """,
        "checkbox": ".po-select-chk", "button": "#btnBulkDeletePOs",
    },
    {
        "name": "Bill",
        "nav": "document.getElementById('btn-billLedger').click();",
        "seed": """
            App.State.globalBills = [
              {billNumber:'SWP-B-1', billDate:'01/01/2026', vendor:'Vendor A', items:[{name:'Item',size:'',qty:1,unit:'Pcs',price:10,gstRatePct:18,lineTotal:10}], totalQty:1, totalAmount:10},
              {billNumber:'SWP-B-2', billDate:'02/01/2026', vendor:'Vendor B', items:[{name:'Item',size:'',qty:1,unit:'Pcs',price:10,gstRatePct:18,lineTotal:10}], totalQty:1, totalAmount:10}
            ];
            App.State.filteredBills = App.State.globalBills;
            App.Bill.renderTable();
        """,
        "checkbox": ".bill-select-chk", "button": "#btnBulkDeleteBills",
    },
    {
        "name": "Return",
        "nav": "document.getElementById('btn-returnLedger').click();",
        "seed": """
            App.State.globalReturns = [
              {returnNumber:'SWP-R-1', returnDate:'01/01/2026', vendor:'Vendor A', items:[{name:'Item',size:'',qty:1,unit:'Pcs',price:10,reason:'x',lineTotal:10}], totalQty:1, totalAmount:10},
              {returnNumber:'SWP-R-2', returnDate:'02/01/2026', vendor:'Vendor B', items:[{name:'Item',size:'',qty:1,unit:'Pcs',price:10,reason:'x',lineTotal:10}], totalQty:1, totalAmount:10}
            ];
            App.State.filteredReturns = App.State.globalReturns;
            App.Return.renderTable();
        """,
        "checkbox": ".return-select-chk", "button": "#btnBulkDeleteReturns",
    },
    {
        "name": "Wastage",
        "nav": "document.getElementById('btn-returnLedger').click();",
        "seed": """
            App.State.globalWastage = [
              {wastageId:'SWP-W-1', date:'01/01/2026', vendor:'', items:[{name:'Item',size:'',qty:1,unit:'Pcs',reason:'x'}], totalQty:1},
              {wastageId:'SWP-W-2', date:'02/01/2026', vendor:'', items:[{name:'Item',size:'',qty:1,unit:'Pcs',reason:'x'}], totalQty:1}
            ];
            App.State.filteredWastage = App.State.globalWastage;
            App.Wastage.renderTable();
            document.getElementById('wastageReviewSection').style.display = 'block';
        """,
        "checkbox": ".wastage-select-chk", "button": "#btnBulkDeleteWastage",
    },
    {
        "name": "Issue",
        "nav": "document.getElementById('btn-productionTab').click();",
        "seed": """
            App.State.globalIssues = [
              {issueId:'SWP-I-1', date:'01/01/2026', issuedTo:'Dept A', reference:'', items:[{name:'Item',size:'',qty:1,unit:'Pcs'}], totalQty:1},
              {issueId:'SWP-I-2', date:'02/01/2026', issuedTo:'Dept B', reference:'', items:[{name:'Item',size:'',qty:1,unit:'Pcs'}], totalQty:1}
            ];
            App.State.filteredIssues = App.State.globalIssues;
            App.Issue.renderTable();
            App.Production.switchSubTab('issuedStockSubTab');
        """,
        "checkbox": ".issue-select-chk", "button": "#btnBulkDeleteIssue",
    },
    {
        "name": "Item",
        "nav": "document.getElementById('btn-itemMaster').click();",
        "seed": """
            App.State.globalItems = [
              {name:'Sweep Item A', size:'', remarks:'', narration:'', specification:'', baseUnit:'Pcs', purchaseUnit:'Pcs', weightPerBaseUnit:0, vendors:[]},
              {name:'Sweep Item B', size:'', remarks:'', narration:'', specification:'', baseUnit:'Pcs', purchaseUnit:'Pcs', weightPerBaseUnit:0, vendors:[]}
            ];
            App.State.filteredItems = App.State.globalItems;
            App.Item.renderTable();
        """,
        "checkbox": ".item-select-chk", "button": "#btnBulkDeleteItems",
    },
    {
        "name": "Vendor",
        "nav": "document.getElementById('btn-vendorMaster').click();",
        "seed": """
            App.State.globalVendors = [
              {name:'Sweep Vendor A', contact:'', address:'', gstin:'', remarks:''},
              {name:'Sweep Vendor B', contact:'', address:'', gstin:'', remarks:''}
            ];
            App.State.filteredVendors = App.State.globalVendors;
            App.Vendor.renderTable();
        """,
        "checkbox": ".vendor-select-chk", "button": "#btnBulkDeleteVendors",
    },
    {
        "name": "BOM",
        "nav": "document.getElementById('btn-productsTab').click(); App.Products.switchSubTab('bomTab');",
        "seed": """
            App.State.globalBOMs = [
              {productId:'SWP-PRD-1', productName:'Sweep Product A', components:[]},
              {productId:'SWP-PRD-2', productName:'Sweep Product B', components:[]}
            ];
            App.State.filteredBOMs = App.State.globalBOMs;
            App.BOM.renderTable();
        """,
        "checkbox": ".bom-select-chk", "button": "#btnBulkDeleteBOMs",
    },
    {
        "name": "Process",
        "nav": "document.getElementById('btn-productsTab').click(); App.Products.switchSubTab('processTab');",
        "seed": """
            App.State.globalProcesses = [
              {processId:'SWP-PRC-1', processName:'Sweep Process A', sequence:1, lotPrefix:'SPA', isFinalStage:false, active:true, remarks:'', outputItemName:'Out A', processType:'', primaryColorAxis:''},
              {processId:'SWP-PRC-2', processName:'Sweep Process B', sequence:2, lotPrefix:'SPB', isFinalStage:false, active:true, remarks:'', outputItemName:'Out B', processType:'', primaryColorAxis:''}
            ];
            App.State.filteredProcesses = App.State.globalProcesses;
            App.Process.renderTable();
        """,
        "checkbox": ".process-select-chk", "button": "#btnBulkDeleteProcesses",
    },
    {
        "name": "Contractor",
        "nav": "document.getElementById('btn-contractorsTab').click();",
        "seed": """
            App.State.globalContractors = [
              {name:'Sweep Contractor A', contact:'', address:'', gstPan:'', remarks:''},
              {name:'Sweep Contractor B', contact:'', address:'', gstPan:'', remarks:''}
            ];
            App.State.filteredContractors = App.State.globalContractors;
            App.Contractor.renderTable();
        """,
        "checkbox": ".contractor-select-chk", "button": "#btnBulkDeleteContractors",
    },
    {
        "name": "Production",
        "nav": "document.getElementById('btn-productionTab').click(); App.Production.switchSubTab('productionLogSubTab');",
        "seed": """
            App.State.globalProduction = [
              {rowIdx:101, date:'01/01/2026', dateRaw:null, processId:'PRC-1', productId:'', productName:'', qty:5, assignedBy:'', assignedTo:'C', status:'Pending', remarks:'', customComponents:[], sheetRemarks:'', lotNumber:'LOT-1', contractorPayable:0, outputItemName:'Out', componentsConsumed:[], color:'', colorBreakdown:[]},
              {rowIdx:102, date:'02/01/2026', dateRaw:null, processId:'PRC-1', productId:'', productName:'', qty:5, assignedBy:'', assignedTo:'C', status:'Pending', remarks:'', customComponents:[], sheetRemarks:'', lotNumber:'LOT-2', contractorPayable:0, outputItemName:'Out', componentsConsumed:[], color:'', colorBreakdown:[]}
            ];
            App.State.filteredProduction = App.State.globalProduction;
            App.Production.renderTable();
        """,
        "checkbox": ".production-select-chk", "button": "#btnBulkDeleteProduction",
    },
    {
        "name": "Client (Clients)",
        "nav": "document.getElementById('btn-clientsTab').click(); App.Client.switchSubTab('clientsSubTab');",
        "seed": """
            App.State.globalClients = [
              {name:'Sweep Client A', contact:'', address:'', gstin:'', remarks:''},
              {name:'Sweep Client B', contact:'', address:'', gstin:'', remarks:''}
            ];
            App.State.filteredClients = App.State.globalClients;
            App.Client.renderClientsTable();
        """,
        "checkbox": ".client-select-chk", "button": "#btnBulkDeleteClients",
    },
    {
        "name": "Client (Orders)",
        "nav": "document.getElementById('btn-clientsTab').click(); App.Client.switchSubTab('clientOrdersSubTab');",
        "seed": """
            App.State.globalOrders = [
              {orderNumber:'SWP-PI-1', orderDate:'01/01/2026', clientName:'Sweep Client A', status:'Estimate', orderRemarks:'', lines:[{productId:'P',productName:'Prod',qty:1,lineRemarks:''}]},
              {orderNumber:'SWP-PI-2', orderDate:'02/01/2026', clientName:'Sweep Client B', status:'Estimate', orderRemarks:'', lines:[{productId:'P',productName:'Prod',qty:1,lineRemarks:''}]}
            ];
            App.State.filteredOrders = App.State.globalOrders;
            App.Client.renderOrdersTable();
        """,
        "checkbox": ".order-select-chk", "button": "#btnBulkDeleteOrders",
    },
    {
        "name": "Dispatch",
        "nav": "document.getElementById('btn-dispatchTab').click(); App.Dispatch.switchSubTab('dispatchedGoodsSubTab');",
        "seed": """
            App.State.globalDispatch = [
              {rowIdx:201, dispatchNumber:'SWP-DSP-1', dispatchDate:'01/01/2026', orderNumber:'', clientName:'Client A', productId:'P', productName:'Prod', qty:1, transport:'', remarks:'', invoiceNumber:'', grNumber:'', logisticsContractor:'', logisticsRate:0, logisticsCost:0},
              {rowIdx:202, dispatchNumber:'SWP-DSP-2', dispatchDate:'02/01/2026', orderNumber:'', clientName:'Client B', productId:'P', productName:'Prod', qty:1, transport:'', remarks:'', invoiceNumber:'', grNumber:'', logisticsContractor:'', logisticsRate:0, logisticsCost:0}
            ];
            App.State.filteredDispatch = App.State.globalDispatch;
            App.Dispatch.renderDispatchTable();
        """,
        "checkbox": ".dispatch-select-chk", "button": "#btnBulkDeleteDispatch",
    },
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

        # Mock window.google.script.run (Api.call's transport) for the whole
        # script, BEFORE any tab navigation. `Api` itself is an unreachable
        # top-level `const` inside Script.html (not a window property, so it
        # can't be reassigned from page.evaluate — same quirk documented in
        # verify_p1_3_p1_4_fixes.py) but google.script.run IS reachable since
        # it's what google.script.run's real caller reads off `window`.
        # Without this, clicking a nav tab kicks off that module's real
        # loadData() -> Api.call -> google.script.run, whose handler can
        # resolve asynchronously AFTER our seeded state is set and clobber/
        # re-render the table out from under the test (this is what caused
        # the original flaky "only 1 checkbox" / mid-check timeout failures).
        # A Proxy that counts every call and never invokes the success
        # callback makes every module's own loadData() call permanently inert.
        page.evaluate("""
            window.__apiCallCount = 0;
            window.google = {
                script: {
                    run: {
                        withSuccessHandler(cb) {
                            const runner = { withFailureHandler() { return runner; } };
                            return new Proxy(runner, {
                                get(target, prop) {
                                    if (prop in target) return target[prop];
                                    return (...args) => { window.__apiCallCount++; };
                                }
                            });
                        }
                    }
                }
            };
        """)

        for mod in MODULES:
            print(f"\n[{mod['name']}] bulk-select 2 + Delete Selected -> confirmation must appear")
            page.evaluate(mod["nav"])
            page.wait_for_timeout(50)
            page.evaluate(mod["seed"])
            page.wait_for_timeout(50)

            count = page.locator(mod["checkbox"]).count()
            check(count >= 2, f"at least 2 checkboxes rendered (got {count})")
            if count < 2:
                continue

            # Driven via direct DOM calls (checked + a real 'change' event,
            # then .click() on the button), not Playwright's mouse-simulated
            # check()/click() — the previous module's confirm modal can
            # still be mid-fade-out here (Bootstrap's hide() transition
            # isn't synchronous with the JS-side cleanup below), which
            # physically intercepts a screen-coordinate click on an
            # otherwise-correct, visible checkbox. We care about resulting
            # application state here, not pixel-level click realism.
            page.evaluate(f"""
                window.__apiCallCount = 0;
                App.State.confirmCallback = null;
                const boxes = document.querySelectorAll('{mod["checkbox"]}');
                [boxes[0], boxes[1]].forEach(b => {{
                    b.checked = true;
                    b.dispatchEvent(new Event('change', {{ bubbles: true }}));
                }});
                document.querySelector('{mod["button"]}').click();
            """)
            page.wait_for_timeout(50)

            api_calls = page.evaluate("window.__apiCallCount")
            check(api_calls == 0, f"Api.call NOT invoked yet (delete is gated behind confirm) (got {api_calls} call(s))")

            has_callback = page.evaluate("typeof App.State.confirmCallback === 'function'")
            check(has_callback, "App.State.confirmCallback is set (confirmAction fired)")

            msg_text = page.evaluate("document.getElementById('confirmMessageText')?.innerText || ''")
            check("2" in msg_text, f"confirm message mentions the count of 2 selected (got: {msg_text!r})")

            # The modal is left open (we never confirm/cancel it as part of
            # the assertions above) — close it via its own real Cancel
            # button (data-bs-dismiss="modal") so Bootstrap's native
            # dismiss lifecycle handles teardown itself, rather than us
            # racing/fighting its internal instance state with manual
            # hide()/dispose()/classList calls (which left it out of sync
            # and threw stray "Cannot read properties of null" errors on
            # the next module's confirmAction call).
            page.evaluate("""
                document.querySelector('#confirmModal [data-bs-dismiss="modal"]').click();
            """)
            page.wait_for_timeout(400)

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
