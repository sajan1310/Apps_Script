"""
Verification script: Excel-style per-column funnel filters added to the
Production Log table (Process/Lot #, Output Item, Product Tag, Assigned By,
Assigned To, Status), mirroring the existing Item Master pattern
(Script_Items.html's toggleColumnFilter/applyColumnFilters).

Covers: opening a column's checklist dropdown, checking values narrows the
table (AND across columns, OR within one column's checked values), the
funnel icon gets an "active" class while a filter is set, "Clear" empties
just that column's filter, and the free-text search box still combines with
column filters instead of one overriding the other.

Run: python .pw-test/verify_production_column_filters.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

MOCK_PROCESSES = [
    {"processId": "PRC-1", "processName": "Packing", "sequence": 3, "lotPrefix": "PACK",
     "outputItemName": "Packed Cycle", "isFinalStage": True, "active": True, "processType": "General"},
    {"processId": "PRC-2", "processName": "Painting", "sequence": 2, "lotPrefix": "PAINT",
     "outputItemName": "Painted Frame", "isFinalStage": False, "active": True, "processType": "General"},
]

MOCK_PRODUCTION = [
    {"rowIdx": 1, "date": "01/07/2026", "dateRaw": "2026-07-01", "lotNumber": "PACK-1", "processId": "PRC-1",
     "outputItemName": "Packed Cycle A", "productId": "PT-1", "productName": "Model A", "qty": 10,
     "assignedBy": "Alice", "assignedTo": "Bob", "status": "Pending", "color": "", "colorBreakdown": []},
    {"rowIdx": 2, "date": "02/07/2026", "dateRaw": "2026-07-02", "lotNumber": "PACK-2", "processId": "PRC-1",
     "outputItemName": "Packed Cycle B", "productId": "PT-2", "productName": "Model B", "qty": 20,
     "assignedBy": "Alice", "assignedTo": "Carol", "status": "Completed", "color": "", "colorBreakdown": []},
    {"rowIdx": 3, "date": "03/07/2026", "dateRaw": "2026-07-03", "lotNumber": "PAINT-1", "processId": "PRC-2",
     "outputItemName": "Painted Frame", "productId": "PT-1", "productName": "Model A", "qty": 15,
     "assignedBy": "Dave", "assignedTo": "Bob", "status": "Pending", "color": "", "colorBreakdown": []},
]

MOCK_API_RESPONSES = {
    "getProcessData": {"success": True, "data": MOCK_PROCESSES},
    "getColors": {"success": True, "data": []},
    "getProductionData": {"success": True, "data": MOCK_PRODUCTION},
    "getIssuedStockData": {"success": True, "data": []},
}


def visible_lot_numbers(page):
    return page.evaluate("""
        Array.from(document.querySelectorAll('#productionTableBody tr'))
             .map(tr => tr.querySelector('td:nth-child(3) .badge')?.textContent.trim())
             .filter(Boolean)
    """)


def open_filter(page, key):
    """Opens the column's filter panel — a no-op if it's already open for
    this same key (clicking the funnel a 2nd time toggles it CLOSED, same
    as Item Master's identical toggleColumnFilter)."""
    existing_key = page.evaluate("document.getElementById('productionColFilterPanel')?.dataset.key || null")
    if existing_key != key:
        page.locator(f'.th-filter-btn[data-filter-key="{key}"]').click()
        page.locator('#productionColFilterPanel').wait_for(state="visible", timeout=TIMEOUT)
    page.wait_for_timeout(100)


def check_option(page, value):
    page.locator(f'#productionColFilterPanel .po-ms-option[data-value="{value}"]').click()
    page.wait_for_timeout(150)


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context()
        page = ctx.new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        page.evaluate(f"""
            window.__mockResponses = {json.dumps(MOCK_API_RESPONSES)};
            window.google = {{ script: {{ run: {{
                withSuccessHandler(cb) {{
                    const runner = {{ withFailureHandler() {{ return runner; }} }};
                    Object.keys(window.__mockResponses).forEach(method => {{
                        runner[method] = () => setTimeout(() => cb(window.__mockResponses[method]), 30);
                    }});
                    return runner;
                }}
            }} }} }};
        """)

        failures = []

        def check(cond, msg):
            print(("PASS: " if cond else "FAIL: ") + msg)
            if not cond:
                failures.append(msg)

        print("[Step 1] Load the Production tab...")
        page.evaluate("App.Navigation.showTab('productionTab')")
        page.wait_for_timeout(500)
        lots = visible_lot_numbers(page)
        print("  visible lots:", lots)
        check(sorted(lots) == ["PACK-1", "PACK-2", "PAINT-1"], f"All 3 lots visible before any filter (got {lots})")

        print("\n[Step 2] Filter Process = Packing...")
        open_filter(page, "process")
        check_option(page, "Packing")
        lots = visible_lot_numbers(page)
        print("  visible lots:", lots)
        check(sorted(lots) == ["PACK-1", "PACK-2"], f"Only Packing's 2 lots show (got {lots})")

        icon_active = page.evaluate("""
            document.querySelector('.th-filter-btn[data-filter-key="process"]').classList.contains('active')
        """)
        check(icon_active, "Process funnel icon shows active state")

        print("\n[Step 3] AND with Status = Completed (only PACK-2 qualifies)...")
        open_filter(page, "status")
        check_option(page, "Completed")
        lots = visible_lot_numbers(page)
        print("  visible lots:", lots)
        check(lots == ["PACK-2"], f"Process=Packing AND Status=Completed narrows to PACK-2 only (got {lots})")

        print("\n[Step 4] Clear the Status filter -- Process filter alone still applies...")
        open_filter(page, "status")
        page.locator('#productionColFilterPanel [data-action="clear"]').click()
        page.wait_for_timeout(150)
        lots = visible_lot_numbers(page)
        print("  visible lots:", lots)
        check(sorted(lots) == ["PACK-1", "PACK-2"], f"Back to Packing's 2 lots after clearing Status (got {lots})")

        status_icon_active = page.evaluate("""
            document.querySelector('.th-filter-btn[data-filter-key="status"]').classList.contains('active')
        """)
        check(not status_icon_active, "Status funnel icon no longer shows active state after Clear")

        print("\n[Step 5] Clear Process too, then filter Product Tag = PT-1 (spans 2 different processes)...")
        open_filter(page, "process")
        page.locator('#productionColFilterPanel [data-action="clear"]').click()
        page.wait_for_timeout(150)
        open_filter(page, "productTag")
        check_option(page, "PT-1")
        lots = visible_lot_numbers(page)
        print("  visible lots:", lots)
        check(sorted(lots) == ["PACK-1", "PAINT-1"], f"Product Tag PT-1 shows both its lots regardless of process (got {lots})")

        print("\n[Step 6] Free-text search still combines with the active column filter...")
        # The search box is wired via onkeyup (not oninput), so .fill() alone
        # (which only dispatches input/change) wouldn't trigger it -- type
        # the real key events, same as an operator typing.
        page.locator('#searchProduction').press_sequentially('Painted')
        page.wait_for_timeout(150)
        lots = visible_lot_numbers(page)
        print("  visible lots:", lots)
        check(lots == ["PAINT-1"], f"Search 'Painted' AND Product Tag PT-1 narrows to PAINT-1 only (got {lots})")

        check(len(console_errors) == 0, f"No uncaught page errors (got {console_errors})")

        print("\n" + "=" * 60)
        if failures:
            print(f"FAILURES: {len(failures)}")
            for f in failures:
                print("  -", f)
        else:
            print("ALL CHECKS PASSED")
        print("=" * 60)

        browser.close()
        return len(failures) == 0


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
