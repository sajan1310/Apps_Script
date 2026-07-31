"""
Verification script: Excel-style per-column funnel filters added to the
Processes table (Process Type, Output Item Name, Final Stage, Active),
mirroring the existing Production Log / Item Master pattern
(App.Production.toggleColumnFilter / App.Item.toggleColumnFilter).

Covers: opening a column's checklist dropdown, checking values narrows the
table (AND across columns, OR within one column's checked values), the
funnel icon gets an "active" class while a filter is set, "Clear" empties
just that column's filter, and the free-text search box still combines with
column filters instead of one overriding the other.

Run: python .pw-test/verify_process_column_filters.py
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
    {"processId": "PRC-1", "processName": "Packing 16 inch", "sequence": 3, "lotPrefix": "PACK16",
     "outputItemName": "16 inch Packed Cycle", "isFinalStage": True, "active": True, "processType": "Packing"},
    {"processId": "PRC-2", "processName": "Painting 16 inch", "sequence": 2, "lotPrefix": "PAINT16",
     "outputItemName": "16 inch Painted Frame", "isFinalStage": False, "active": True, "processType": "Painting"},
    {"processId": "PRC-3", "processName": "Painting 20 inch", "sequence": 2, "lotPrefix": "PAINT20",
     "outputItemName": "20 inch Painted Frame", "isFinalStage": False, "active": False, "processType": "Painting"},
]

MOCK_API_RESPONSES = {
    "getProcessData": {"success": True, "data": MOCK_PROCESSES},
    "getModels": {"success": True, "data": []},
    "getProcessTypes": {"success": True, "data": []},
}


def visible_process_ids(page):
    return page.evaluate("""
        Array.from(document.querySelectorAll('#processTableBody tr'))
             .map(tr => tr.querySelector('td:nth-child(3) .badge')?.textContent.trim())
             .filter(Boolean)
    """)


def open_filter(page, key):
    """Opens the column's filter panel -- a no-op if it's already open for
    this same key (clicking the funnel a 2nd time toggles it CLOSED)."""
    existing_key = page.evaluate("document.getElementById('processColFilterPanel')?.dataset.key || null")
    if existing_key != key:
        page.locator(f'#processTab .th-filter-btn[data-filter-key="{key}"]').click()
        page.locator('#processColFilterPanel').wait_for(state="visible", timeout=TIMEOUT)
    page.wait_for_timeout(100)


def check_option(page, value):
    page.locator(f'#processColFilterPanel .po-ms-option[data-value="{value}"]').click()
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

        print("[Step 1] Load the Products & Processes tab...")
        page.evaluate("App.Navigation.showTab('productsTab')")
        page.wait_for_timeout(500)
        ids = visible_process_ids(page)
        print("  visible processes:", ids)
        check(sorted(ids) == ["PRC-1", "PRC-2", "PRC-3"], f"All 3 processes visible before any filter (got {ids})")

        print("\n[Step 2] Filter Process Type = Painting...")
        open_filter(page, "processType")
        check_option(page, "Painting")
        ids = visible_process_ids(page)
        print("  visible processes:", ids)
        check(sorted(ids) == ["PRC-2", "PRC-3"], f"Only the 2 Painting processes show (got {ids})")

        icon_active = page.evaluate("""
            document.querySelector('.th-filter-btn[data-filter-key="processType"]').classList.contains('active')
        """)
        check(icon_active, "Process Type funnel icon shows active state")

        print("\n[Step 3] AND with Active = Inactive (only PRC-3 qualifies)...")
        open_filter(page, "active")
        check_option(page, "Inactive")
        ids = visible_process_ids(page)
        print("  visible processes:", ids)
        check(ids == ["PRC-3"], f"Process Type=Painting AND Active=Inactive narrows to PRC-3 only (got {ids})")

        print("\n[Step 4] Clear the Active filter -- Process Type filter alone still applies...")
        open_filter(page, "active")
        page.locator('#processColFilterPanel [data-action="clear"]').click()
        page.wait_for_timeout(150)
        ids = visible_process_ids(page)
        print("  visible processes:", ids)
        check(sorted(ids) == ["PRC-2", "PRC-3"], f"Back to the 2 Painting processes after clearing Active (got {ids})")

        active_icon_active = page.evaluate("""
            document.querySelector('.th-filter-btn[data-filter-key="active"]').classList.contains('active')
        """)
        check(not active_icon_active, "Active funnel icon no longer shows active state after Clear")

        print("\n[Step 5] Clear Process Type too, then filter Final Stage = Yes...")
        open_filter(page, "processType")
        page.locator('#processColFilterPanel [data-action="clear"]').click()
        page.wait_for_timeout(150)
        open_filter(page, "finalStage")
        check_option(page, "Yes")
        ids = visible_process_ids(page)
        print("  visible processes:", ids)
        check(ids == ["PRC-1"], f"Final Stage=Yes shows only PRC-1 (got {ids})")

        print("\n[Step 6] Free-text search still combines with the active column filter...")
        open_filter(page, "finalStage")
        page.locator('#processColFilterPanel [data-action="clear"]').click()
        page.wait_for_timeout(150)
        open_filter(page, "outputItem")
        check_option(page, "16 inch Painted Frame")
        # The search box is wired via onkeyup -- type the real key events.
        page.locator('#searchProcess').press_sequentially('Painting')
        page.wait_for_timeout(150)
        ids = visible_process_ids(page)
        print("  visible processes:", ids)
        check(ids == ["PRC-2"], f"Search 'Painting' AND Output Item filter narrows to PRC-2 only (got {ids})")

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
