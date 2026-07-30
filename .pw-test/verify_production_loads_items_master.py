"""
Verification: the Production tab loads Items Master itself, so component
narration and Base Unit render correctly even when the operator never
visited the Item Master tab.

Reported bug: "item's narration isn't getting refreshed, neither is updated
item's base unit". Root cause was NOT the server-side refresh pass (which
rewrites the stored JSON correctly) but the CLIENT: App.State.globalItems is
populated ONLY by App.Item.loadData(), i.e. only if the Item Master tab was
opened. Production resolves component narration and unit live against that
cache at render time (_resolveDisplayNarration / _resolveDisplayUnit), so on
a session that went straight to Production, every component silently showed
its stored narration and a blanket 'Pcs' fallback unit -- looking exactly
like "the refresh did nothing".

Fixed by adding App.Item.ensureLoaded() (Script_Items.html, mirroring
App.Process.ensureLoaded) and awaiting it in App.Production.loadData().

Covered:
  1. Opening Production WITHOUT ever visiting Item Master populates
     globalItems, and narration/unit resolve to real Items Master values
     (not '' / the 'Pcs' fallback).
  2. The Refresh button force-reloads Items Master (App.Item.loadData, not
     ensureLoaded) so edits made after the first load are picked up rather
     than served from the already-warm cache.

Run: python .pw-test/verify_production_loads_items_master.py
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
]

# 'Adhesive Tape' is deliberately NOT tracked in Pcs -- the blanket 'Pcs'
# fallback is what the bug produced, so a non-Pcs base unit is what proves
# the real lookup ran.
ITEMS_V1 = [
    {"name": "Carton Box", "size": "", "narration": "Corrugated 5-ply", "baseUnit": "Pcs"},
    {"name": "Adhesive Tape", "size": "", "narration": "Original tape note", "baseUnit": "Kg"},
]
ITEMS_V2 = [
    {"name": "Carton Box", "size": "", "narration": "Corrugated 5-ply", "baseUnit": "Pcs"},
    {"name": "Adhesive Tape", "size": "", "narration": "REVISED tape note", "baseUnit": "Set"},
]

MOCK_PRODUCTION = [
    {"rowIdx": 1, "date": "01/07/2026", "dateRaw": "2026-07-01", "lotNumber": "PACK-1", "processId": "PRC-1",
     "outputItemName": "Packed Cycle", "productId": "PT-1", "productName": "Model A", "qty": 10,
     "assignedBy": "Alice", "assignedTo": "Bob", "status": "Pending", "color": "", "colorBreakdown": []},
]


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_context().new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        # getItemsData is served from a swappable slot so we can simulate the
        # operator editing Items Master between the two loads.
        page.evaluate(f"""
            window.__itemsV1 = {json.dumps(ITEMS_V1)};
            window.__itemsV2 = {json.dumps(ITEMS_V2)};
            window.__itemsCurrent = window.__itemsV1;
            window.__itemsCallCount = 0;
            window.__mockResponses = {{
                "getProcessData": {{"success": true, "data": {json.dumps(MOCK_PROCESSES)}}},
                "getColors": {{"success": true, "data": []}},
                "getProductionData": {{"success": true, "data": {json.dumps(MOCK_PRODUCTION)}}},
                "getStockData": {{"success": true, "data": []}},
                "getIssuedStockData": {{"success": true, "data": []}},
                "refreshProductionComponentsFromItemsMaster": {{"success": true, "message": "Refreshed 2 field(s).", "data": {{"fieldsUpdated": 2}}}}
            }};
            window.google = {{ script: {{ run: {{
                withSuccessHandler(cb) {{
                    const runner = {{ withFailureHandler() {{ return runner; }} }};
                    Object.keys(window.__mockResponses).forEach(method => {{
                        runner[method] = () => setTimeout(() => cb(window.__mockResponses[method]), 20);
                    }});
                    runner.getItemsData = () => setTimeout(() => {{
                        window.__itemsCallCount++;
                        cb({{success: true, data: window.__itemsCurrent}});
                    }}, 20);
                    return runner;
                }}
            }} }} }};
        """)

        failures = []

        def check(cond, msg):
            print(("PASS: " if cond else "FAIL: ") + msg)
            if not cond:
                failures.append(msg)

        print("[Step 1] Go STRAIGHT to Production -- never open the Item Master tab...")
        # The shell's own bootstrap may have seeded globalItems before the stub
        # was installed; clear it so this genuinely reproduces "the operator
        # never opened Item Master this session".
        page.evaluate("() => { App.State.globalItems = []; }")
        globals_before = page.evaluate("(App.State.globalItems || []).length")
        check(globals_before == 0, f"globalItems starts empty (got {globals_before})")

        page.evaluate("App.Navigation.showTab('productionTab')")
        page.wait_for_timeout(900)

        loaded = page.evaluate("(App.State.globalItems || []).length")
        check(loaded == 2, f"Production's own load populated globalItems without visiting Item Master (got {loaded})")

        print("\n[Step 2] Narration and Base Unit resolve to real Items Master values...")
        resolved = page.evaluate("""
            ({
                narration: App.Production._resolveDisplayNarration('Adhesive Tape', '', 'STORED FALLBACK'),
                unit: App.Production._resolveDisplayUnit('Adhesive Tape', '')
            })
        """)
        print("  resolved:", resolved)
        check(resolved["narration"] == "Original tape note",
              f"narration resolves from Items Master, not the stored fallback (got '{resolved['narration']}')")
        check(resolved["unit"] == "Kg",
              f"unit resolves to the item's real Base Unit 'Kg', not the blanket 'Pcs' fallback (got '{resolved['unit']}')")

        print("\n[Step 3] Operator edits Items Master; the Refresh button must pick the new values up...")
        page.evaluate("window.__itemsCurrent = window.__itemsV2;")
        calls_before = page.evaluate("window.__itemsCallCount")

        # confirmAction gates the button; auto-accept so the action runs.
        page.evaluate("() => { App.Utils.confirmAction = (msg, cb) => cb(); }")
        page.evaluate("() => App.Production.refreshFromItemsMaster()")
        page.wait_for_timeout(1200)

        calls_after = page.evaluate("window.__itemsCallCount")
        check(calls_after > calls_before,
              f"Refresh re-fetched Items Master rather than reusing the warm cache (calls {calls_before} -> {calls_after})")

        resolved2 = page.evaluate("""
            ({
                narration: App.Production._resolveDisplayNarration('Adhesive Tape', '', 'STORED FALLBACK'),
                unit: App.Production._resolveDisplayUnit('Adhesive Tape', '')
            })
        """)
        print("  resolved after refresh:", resolved2)
        check(resolved2["narration"] == "REVISED tape note",
              f"narration now shows the EDITED Items Master value (got '{resolved2['narration']}')")
        check(resolved2["unit"] == "Set",
              f"unit now shows the EDITED Base Unit (got '{resolved2['unit']}')")

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
    sys.exit(0 if run() else 1)
