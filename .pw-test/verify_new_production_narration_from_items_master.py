"""
Verifies the 2026-08-01 fix: creating a NEW Production lot from a Process
recipe now shows each component's LIVE Items Master narration, not the
possibly-stale narration snapshot frozen on the Process recipe row.

Reported bug: "old narration is still feeding into New Production" -- the
Process module stores its own narration on each recipe component
(PROCESS_COMPONENTS sheet), refreshed only when a human re-touches that
row's Item/Size picker in the Process editor. Editing an item's narration
in Items Master afterward does NOT cascade into already-saved recipe rows
(by design -- narration is meant to be resolved live at read time, the same
way the already-logged Production Sheet's print/export path already does
via _resolveDisplayNarration/_lookupSheetItem). But the three functions
that populate a NEW lot's Components table from a Process recipe
(populateComponentsFromProcess, populateCommonComponentsFromProcess,
populateColorMatrixForColors) were copying the recipe's raw stored
narration straight into the row, never calling that live-resolution
helper -- so a NEW lot silently showed stale text even though Items
Master had since been corrected.

Covers all three population paths:
  1. populateComponentsFromProcess -- process with NO configured colors.
  2. populateCommonComponentsFromProcess -- process WITH colors, the
     Common Components table.
  3. populateColorMatrixForColors -- process WITH colors, the Per-Color
     matrix's color-tagged-row path (addMergedMatrixRow's first narration
     arg). Its second call site, the commonOverrideComps fallback, uses the
     exact same _resolveDisplayNarration call but only actually creates a
     row (rather than filling an existing one) when no per-color sibling
     already claimed that display name -- a narrower case not exercised
     here.

Run: python .pw-test/verify_new_production_narration_from_items_master.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"

PROCESS_NO_COLOR = {
    "processId": "PRC-NC", "processName": "Packing (No Color)", "sequence": 3, "lotPrefix": "PACKNC",
    "outputItemName": "Packed Cycle", "isFinalStage": False, "active": True, "processType": "General"
}
PROCESS_WITH_COLOR = {
    "processId": "PRC-WC", "processName": "Packing (Color)", "sequence": 3, "lotPrefix": "PACKWC",
    "outputItemName": "Packed Cycle Color", "isFinalStage": False, "active": True, "processType": "General"
}

ITEMS = [
    {"name": "Carton Box", "size": "", "narration": "FRESH carton note", "baseUnit": "Pcs"},
    {"name": "Frame Sticker", "size": "", "narration": "FRESH sticker note", "baseUnit": "Pcs"},
]

# Recipe rows carry an intentionally STALE narration, distinct from Items
# Master's current value above -- this is the "frozen at recipe-save-time"
# snapshot the bug leaves on screen.
COMPONENTS_NO_COLOR = [
    {"itemName": "Carton Box", "size": "", "narration": "STALE carton note (recipe)", "qtyPerUnit": 1,
     "sourceType": "ITEM", "colorGroup": "", "unit": "Pcs"},
]
COMPONENTS_WITH_COLOR = [
    # Common (no colorGroup) -> populateCommonComponentsFromProcess
    {"itemName": "Carton Box", "size": "", "narration": "STALE carton note (recipe)", "qtyPerUnit": 1,
     "sourceType": "ITEM", "colorGroup": "", "unit": "Pcs"},
    # Color-tagged -> populateColorMatrixForColors' color-comps path
    {"itemName": "Frame Sticker", "size": "", "narration": "STALE sticker note (recipe)", "qtyPerUnit": 1,
     "sourceType": "ITEM", "colorGroup": "Red", "unit": "Pcs"},
]

MOCK_API_RESPONSES = {
    "getProcessData": {"success": True, "data": [PROCESS_NO_COLOR, PROCESS_WITH_COLOR]},
    "getColors": {"success": True, "data": [{"name": "Red"}]},
    "getProcessColorAxes": {"success": True, "data": {"axes": [], "primaryAxisKey": ""}},
    "getItemsData": {"success": True, "data": ITEMS},
    "getProductionData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
    "getIssuedStockData": {"success": True, "data": []},
    "getWarehousePoolData": {"success": True, "data": []},
    "getModels": {"success": True, "data": []},
    "getProcessTypes": {"success": True, "data": []},
}


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
            window.__componentsByProcess = {{
                'PRC-NC': {json.dumps(COMPONENTS_NO_COLOR)},
                'PRC-WC': {json.dumps(COMPONENTS_WITH_COLOR)}
            }};
            window.__colorGroupsByProcess = {{ 'PRC-NC': [], 'PRC-WC': ['Red'] }};
            window.google = {{ script: {{ run: {{
                withSuccessHandler(cb) {{
                    const runner = {{ withFailureHandler() {{ return runner; }} }};
                    Object.keys(window.__mockResponses).forEach(method => {{
                        runner[method] = () => setTimeout(() => cb(window.__mockResponses[method]), 20);
                    }});
                    runner.getProcessComponentsData = (processId) => setTimeout(
                        () => cb({{success: true, data: window.__componentsByProcess[processId] || []}}), 20);
                    runner.getProcessColorGroups = (processId) => setTimeout(
                        () => cb({{success: true, data: window.__colorGroupsByProcess[processId] || []}}), 20);
                    return runner;
                }}
            }} }} }};
        """)

        failures = []

        def check(cond, msg):
            print(("PASS: " if cond else "FAIL: ") + msg)
            if not cond:
                failures.append(msg)

        print("[Setup] Load the Production tab (populates globalItems/globalProcesses)...")
        page.evaluate("App.Navigation.showTab('productionTab')")
        page.wait_for_timeout(700)

        # The shell's own demo bootstrap may have already seeded
        # App.State.globalItems before the google.script.run stub above was
        # installed (App.Item.ensureLoaded() short-circuits once _loaded is
        # true) -- force our own ITEMS fixture in directly, the same
        # workaround verify_production_loads_items_master.py uses.
        page.evaluate("(items) => { App.State.globalItems = items; }", ITEMS)

        # ---- 1. populateComponentsFromProcess (no-color path) -------------
        print("\n[1] populateComponentsFromProcess (process with no configured colors)")
        page.evaluate("async () => { await App.Production.handleProcessChange('PRC-NC'); }")
        page.wait_for_timeout(300)
        narrations = page.evaluate("""
            Array.from(document.querySelectorAll('#productionComponentsBody .prod-comp-narration')).map(i => i.value)
        """)
        print("  narrations:", narrations)
        check(narrations == ["FRESH carton note"],
              f"Carton Box row shows Items Master's FRESH narration, not the recipe's stale one (got {narrations})")

        # ---- 2/3. populateCommonComponentsFromProcess + populateColorMatrixForColors (with-color path) --
        print("\n[2/3] Process WITH colors -> Common table + Per-Color matrix")
        page.evaluate("async () => { await App.Production.handleProcessChange('PRC-WC'); }")
        page.wait_for_timeout(300)
        # Drives populateColorMatrixForColors the same way
        # handleColorCheckToggle does (add the column, then populate it) --
        # calling it directly rather than through the full checkbox-toggle
        # cascade (auto-sync of non-primary rows, pool availability refresh,
        # etc.), which is unrelated pre-existing behavior this test isn't
        # trying to cover.
        page.evaluate("""async () => {
            App.Production.addMatrixColorColumn('Red');
            await App.Production.populateColorMatrixForColors('PRC-WC', ['Red'], App.Production._compLoadSeq, undefined);
        }""")
        page.wait_for_timeout(300)

        common_narrations = page.evaluate("""
            Array.from(document.querySelectorAll('#productionComponentsBody .prod-comp-narration')).map(i => i.value)
        """)
        print("  common table narrations:", common_narrations)
        check("FRESH carton note" in common_narrations,
              f"Common Components table shows FRESH narration for Carton Box (got {common_narrations})")
        check("STALE carton note (recipe)" not in common_narrations,
              "Common Components table does not show the stale recipe narration")

        matrix_narrations = page.evaluate("""
            Array.from(document.querySelectorAll('#productionColorMatrixBody .prod-comp-narration')).map(i => i.value)
        """)
        print("  matrix narrations:", matrix_narrations)
        check("FRESH sticker note" in matrix_narrations,
              f"Per-Color matrix shows FRESH narration for Frame Sticker (Red-tagged row) (got {matrix_narrations})")
        check(not any('STALE' in n for n in matrix_narrations),
              f"no stale recipe narration leaks into the Per-Color matrix (got {matrix_narrations})")

        if errors := console_errors:
            print("\n  Console/page errors:")
            for e in errors:
                print(f"    {e}")
            failures.append("console errors")

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
