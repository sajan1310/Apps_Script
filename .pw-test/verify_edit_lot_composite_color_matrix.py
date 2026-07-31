"""
Regression test: reopening an EXISTING saved lot for Edit when the lot's
own color is a COMPOSITE built from two paired Color Axes (e.g.
"Blue-White / BCP") silently emptied the Per-Color Components table --
every one of its color-tagged components fell into the Common Components
table instead, so the matrix looked like it never loaded at all.

Root cause: populateComponentsConsumedDirect (Script_Production.html)
matched a saved component's colorGroup (a single axis token, e.g.
"Blue-White") against the lot's own breakdown colors with a plain string
comparison. A saved component is tagged with just ONE axis's own color,
while the lot's checked/breakdown color is the full composite string
("Blue-White / BCP") -- so the match never hit. The CREATE/change-process
path (populateColorMatrixForColors) already handles this correctly via
_matchedColorToken (fixed for the exact same symptom, per that function's
own doc comment); populateComponentsConsumedDirect never got the same fix.

Fixed by routing populateComponentsConsumedDirect's own colorGroup-match
through _matchedColorToken too, filling every composite color a token
matches (mirroring populateColorMatrixForColors's "both tokens legitimately
apply at once" rule).

Run: python .pw-test/verify_edit_lot_composite_color_matrix.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

MOCK_PROCESS = {
    "processId": "PRC-FIT", "processName": "Fitting Frame", "sequence": 5,
    "lotPrefix": "FIT", "outputItemName": "Fitted Frame Assembled", "isFinalStage": False,
    "active": True, "processType": "General", "primaryColorAxis": ""
}
# Explicitly color-tagged (not pool), one literal item per axis token --
# the "Frame Sticker---Blue-White" pattern populateComponentsConsumedDirect's
# own comments describe.
MOCK_COMPONENTS = [
    {"itemName": "Frame Sticker---Blue-White", "size": "", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "Blue-White"},
    {"itemName": "Frame Sticker---Orange", "size": "", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "Orange"},
]
MOCK_AXES = {
    "success": True,
    "data": {
        "axes": [
            {"key": "frame-axis", "label": "Frame Color", "colors": ["Blue-White", "Orange"], "source": "recipe"},
            {"key": "rim-axis", "label": "Rim Color", "colors": ["BCP", "Gold"], "source": "recipe"}
        ],
        "primaryColorAxis": "Frame Color",
        "primaryAxisKey": "frame-axis"
    }
}
# This lot's own output color is a COMPOSITE of the two axes above -- saved
# BEFORE this fix, so its componentsConsumed rows are tagged with just their
# own single axis token ("Blue-White"), never the full composite string.
SAVED_LOT = {
    "lotNumber": "FIT-0001", "processId": "PRC-FIT", "status": "Completed", "rowIdx": 3,
    "qty": 10, "color": "Blue-White / BCP", "assignedTo": "Ravi",
    "colorBreakdown": [{"color": "Blue-White / BCP", "qty": 10}],
    "componentsConsumed": [
        {"itemName": "Frame Sticker---Blue-White", "size": "", "sourceType": "ITEM", "qty": 10, "colorGroup": "Blue-White"},
    ],
}
MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": ["Blue-White", "Orange", "BCP", "Gold"]},
    "getProcessColorAxes": MOCK_AXES,
    "getProcessComponentsData": {"success": True, "data": MOCK_COMPONENTS},
    "getWarehousePoolData": {"success": True, "data": []},
    "getProcessWipData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
}


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_context().new_page()
        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.resolve().as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)
        page.evaluate(f"""
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
            App.State.globalItems = [];
            App.State.globalColors = [];
            App.State.globalContractors = [{{ contractorId: 'C1', name: 'Ravi', active: true }}];
            App.State.globalProduction = [{json.dumps(SAVED_LOT)}];
            window.__mockResponses = {json.dumps(MOCK_API_RESPONSES)};
            window.google = {{
                script: {{
                    run: {{
                        withSuccessHandler(cb) {{
                            const runner = {{ withFailureHandler() {{ return runner; }} }};
                            Object.keys(window.__mockResponses).forEach(method => {{
                                runner[method] = (...args) => setTimeout(() => cb(window.__mockResponses[method]), 20);
                            }});
                            return runner;
                        }}
                    }}
                }}
            }};
        """)

        ok = True
        print("[1] Open Edit modal for a lot whose color is a COMPOSITE ('Blue-White / BCP')...")
        page.evaluate("App.Production.openEditModal(0)")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(800)

        matrix_rows = page.evaluate("""
            Array.from(document.querySelectorAll('#productionColorMatrixBody tr')).map(tr => ({
                name: tr.querySelector('.prod-comp-display-name')?.value,
                narration: tr.querySelector('.prod-comp-narration')?.value,
            }))
        """)
        print(f"  Per-Color matrix rows: {matrix_rows}")
        if len(matrix_rows) == 0:
            print("  FAIL: Per-Color Components matrix is empty -- the exact reported bug")
            ok = False
        elif not any(r["name"] == "Frame Sticker" for r in matrix_rows):
            print(f"  FAIL: expected a 'Frame Sticker' row in the matrix, got {matrix_rows}")
            ok = False
        else:
            print("  PASS: Frame Sticker landed in the Per-Color matrix, not Common Components")

        common_rows = page.evaluate("""
            Array.from(document.querySelectorAll('#productionComponentsBody tr'))
                .map(tr => tr.querySelector('.prod-comp-item-select')?.value || tr.querySelector('select')?.value)
        """)
        print(f"  Common Components rows: {common_rows}")

        cell_value = page.evaluate("""
            (() => {
                const row = Array.from(document.querySelectorAll('#productionColorMatrixBody tr'))
                    .find(tr => tr.querySelector('.prod-comp-display-name')?.value === 'Frame Sticker');
                if (!row) return null;
                const th = Array.from(document.querySelectorAll('#productionColorMatrixHeaderRow th[data-color]'))
                    .find(t => t.dataset.color === 'Blue-White / BCP');
                if (!th) return { error: 'no Blue-White / BCP column found' };
                const idx = Array.from(th.parentElement.children).indexOf(th);
                const cell = row.children[idx];
                return { qty: cell?.querySelector('.matrix-qty')?.value };
            })()
        """)
        print(f"  'Blue-White / BCP' column cell: {cell_value}")
        if not cell_value or cell_value.get("qty") != "10":
            print(f"  FAIL: expected qty 10 in the 'Blue-White / BCP' column, got {cell_value}")
            ok = False
        else:
            print("  PASS: the saved qty (10) landed in the lot's own composite color column")

        if console_errors:
            print("\n  Console/page errors:")
            for e in console_errors:
                print(f"    {e}")
            ok = False

        browser.close()
        return ok


if __name__ == "__main__":
    ok = run()
    print("\n" + ("ALL PASS" if ok else "SOME FAILED"))
    sys.exit(0 if ok else 1)
