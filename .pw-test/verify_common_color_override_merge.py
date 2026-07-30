"""
Verification script: Production Lot form's Per-Color Components merge for
"Common item that's really a per-color sibling in disguise".

Reproduces the reported bug: a recipe has one Common-tagged component
(e.g. "CRYSTA STICKER Chain Cover", applies to every color) and one
explicitly color-tagged sibling for a single color (e.g. "CRYSTA STICKER
Chain Cover Green", colorGroup Green) that's really the SAME physical item.
Before the fix, these rendered as two separate, seemingly-unrelated rows —
one flat in Common Components, one alone in the Per-Color matrix. After the
fix they must collapse into ONE merged Per-Color row with a genuine
per-color quantity for every color, and nothing left behind in Common
Components for that item.

Also verifies the same merge applies when reopening a legacy-saved lot
(populateComponentsConsumedDirect) and when viewing/printing an existing
lot's Production Sheet (groupComponentsForSheet).

Run: python .pw-test/verify_common_color_override_merge.py
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
    "processId": "PRC-1", "processName": "Frame Painting", "sequence": 1,
    "lotPrefix": "FP", "outputItemName": "Painted Frame", "isFinalStage": False,
    "active": True, "processType": "General"
}
MOCK_COMPONENTS = [
    {"itemName": "Brush", "size": "", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "COMMON"},
    {"itemName": "CRYSTA STICKER Chain Cover", "size": "General", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "COMMON"},
    {"itemName": "CRYSTA STICKER Chain Cover Green", "size": "General", "sourceType": "ITEM", "qtyPerUnit": 2, "colorGroup": "Green"},
]
MOCK_ITEMS = [{"name": c["itemName"], "size": c["size"]} for c in MOCK_COMPONENTS]
MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": ["Blue", "Green"]},
    "getProcessComponentsData": {"success": True, "data": MOCK_COMPONENTS},
    "getWarehousePoolData": {"success": True, "data": []},
    "getProcessWipData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
}


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_context(viewport={"width": 1400, "height": 900}).new_page()
        page.goto(DIST_HTML.resolve().as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        page.evaluate(f"""
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
            App.State.globalItems = {json.dumps(MOCK_ITEMS)};
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

        failures = []
        def check(cond, msg):
            print(("PASS: " if cond else "FAIL: ") + msg)
            if not cond:
                failures.append(msg)

        # ─────────────────────────────────────────────────────────────
        print("\n[Scenario 1] Fresh Create form: check both colors, verify merge...")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.evaluate("document.getElementById('productionProcessId').value='PRC-1'; App.Production.handleProcessChange('PRC-1');")
        page.wait_for_timeout(800)

        page.evaluate("document.querySelector('#productionColorChecklist .production-color-row[data-color=\"Blue\"] .production-color-check').click()")
        page.wait_for_timeout(300)
        page.locator("#productionColorChecklist .production-color-row[data-color='Blue'] .production-color-qty").fill("10")
        page.wait_for_timeout(300)
        page.evaluate("document.querySelector('#productionColorChecklist .production-color-row[data-color=\"Green\"] .production-color-check').click()")
        page.wait_for_timeout(300)
        page.locator("#productionColorChecklist .production-color-row[data-color='Green'] .production-color-qty").fill("4")
        page.wait_for_timeout(600)

        common_names = page.evaluate("""
            Array.from(document.querySelectorAll('#productionComponentsBody tr .prod-comp-item-select')).map(sel => {
                const opt = sel.options[sel.selectedIndex];
                return opt ? (opt.dataset.name || opt.textContent) : null;
            })
        """)
        check("CRYSTA STICKER Chain Cover" not in common_names, f"Common table no longer has the duplicate row (got {common_names})")

        matrix_rows = page.evaluate("""
            Array.from(document.querySelectorAll('#productionColorMatrixBody tr')).map(row => ({
                name: row.querySelector('.prod-comp-display-name')?.value,
                cellVals: Array.from(row.querySelectorAll('.matrix-qty')).map(i => i.value)
            }))
        """)
        chain_cover_rows = [r for r in matrix_rows if r["name"] == "CRYSTA STICKER Chain Cover"]
        check(len(chain_cover_rows) == 1, f"Exactly one merged matrix row for the item (got {len(chain_cover_rows)})")
        if chain_cover_rows:
            check(chain_cover_rows[0]["cellVals"] == ["10", "8"], f"Blue=10 (fallback qtyPerUnit*10), Green=8 (explicit qtyPerUnit 2*4) (got {chain_cover_rows[0]['cellVals']})")

        saved = json.loads(page.evaluate("""
            JSON.stringify([
                ...App.Production.serializeComponentsConsumed(),
                ...App.Production.serializeColorMatrix(),
                ...App.Production.serializePoolColorGroups()
            ])
        """))
        chain_saved = [c for c in saved if "Chain Cover" in c["itemName"]]
        check(len(chain_saved) == 2 and all(c["colorGroup"] != "COMMON" for c in chain_saved),
              f"Saved payload has 2 per-color entries, no combined COMMON row for this item (got {chain_saved})")

        # ─────────────────────────────────────────────────────────────
        print("\n[Scenario 2] Reopening a legacy-saved lot upgrades the same way...")
        legacy_result = page.evaluate("""
            (async () => {
                document.getElementById('productionColorMatrixBody').innerHTML = '';
                document.querySelectorAll('#productionColorMatrixHeaderRow th[data-color]').forEach(th => th.remove());
                // populateComponentsConsumedDirect takes a colorBreakdown —
                // {color, qty} objects, not bare color-name strings. This used to
                // pass ['Blue','Green'], so every entry.color was undefined and
                // the call threw before asserting anything.
                const breakdown = [{ color: 'Blue', qty: 6 }, { color: 'Green', qty: 8 }];
                const colors = breakdown.map(b => b.color);
                colors.forEach(c => App.Production.addMatrixColorColumn(c));
                const legacyComponents = [
                    { itemName: 'Brush', size: '', sourceType: 'ITEM', qty: 14, colorGroup: 'COMMON' },
                    { itemName: 'CRYSTA STICKER Chain Cover', size: 'General', sourceType: 'ITEM', qty: 14, colorGroup: 'COMMON' },
                    { itemName: 'CRYSTA STICKER Chain Cover Green', size: 'General', sourceType: 'ITEM', qty: 8, colorGroup: 'Green' }
                ];
                document.getElementById('productionComponentsBody').innerHTML = '';
                await App.Production.populateComponentsConsumedDirect(legacyComponents, breakdown);
                const commonRows = Array.from(document.querySelectorAll('#productionComponentsBody tr .prod-comp-item-select')).map(sel => {
                    const opt = sel.options[sel.selectedIndex];
                    return opt ? (opt.dataset.name || opt.textContent) : null;
                });
                const matrixRows = Array.from(document.querySelectorAll('#productionColorMatrixBody tr')).map(row => ({
                    name: row.querySelector('.prod-comp-display-name')?.value,
                    cellVals: Array.from(row.querySelectorAll('.matrix-qty')).map(i => i.value)
                }));
                return { commonRows, matrixRows };
            })()
        """)
        check("CRYSTA STICKER Chain Cover" not in legacy_result["commonRows"], f"Legacy reopen: item removed from Common table (got {legacy_result['commonRows']})")
        legacy_chain = [r for r in legacy_result["matrixRows"] if r["name"] == "CRYSTA STICKER Chain Cover"]
        check(len(legacy_chain) == 1 and legacy_chain[0]["cellVals"] == ["14", "8"],
              f"Legacy reopen: one merged row, Blue gets full fallback qty 14, Green keeps its saved 8 (got {legacy_chain})")

        # ─────────────────────────────────────────────────────────────
        print("\n[Scenario 3] Production Sheet grouping merges the same way for viewing/printing...")
        sheet_result = page.evaluate("""
            App.Production.groupComponentsForSheet([
                { itemName: 'Brush', size: '', sourceType: 'ITEM', qty: 14, colorGroup: 'COMMON' },
                { itemName: 'CRYSTA STICKER Chain Cover', size: 'General', sourceType: 'ITEM', qty: 14, colorGroup: 'COMMON' },
                { itemName: 'CRYSTA STICKER Chain Cover Green', size: 'General', sourceType: 'ITEM', qty: 8, colorGroup: 'Green' }
            ], ['Blue', 'Green'])
        """)
        common_item_names = [c["itemName"] for c in sheet_result["common"]]
        check("CRYSTA STICKER Chain Cover" not in common_item_names, f"Sheet: item not left in flat Common list (got {common_item_names})")
        sheet_slot = next((s for s in sheet_result["matrixSlots"] if s["itemName"] == "CRYSTA STICKER Chain Cover"), None)
        check(sheet_slot is not None and sheet_slot["colors"] == {"Green": 8, "Blue": 14},
              f"Sheet: one slot with Green=8 (saved) and Blue=14 (fallback split) (got {sheet_slot})")

        # ─────────────────────────────────────────────────────────────
        # A colorBreakdown entry with no color name must not take down the whole
        # Edit-Lot reopen. populateComponentsConsumedDirect used to do
        # `entry.color.toLowerCase()`, which threw on exactly this shape; the
        # caller that builds breakdownColors already guards with filter(Boolean),
        # so blank colors are a known possibility on this data.
        print("\n[Scenario 4] A breakdown entry with a missing color must not crash the reopen...")
        malformed = page.evaluate("""
            (async () => {
                document.getElementById('productionColorMatrixBody').innerHTML = '';
                document.querySelectorAll('#productionColorMatrixHeaderRow th[data-color]').forEach(th => th.remove());
                document.getElementById('productionComponentsBody').innerHTML = '';
                const breakdown = [
                    { color: 'Blue', qty: 6 },
                    { color: '', qty: 3 },          // blank name
                    { qty: 4 },                     // missing key entirely
                    { color: null, qty: 2 },        // explicit null
                    { color: 'Green', qty: 8 }
                ];
                const components = [
                    { itemName: 'Brush', size: '', sourceType: 'ITEM', qty: 14, colorGroup: 'COMMON' },
                    { itemName: 'CRYSTA STICKER Chain Cover Green', size: 'General', sourceType: 'ITEM', qty: 8, colorGroup: 'Green' }
                ];
                try {
                    await App.Production.populateComponentsConsumedDirect(components, breakdown);
                } catch (err) {
                    return { threw: String(err && err.message || err) };
                }
                return {
                    threw: null,
                    // Only the two real colors may become matrix columns -- a blank
                    // must never render as an 'undefined' column header.
                    columns: Array.from(
                        document.querySelectorAll('#productionColorMatrixHeaderRow th[data-color]')
                    ).map(th => th.dataset.color)
                };
            })()
        """)
        check(malformed["threw"] is None,
              f"Reopen survives a malformed breakdown entry (threw: {malformed['threw']})")
        if malformed["threw"] is None:
            check(malformed["columns"] == ["Blue", "Green"],
                  f"Only the named colors become matrix columns, no blank/undefined column (got {malformed['columns']})")

        print("\n" + ("ALL TESTS PASSED" if not failures else f"{len(failures)} TEST(S) FAILED"))
        browser.close()
        return len(failures) == 0


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
