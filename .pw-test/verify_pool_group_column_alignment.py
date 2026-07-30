"""
Regression test for a real bug found via screenshot inspection: when
columns are added one at a time to a Per-Process Pool Components table
(see syncPoolColorGroupColumns in Script.html), each new <td> was being
appended AFTER the row's trailing "✕" remove-row button instead of before
it — because the insertion index was read from the header row AFTER the
new <th> had already been inserted (shifting its own position), while the
<tr> being patched still had its OLD (pre-insertion) structure. Every
value ended up one column to the right of its own header, with the ✕
button landing in the first color's cell.

This test checks actual VALUE-TO-HEADER alignment (by data-color on each
cell's own input), not just header order, which a prior test missed.

Run: python .pw-test/verify_pool_group_column_alignment.py
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
    "processId": "PRC-FTF", "processName": "Fitting Frame", "sequence": 5,
    "lotPrefix": "FTF", "outputItemName": "Fitted Frame Assembled", "isFinalStage": False,
    "active": True, "processType": "General", "primaryColorAxis": "Painted Frame Crysta 20 inch D/Gaddi"
}
MOCK_COMPONENTS = [
    {"itemName": "Painted Frame Crysta 20 inch D/Gaddi", "size": "", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
    {"itemName": "Fitted Rim 20 inch", "size": "", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
    {"itemName": "20 inch Mudguard", "size": "", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON"},
]
PAINTED_COLORS = ["Blue-White", "Orange-White", "Pink-White", "Purple-White", "Red-White", "Sea Green-White"]
RIM_COLORS = ["BCP", "Black"]
MUDGUARD_COLORS = ["Blue", "Orange", "Pink", "Purple", "Red", "Sea Green"]

MOCK_POOL = (
    [{"outputItemName": "Painted Frame Crysta 20 inch D/Gaddi", "color": c, "processId": "PRC-P", "qty": 10} for c in PAINTED_COLORS]
    + [{"outputItemName": "Fitted Rim 20 inch", "color": c, "processId": "PRC-R", "qty": 10} for c in RIM_COLORS]
    + [{"outputItemName": "20 inch Mudguard", "color": c, "processId": "PRC-M", "qty": 10} for c in MUDGUARD_COLORS]
)
MOCK_AXES = {
    "success": True,
    "data": {
        "axes": [
            {"key": "pool:painted frame crysta 20 inch d/gaddi", "label": "Painted Frame Crysta 20 inch D/Gaddi", "colors": PAINTED_COLORS, "source": "pool"},
            {"key": "pool:fitted rim 20 inch", "label": "Fitted Rim 20 inch", "colors": RIM_COLORS, "source": "pool"},
            {"key": "pool:20 inch mudguard", "label": "20 inch Mudguard", "colors": MUDGUARD_COLORS, "source": "pool"}
        ],
        "primaryColorAxis": "Painted Frame Crysta 20 inch D/Gaddi",
        "primaryAxisKey": "pool:painted frame crysta 20 inch d/gaddi"
    }
}
MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": sorted(list({r["color"] for r in MOCK_POOL}))},
    "getProcessColorAxes": MOCK_AXES,
    "getProcessComponentsData": {"success": True, "data": MOCK_COMPONENTS},
    "getWarehousePoolData": {"success": True, "data": MOCK_POOL},
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
        print("[1] Open Create modal, select Fitting Frame...")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.evaluate("""
            document.getElementById('productionProcessId').value = 'PRC-FTF';
            App.Production.handleProcessChange('PRC-FTF');
        """)
        page.wait_for_timeout(800)

        print("[2] Check colors one at a time, mirroring the exact reported sequence "
              "(all 6 Painted Frame colors, then BCP, then all 6 Mudguard colors)...")
        checks = (
            [(c, str(v)) for c, v in zip(PAINTED_COLORS, [10, 7, 17, 7, 17, 10])]
            + [("BCP", "68")]
            + [(c, str(v)) for c, v in zip(MUDGUARD_COLORS, [10, 7, 17, 7, 17, 10])]
        )
        for color, qty in checks:
            page.evaluate(f"""
                const row = Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row'))
                    .find(r => r.dataset.color === {json.dumps(color)});
                const chk = row.querySelector('.production-color-check');
                chk.checked = true;
                App.Production.handleColorCheckToggle(chk);
                const qtyInput = row.querySelector('.production-color-qty');
                qtyInput.disabled = false;
                qtyInput.value = {json.dumps(qty)};
                qtyInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
            """)
            page.wait_for_timeout(60)
        page.wait_for_timeout(400)

        print("\n[3] Verify EVERY cell's own data-color attribute lines up under the matching header, "
              "and the row's remove (✕) button is the LAST cell, not the first...")
        report = page.evaluate("""
            Array.from(document.querySelectorAll('#productionPoolColorGroupsContainer table.prod-color-table')).map(table => {
                const headerCells = Array.from(table.querySelectorAll('thead tr th'));
                const headerColors = headerCells.map(th => th.dataset.color || null);
                const rows = Array.from(table.querySelectorAll('tbody tr')).map(tr => {
                    const cells = Array.from(tr.children);
                    return cells.map((td, i) => {
                        const input = td.querySelector('.pool-group-qty');
                        const btn = td.querySelector('button');
                        if (input) return { kind: 'value', color: input.dataset.color, value: input.value };
                        if (btn) return { kind: 'removeBtn' };
                        return { kind: 'other' };
                    });
                });
                return { allColors: table.dataset.allColors, headerColors, rows };
            })
        """)

        for entry in report:
            print(f"\n  Table: {entry['allColors']}")
            print(f"    Header order: {entry['headerColors']}")
            for row in entry['rows']:
                print(f"    Row cells:    {row}")
                # Header layout is [...descriptive columns..., ...colorHeaders..., "✕"]
                # Row layout must mirror it exactly:
                #   [...descriptive tds..., ...value-tds..., removeBtn-td]
                #
                # The number of leading descriptive columns is DERIVED from the
                # header row (the first header carrying a data-color starts the
                # colour block), not hardcoded. It used to be a literal 3, which
                # silently went stale the moment a Narration column was added to
                # these tables: the loop then started one column early, read a
                # descriptive cell as a colour cell and reported a misalignment
                # that wasn't there — the sort of false failure that trains you
                # to ignore the suite.
                n = len(entry['headerColors'])
                if len(row) != n:
                    print(f"    FAIL: row has {len(row)} cells, header has {n}")
                    ok = False
                    continue
                # last cell must be the remove button
                if row[-1]['kind'] != 'removeBtn':
                    print(f"    FAIL: last cell is not the remove (✕) button (got {row[-1]})")
                    ok = False
                first_color = next((i for i, h in enumerate(entry['headerColors']) if h), None)
                if first_color is None:
                    print("    FAIL: no header carries a data-color, so no colour block to align against")
                    ok = False
                    continue
                # every value cell's own data-color must match its header's data-color
                mismatch = False
                for i in range(first_color, n - 1):
                    cell = row[i]
                    header_color = entry['headerColors'][i]
                    if cell['kind'] != 'value':
                        print(f"    FAIL: cell at index {i} (under header '{header_color}') is not a value input (got {cell})")
                        ok = False
                        mismatch = True
                        continue
                    if (cell['color'] or '').lower() != (header_color or '').lower():
                        print(f"    FAIL: cell at index {i} has data-color '{cell['color']}' but sits under header '{header_color}'")
                        ok = False
                        mismatch = True
                # Every descriptive cell before the colour block must NOT be a
                # value input — that's what actually pins the block's start, and
                # catches a colour column drifting left into the labels.
                for i in range(0, first_color):
                    if row[i]['kind'] == 'value':
                        print(f"    FAIL: cell at index {i} is a value input but sits under a descriptive header")
                        ok = False
                        mismatch = True
                if not mismatch and row[-1]['kind'] == 'removeBtn':
                    print(f"    PASS: all {n - 1 - first_color} value cells aligned correctly under their headers "
                          f"({first_color} descriptive columns first), remove button last")

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
