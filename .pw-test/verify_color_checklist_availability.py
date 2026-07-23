"""
Verification script: the Production Lot form's "Colors to Produce" checklist
must show, next to every color, how much of that color is actually available
to build from right now — the Warehouse Pool balance of the upstream item(s)
that drive that group (or plain Stock for an ITEM-sourced tag-axis input).

Covers the real behaviors of _poolAvailByItemColor / _poolAvailForColor /
_groupInputSources / _colorRowAvailability / refreshColorChecklistAvailability
in Script_Production.html:
  * a pool-driven group's colors read that pool item's own color bucket
  * a group fed by SEVERAL pool items shows the smallest (the real ceiling),
    with the per-item breakdown in the tooltip
  * tagged (already committed to a packed product) pool rows are excluded,
    matching getPoolAvailableQtyMap server-side
  * a color with no exact bucket falls back to the composite buckets it is
    one token of (e.g. "Purple" <- "BCP / Purple")
  * zero / negative availability is flagged red+bold instead of blending in
  * an ITEM-sourced tag-axis color reads Stock, not the pool
  * an "Other Colors (Color Master)" row — no configured input at all —
    stays blank rather than asserting a "0 avail." it can't back up

Run: python .pw-test/verify_color_checklist_availability.py
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
    "processId": "PRC-1", "processName": "Fitted Frame Assembly", "sequence": 1,
    "lotPrefix": "FFA", "outputItemName": "Fitted Frame", "isFinalStage": False,
    "active": True, "processType": "General"
}

MOCK_COMPONENTS = [
    {"itemName": "Painted Frame", "size": "", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON", "colorAxis": ""},
    {"itemName": "Painted Mudguard", "size": "", "sourceType": "POOL", "qtyPerUnit": 1, "colorGroup": "COMMON", "colorAxis": ""},
    {"itemName": "Assembly Screws", "size": "", "sourceType": "ITEM", "qtyPerUnit": 4, "colorGroup": "COMMON", "colorAxis": ""},
    {"itemName": "Fitting Black", "size": "", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "Black", "colorAxis": "Fitting"},
]
MOCK_ITEMS = [{"name": c["itemName"], "size": c["size"]} for c in MOCK_COMPONENTS]

# One pool axis fed by BOTH pool items (label is the comma-joined item names,
# exactly what computeColorAxesForProcess emits) + one ITEM-backed tag axis.
MOCK_AXES = {
    "axes": [
        {"key": "pool:1", "label": "Painted Frame, Painted Mudguard",
         "colors": ["Blue", "Green", "Purple", "Red", "Yellow"], "source": "pool"},
        {"key": "tag:fitting", "label": "Fitting", "colors": ["Black"], "source": "tag"},
    ],
    "primaryColorAxis": "Painted Frame, Painted Mudguard",
    "primaryAxisKey": "pool:1",
}

MOCK_POOL = [
    # Painted Frame buckets
    {"outputItemName": "Painted Frame", "processId": "PRC-0", "productTag": "", "color": "Red", "availableQty": 120, "producedQty": 120, "consumedQty": 0},
    {"outputItemName": "Painted Frame", "processId": "PRC-0", "productTag": "", "color": "Blue", "availableQty": 0, "producedQty": 10, "consumedQty": 10},
    {"outputItemName": "Painted Frame", "processId": "PRC-0", "productTag": "", "color": "Green", "availableQty": 8, "producedQty": 8, "consumedQty": 0},
    {"outputItemName": "Painted Frame", "processId": "PRC-0", "productTag": "", "color": "Yellow", "availableQty": -5, "producedQty": 5, "consumedQty": 10},
    # No exact "Purple" bucket — only a composite one it is a token of.
    {"outputItemName": "Painted Frame", "processId": "PRC-0", "productTag": "", "color": "BCP / Purple", "availableQty": 25, "producedQty": 25, "consumedQty": 0},
    # Already packed into a product — must NOT count toward what a new lot can draw.
    {"outputItemName": "Painted Frame", "processId": "PRC-0", "productTag": "PRD-9", "color": "Red", "availableQty": 500, "producedQty": 500, "consumedQty": 0},
    # Painted Mudguard buckets
    {"outputItemName": "Painted Mudguard", "processId": "PRC-0", "productTag": "", "color": "Red", "availableQty": 90, "producedQty": 90, "consumedQty": 0},
    {"outputItemName": "Painted Mudguard", "processId": "PRC-0", "productTag": "", "color": "Blue", "availableQty": 40, "producedQty": 40, "consumedQty": 0},
    {"outputItemName": "Painted Mudguard", "processId": "PRC-0", "productTag": "", "color": "Green", "availableQty": 40, "producedQty": 40, "consumedQty": 0},
    {"outputItemName": "Painted Mudguard", "processId": "PRC-0", "productTag": "", "color": "Yellow", "availableQty": 10, "producedQty": 10, "consumedQty": 0},
    {"outputItemName": "Painted Mudguard", "processId": "PRC-0", "productTag": "", "color": "Purple", "availableQty": 30, "producedQty": 30, "consumedQty": 0},
]

MOCK_STOCK = [
    {"name": "Fitting Black", "size": "", "currentStock": 500},
    {"name": "Assembly Screws", "size": "", "currentStock": 9000},
]

MOCK_API_RESPONSES = {
    # "Cyan" is a Color Master leftover belonging to no axis -> "Other" bucket.
    "getProcessColorGroups": {"success": True, "data": ["Black", "Blue", "Cyan", "Green", "Purple", "Red", "Yellow"]},
    "getProcessColorAxes": {"success": True, "data": MOCK_AXES},
    "getProcessComponentsData": {"success": True, "data": MOCK_COMPONENTS},
    "getWarehousePoolData": {"success": True, "data": MOCK_POOL},
    "getProcessWipData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": MOCK_STOCK},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
}

READ_ROWS_JS = """
    Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row')).map(row => {
        const el = row.querySelector('.production-color-avail');
        return {
            color: row.dataset.color,
            group: row.dataset.group,
            text: el ? el.innerText.trim() : null,
            title: el ? (el.title || '') : '',
            danger: el ? el.classList.contains('text-danger') : false,
            bold: el ? el.classList.contains('fw-bold') : false,
            muted: el ? el.classList.contains('text-muted') : false
        };
    })
"""


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
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
            App.State.globalItems = {json.dumps(MOCK_ITEMS)};
            window.__mockResponses = {json.dumps(MOCK_API_RESPONSES)};
            window.google = {{
                script: {{
                    run: {{
                        withSuccessHandler(cb) {{
                            const runner = {{ withFailureHandler() {{ return runner; }} }};
                            Object.keys(window.__mockResponses).forEach(method => {{
                                runner[method] = (...args) => setTimeout(() => cb(window.__mockResponses[method]), 30);
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

        print("[Step 1] Open Create modal, select the pool-driven process...")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.evaluate("""
            document.getElementById('productionProcessId').value = 'PRC-1';
            App.Production.handleProcessChange('PRC-1');
        """)
        page.wait_for_timeout(1500)

        rows = page.evaluate(READ_ROWS_JS)
        by_color = {r["color"]: r for r in rows}
        print("    rendered:", json.dumps(rows, indent=None))

        check(len(rows) > 0, "checklist rendered color rows")

        print("\n[Step 2] Multi-input pool group shows the SMALLEST available (the real ceiling)...")
        red = by_color.get("Red", {})
        # Frame 120 vs Mudguard 90 -> 90 is what actually caps a Red lot.
        check(red.get("text") == "90 avail.",
              f"Red shows the scarcest input's qty (90), not 120 — got {red.get('text')!r}")
        check("Painted Frame: 120" in red.get("title", "") and "Painted Mudguard: 90" in red.get("title", ""),
              f"Red tooltip breaks the min down per input — got {red.get('title')!r}")
        check(not red.get("danger") and red.get("muted"), "Red (stock available) renders muted, not red")

        print("\n[Step 3] A tagged/packed pool row must NOT inflate what a new lot can draw...")
        # The 500-qty Red bucket carries productTag PRD-9; if it leaked in,
        # Painted Frame's Red would read 620 and the min would still be 90 —
        # so assert on the tooltip, which shows the per-item figure directly.
        check("Painted Frame: 120" in red.get("title", "") and "620" not in red.get("title", ""),
              "tagged (productTag) pool rows are excluded from availability")

        print("\n[Step 4] Zero availability is flagged, not silently gray...")
        blue = by_color.get("Blue", {})
        # Frame Blue is 0 available -> no Blue can be built regardless of the 40 mudguards.
        check(blue.get("text") == "0 avail.", f"Blue shows 0 avail. — got {blue.get('text')!r}")
        check(blue.get("danger") and blue.get("bold"), "Blue's 0 avail. is red + bold")

        print("\n[Step 5] Negative availability is flagged too...")
        yellow = by_color.get("Yellow", {})
        check(yellow.get("text") == "-5 avail.", f"Yellow shows -5 avail. — got {yellow.get('text')!r}")
        check(yellow.get("danger") and yellow.get("bold"), "Yellow's negative avail. is red + bold")

        print("\n[Step 6] Genuinely-low-but-positive availability stays readable...")
        green = by_color.get("Green", {})
        check(green.get("text") == "8 avail.", f"Green shows 8 avail. (min of 8/40) — got {green.get('text')!r}")
        check(not green.get("danger"), "Green (8 available) is not flagged red")

        print("\n[Step 7] A color with no exact bucket resolves through its composite bucket...")
        purple = by_color.get("Purple", {})
        # Painted Frame has only "BCP / Purple" (25); Mudguard has a real "Purple" (30).
        check(purple.get("text") == "25 avail.",
              f"Purple resolves 'BCP / Purple' -> 25 — got {purple.get('text')!r}")

        print("\n[Step 8] An ITEM-sourced tag-axis color reads Stock, not the pool...")
        black = by_color.get("Black", {})
        check(black.get("text") == "500 avail.",
              f"Black (Fitting axis, ITEM source) shows its Stock of 500 — got {black.get('text')!r}")

        print("\n[Step 9] A color with no configured input stays blank...")
        cyan = by_color.get("Cyan", {})
        check(cyan.get("group") == "other", f"Cyan landed in the 'Other' bucket — got {cyan.get('group')!r}")
        check(cyan.get("text") == "", f"Cyan shows nothing rather than a bogus 0 — got {cyan.get('text')!r}")
        check(not cyan.get("danger"), "Cyan is not flagged red")

        print("\n[Step 10] Hints survive checking a color (no wipe on re-render/refresh)...")
        page.evaluate("""
            const row = Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row'))
                .find(r => r.dataset.color === 'Red');
            const cb = row.querySelector('.production-color-check');
            cb.checked = true;
            App.Production.handleColorCheckToggle(cb);
        """)
        page.wait_for_timeout(1200)
        after = {r["color"]: r for r in page.evaluate(READ_ROWS_JS)}
        check(after.get("Red", {}).get("text") == "90 avail.",
              f"Red still shows 90 avail. after being checked — got {after.get('Red', {}).get('text')!r}")
        check(after.get("Black", {}).get("text") == "500 avail.",
              f"Black still shows its Stock after the availability refresh — got {after.get('Black', {}).get('text')!r}")

        print("\n[Step 11] No JS errors...")
        check(len(console_errors) == 0, f"no page errors (got {console_errors})")

        browser.close()

        print("\n" + "=" * 60)
        if failures:
            print(f"{len(failures)} CHECK(S) FAILED:")
            for f in failures:
                print("  - " + f)
            sys.exit(1)
        print("ALL CHECKS PASSED")


if __name__ == "__main__":
    run()
