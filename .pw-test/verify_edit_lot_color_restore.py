"""
Verification script: re-opening a saved multi-color Production Lot for Edit
must restore every checked color to the EXACT row it was saved from, and
must never silently drop a color's checked/qty state.

Covers two gaps found during the 2026-07-08 Production modal audit
(openEditModal in Script.html):

1. Axis-identity collision on restore. Two independent Color Axes can share
   a literal color name (see bug_axis_color_name_collision_qty in project
   memory — e.g. a Rim axis and a Frame axis each having their own
   "Purple"). openEditModal used to restore a saved colorBreakdown by
   `.find(r => r.dataset.color === color)` — color name ONLY, no axis
   check — so re-opening a lot with such a colliding pair checked/filled the
   SAME first-found row twice while its true counterpart stayed blank.
   Fixed by matching on dataset.group === entry.axisKey first, with a
   claimed-rows guard so two colliding entries can never land on the same
   DOM row.

2. Orphaned/custom color data loss. populateColorChecklist only ever
   renders rows for colors the PROCESS currently knows about (see
   getProcessColorGroups/getProcessColorAxes) — never this lot's own saved
   breakdown. A lot saved with a one-off custom sub-group (see
   addCustomColorRow, e.g. "+ Add Custom Sub-Group") had NO row to restore
   into on re-open, so its checked/qty state silently vanished from the
   Edit form — and saving that edit again would then permanently drop it
   from the lot's record with no warning. Fixed by synthesizing a row (in
   the flat "Custom" bucket) for any breakdown entry with no matching row,
   preserving its original countsTowardTotal semantics.

Run: python .pw-test/verify_edit_lot_color_restore.py
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
    {"itemName": "Assembly Screws", "size": "", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "COMMON"},
    {"itemName": "Painted Rim - Blue-White", "size": "", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "Blue-White", "colorAxis": "Rim Color"},
    {"itemName": "Painted Rim - Purple", "size": "", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "Purple", "colorAxis": "Rim Color"},
    {"itemName": "Frame Tag - Purple", "size": "", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "Purple", "colorAxis": "Frame Color"},
    {"itemName": "Frame Tag - Green", "size": "", "sourceType": "ITEM", "qtyPerUnit": 1, "colorGroup": "Green", "colorAxis": "Frame Color"},
]

MOCK_AXES = {
    "axes": [
        {"key": "tag:rim color", "label": "Rim Color", "colors": ["Blue-White", "Purple"], "source": "tag"},
        {"key": "tag:frame color", "label": "Frame Color", "colors": ["Purple", "Green"], "source": "tag"},
    ],
    "primaryColorAxis": "Rim Color",
    "primaryAxisKey": "tag:rim color",
}

MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": ["Blue-White", "Green", "Purple"]},
    "getProcessColorAxes": {"success": True, "data": MOCK_AXES},
    "getProcessComponentsData": {"success": True, "data": MOCK_COMPONENTS},
    "getWarehousePoolData": {"success": True, "data": []},
    "getProcessWipData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
    "getProcessData": {"success": True, "data": [MOCK_PROCESS]},
    "getContractorsData": {"success": True, "data": []},
    "getModels": {"success": True, "data": []},
    "getProcessTypes": {"success": True, "data": []},
}

# The saved lot being re-opened for edit: Rim=Blue-White(6)+Purple(6) as the
# PRIMARY axis (counts toward the lot total), Frame=Purple(6)+Green(6) as a
# non-primary axis (does NOT count — describes the same physical batch from
# a second angle), plus one orphaned custom color with no backing row at all.
MOCK_LOT = {
    "rowIdx": 7,
    "date": "08/07/2026",
    "dateRaw": None,
    "productId": "", "productName": "",
    "qty": 12,
    "assignedBy": "Supervisor A",
    "assignedTo": "Contractor X",
    "status": "Pending",
    "remarks": "",
    "customComponents": [],
    "sheetRemarks": "",
    "processId": "PRC-1",
    "lotNumber": "LOT-FFA-0007",
    "contractorRate": 0,
    "contractorPayable": 0,
    "outputItemName": "Fitted Frame",
    "componentsConsumed": [],
    "color": "Blue-White, Purple, Purple, Green, One-Off Extra",
    "colorBreakdown": [
        {"color": "Blue-White", "qty": 6, "size": "", "isCustom": False, "countsTowardTotal": True, "axisKey": "tag:rim color"},
        {"color": "Purple", "qty": 6, "size": "", "isCustom": False, "countsTowardTotal": True, "axisKey": "tag:rim color"},
        {"color": "Purple", "qty": 6, "size": "", "isCustom": False, "countsTowardTotal": False, "axisKey": "tag:frame color"},
        {"color": "Green", "qty": 6, "size": "", "isCustom": False, "countsTowardTotal": False, "axisKey": "tag:frame color"},
        {"color": "One-Off Extra", "qty": 4, "size": "", "isCustom": True, "countsTowardTotal": True, "axisKey": ""},
    ],
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
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
            App.State.globalItems = [];
            App.State.globalColors = [];
            App.State.globalProduction = [{json.dumps(MOCK_LOT)}];
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

        print("[Step 1] Open Edit modal for the saved lot (2-axis Purple collision + orphaned custom color)...")
        page.evaluate("App.Production.openEditModal(0)")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(900)

        rows = page.evaluate("""
            Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row')).map(r => ({
                color: r.dataset.color, group: r.dataset.group,
                checked: r.querySelector('.production-color-check').checked,
                qty: r.querySelector('.production-color-qty').value
            }))
        """)
        print("  rows:", json.dumps(rows))

        rim_purple = next((r for r in rows if r['group'] == 'tag:rim color' and r['color'] == 'Purple'), None)
        frame_purple = next((r for r in rows if r['group'] == 'tag:frame color' and r['color'] == 'Purple'), None)
        rim_bluewhite = next((r for r in rows if r['group'] == 'tag:rim color' and r['color'] == 'Blue-White'), None)
        frame_green = next((r for r in rows if r['group'] == 'tag:frame color' and r['color'] == 'Green'), None)
        custom_row = next((r for r in rows if r['color'] == 'One-Off Extra'), None)

        print("\n[Step 2] Both colliding 'Purple' rows (Rim axis + Frame axis) restore independently...")
        check(rim_purple is not None and rim_purple['checked'] and rim_purple['qty'] == '6',
              f"Rim axis's own Purple row is checked with qty 6 (got {rim_purple})")
        check(frame_purple is not None and frame_purple['checked'] and frame_purple['qty'] == '6',
              f"Frame axis's own Purple row is ALSO checked with qty 6, not left blank (got {frame_purple})")
        check(rim_bluewhite is not None and rim_bluewhite['checked'] and rim_bluewhite['qty'] == '6',
              f"Rim axis's Blue-White row restores normally (got {rim_bluewhite})")
        check(frame_green is not None and frame_green['checked'] and frame_green['qty'] == '6',
              f"Frame axis's Green row restores normally (got {frame_green})")

        print("\n[Step 3] Orphaned custom color ('One-Off Extra') is synthesized as its own row, not silently dropped...")
        check(custom_row is not None, f"a row exists for the orphaned custom color (got {custom_row})")
        if custom_row:
            check(custom_row['checked'] and custom_row['qty'] == '4',
                  f"custom row is checked with its saved qty 4 (got {custom_row})")

        print("\n[Step 4] getCheckedColorQtys() reflects all 5 original entries, correctly attributed...")
        checked = json.loads(page.evaluate("JSON.stringify(App.Production.getCheckedColorQtys())"))
        print("  checked:", json.dumps(checked))
        check(len(checked) == 5, f"exactly 5 checked rows survive the reopen (got {len(checked)})")
        rim_entries = [c for c in checked if c['axisKey'] == 'tag:rim color']
        frame_entries = [c for c in checked if c['axisKey'] == 'tag:frame color']
        check(len(rim_entries) == 2 and all(c['qty'] == 6 for c in rim_entries),
              f"Rim axis contributes 2 entries at qty 6 each (got {rim_entries})")
        check(len(frame_entries) == 2 and all(c['qty'] == 6 for c in frame_entries),
              f"Frame axis contributes 2 entries at qty 6 each, distinct from Rim's (got {frame_entries})")
        custom_entry = next((c for c in checked if c['color'] == 'One-Off Extra'), None)
        check(custom_entry is not None and custom_entry['qty'] == 4 and custom_entry['countsTowardTotal'] is True,
              f"custom entry present with qty 4 and countsTowardTotal=true preserved (got {custom_entry})")

        if console_errors:
            print("\n⚠️  Console/page errors:")
            for e in console_errors:
                print(f"    {e}")
            failures.append("console errors present")

        print("\n" + ("ALL TESTS PASSED" if not failures else f"{len(failures)} TEST(S) FAILED"))
        browser.close()
        return len(failures) == 0


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
