"""
Verification script — Fix #8 (Per-Color matrix axis-collision guard), one of
the 9 architectural gaps verified+fixed on 2026-07-13 (see
verification_2026_07_13_architectural_gaps in project memory).

The Per-Color matrix table's column identity is color-string-only, shared
across every independent Color Axis (getMatrixColumnIndex/addMatrixColorColumn,
Script_Production.html) — a real composite (axis, color) identity would be a
much larger refactor. Two bugs were found in the two live code paths that
still touch that shared column with no axis awareness at all:

1. onColorQtyChanged used to overwrite EVERY row's cell in a shared column
   whenever ANY axis's same-named color qty changed -- so editing a Rim
   axis's "Purple" quantity would silently stomp a Frame axis's own
   "Purple"-tagged component cells with the wrong qtyPerUnit basis.
2. removeMatrixColorColumn used to delete the ENTIRE shared column the
   moment any ONE axis's same-named color was unchecked -- wiping cells
   still in use by a different, still-checked axis.

Both are now guarded: onColorQtyChanged skips the shared-column update
entirely while 2+ axes have the same color name checked at once (ambiguous,
can't correctly represent both), and removeMatrixColorColumn checks whether
the color is still checked under ANY row before deleting the column.

This test constructs the two-axis-collision DOM state directly (two
checklist rows both named "Purple" under different axis groups, and a
matrix table with one shared "Purple" column feeding two different items)
rather than driving the full natural process-recipe-population flow, since
this is testing the guard functions' own logic in isolation.

Run: python .pw-test/verify_matrix_axis_collision_guard.py
"""
import sys
import io
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

failures = []
def check(cond, msg):
    print(("PASS: " if cond else "FAIL: ") + msg)
    if not cond:
        failures.append(msg)


SETUP_JS = """
    () => {
        const checklist = document.getElementById('productionColorChecklist');
        checklist.innerHTML = `
            <div class="form-check d-flex align-items-center gap-2 production-color-row" data-color="Purple" data-group="tag:rim color" data-primary="true">
                <input class="form-check-input production-color-check" type="checkbox" onchange="App.Production.handleColorCheckToggle(this)" checked>
                <label>Purple (Rim)</label>
                <input type="number" class="form-control form-control-sm production-color-qty" value="10" oninput="App.Production.onColorQtyChanged(this.closest('.production-color-row'))">
            </div>
            <div class="form-check d-flex align-items-center gap-2 production-color-row" data-color="Purple" data-group="tag:frame color" data-primary="false">
                <input class="form-check-input production-color-check" type="checkbox" onchange="App.Production.handleColorCheckToggle(this)" checked>
                <label>Purple (Frame)</label>
                <input type="number" class="form-control form-control-sm production-color-qty" value="6" oninput="App.Production.onColorQtyChanged(this.closest('.production-color-row'))">
            </div>
        `;

        const headerRow = document.getElementById('productionColorMatrixHeaderRow');
        headerRow.innerHTML = `
            <th>Item / Pool Name</th>
            <th data-color="Purple">Purple</th>
            <th></th>
        `;
        const body = document.getElementById('productionColorMatrixBody');
        body.innerHTML = `
            <tr data-merged="false">
                <td>Rim Screw</td>
                <td><input type="number" class="matrix-qty" data-qty-per-unit="2" value="99"></td>
                <td></td>
            </tr>
            <tr data-merged="false">
                <td>Frame Bolt</td>
                <td><input type="number" class="matrix-qty" data-qty-per-unit="5" value="88"></td>
                <td></td>
            </tr>
        `;
    }
"""


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context()
        page = ctx.new_page()

        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(500)

        print("[Setup] Open the Create Production Lot modal (for its DOM containers), inject a 2-axis 'Purple' collision...")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.evaluate(SETUP_JS)

        print("\n[Test 1] onColorQtyChanged skips the shared column while 'Purple' is checked under 2 axes at once...")
        rim_screw_before = page.evaluate("document.querySelectorAll('.matrix-qty')[0].value")
        frame_bolt_before = page.evaluate("document.querySelectorAll('.matrix-qty')[1].value")
        check(rim_screw_before == "99" and frame_bolt_before == "88", f"sanity: seeded cell values are 99/88 before any change (got {rim_screw_before}/{frame_bolt_before})")

        page.evaluate("""
            const rimRow = document.querySelector('.production-color-row[data-group="tag:rim color"]');
            App.Production.onColorQtyChanged(rimRow);
        """)
        page.wait_for_timeout(100)

        rim_screw_after = page.evaluate("document.querySelectorAll('.matrix-qty')[0].value")
        frame_bolt_after = page.evaluate("document.querySelectorAll('.matrix-qty')[1].value")
        check(rim_screw_after == "99", f"Rim Screw's cell is UNCHANGED while the collision is live (got {rim_screw_after!r}, would be '20' = 10*2 if the old unguarded code ran)")
        check(frame_bolt_after == "88", f"Frame Bolt's cell is UNCHANGED too (got {frame_bolt_after!r}, would be '50' = 10*5 under the OLD bug -- Rim's qty wrongly applied to Frame's item)")

        print("\n[Test 2] once the collision resolves (only one axis checked), onColorQtyChanged updates normally again...")
        page.evaluate("""
            const frameRow = document.querySelector('.production-color-row[data-group="tag:frame color"]');
            frameRow.querySelector('.production-color-check').checked = false;
        """)
        page.evaluate("""
            const rimRow = document.querySelector('.production-color-row[data-group="tag:rim color"]');
            rimRow.querySelector('.production-color-qty').value = '10';
            App.Production.onColorQtyChanged(rimRow);
        """)
        page.wait_for_timeout(100)
        rim_screw_resolved = page.evaluate("document.querySelectorAll('.matrix-qty')[0].value")
        check(rim_screw_resolved == "20", f"with only Rim checked (collision resolved), the shared column updates normally: 10 * qtyPerUnit(2) = 20 (got {rim_screw_resolved!r})")

        print("\n[Test 3] removeMatrixColorColumn does NOT delete a column another axis's checked row still needs...")
        page.evaluate("""
            const frameRow = document.querySelector('.production-color-row[data-group="tag:frame color"]');
            frameRow.querySelector('.production-color-check').checked = true; // both checked again
        """)
        colIndexBefore = page.evaluate("App.Production.getMatrixColumnIndex('Purple')")
        check(colIndexBefore != -1, "sanity: the 'Purple' column exists before any removal attempt")

        page.evaluate("""
            // Simulate unchecking Rim's row (the row itself reflects unchecked,
            // as handleColorCheckToggle's own flow would leave it) then calling
            // the column-removal function, exactly as that handler does.
            const rimRow = document.querySelector('.production-color-row[data-group="tag:rim color"]');
            rimRow.querySelector('.production-color-check').checked = false;
            App.Production.removeMatrixColorColumn('Purple');
        """)
        colIndexAfterOne = page.evaluate("App.Production.getMatrixColumnIndex('Purple')")
        check(colIndexAfterOne != -1, "column survives -- Frame's row is STILL checked and still needs it")
        frame_bolt_survives = page.evaluate("document.querySelectorAll('.matrix-qty')[1] ? document.querySelectorAll('.matrix-qty')[1].value : null")
        check(frame_bolt_survives is not None, "Frame Bolt's cell data was not wiped out along with a (non-)removed column")

        print("\n[Test 4] removeMatrixColorColumn DOES delete the column once nothing checked needs it anymore...")
        page.evaluate("""
            const frameRow = document.querySelector('.production-color-row[data-group="tag:frame color"]');
            frameRow.querySelector('.production-color-check').checked = false;
            App.Production.removeMatrixColorColumn('Purple');
        """)
        colIndexAfterBoth = page.evaluate("App.Production.getMatrixColumnIndex('Purple')")
        check(colIndexAfterBoth == -1, "column is removed once NEITHER axis's row is checked anymore (normal, non-collision cleanup still works)")

        if console_errors:
            print("\nConsole/page errors:")
            for e in console_errors:
                print(f"    {e}")
            failures.extend(console_errors)

        browser.close()

        if failures:
            print(f"\n{len(failures)} CHECK(S)/ERROR(S) FAILED")
        else:
            print("\nALL CHECKS PASSED")
        return not failures


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
