"""
Verifies the new Warehouse Pool breakdown dialog features:
  1. A real-history combination ("Red") shows a disabled/protected X + checkbox.
  2. A removable placeholder combination ("Blue") shows an enabled X; clicking
     it (after confirm) removes it from the list.
  3. Bulk delete via checkboxes removes multiple removable combinations at once.
  4. "Add Combination" adds a new combination that then appears in the list.
Uses a small stateful mock (not static canned responses) so
excludeWarehousePoolColors/includeWarehousePoolColor actually affect what
getAllProcessColorGroups returns on the next refresh — a static per-method
mock can't represent that.
"""
import sys
import io
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
DIST_HTML = Path(r"C:\Users\erkar\my-app-script-project\dist\index.html")
TIMEOUT = 8000

MOCK_JS = r"""
App.State.globalProcesses = [{
  processId: 'PRC-1', processName: 'Fitting Frame', sequence: 1, lotPrefix: 'FF',
  outputItemName: 'Fitted Frame 16 inch', isFinalStage: false, active: true, processType: 'General'
}];

// Server-side state this mock simulates.
window.__mockState = {
  pool: [
    { outputItemName: 'Fitted Frame 16 inch', processId: 'PRC-1', productTag: '', color: 'Red',
      producedQty: 10, consumedQty: 2, availableQty: 8 }
  ],
  overrides: {} // processId -> { colorLower: 'INCLUDE'|'EXCLUDE' }
};

function computeKnownColors() {
  // Base "always known" colors for PRC-1: Red (real bucket) + Blue + Green (placeholders).
  const base = ['Red', 'Blue', 'Green'];
  const overrides = window.__mockState.overrides['PRC-1'] || {};
  const set = new Map(base.map(c => [c.toLowerCase(), c]));
  Object.keys(overrides).forEach(cLower => {
    if (overrides[cLower].action === 'EXCLUDE') set.delete(cLower);
    else set.set(cLower, overrides[cLower].color);
  });
  const colors = Array.from(set.values()).sort();
  // "Red" is the only one with real bucket history -> protected/not removable.
  const removable = colors.filter(c => c.toLowerCase() !== 'red');
  return { colors, removable };
}

window.google = {
  script: {
    run: {
      withSuccessHandler(cb) {
        const runner = { withFailureHandler() { return runner; } };
        runner.getWarehousePoolData = () => setTimeout(() => cb({ success: true, data: window.__mockState.pool }), 20);
        runner.getAllProcessColorGroups = () => setTimeout(() => cb({
          success: true, data: { 'PRC-1': computeKnownColors() }
        }), 20);
        runner.excludeWarehousePoolColors = (processId, colors) => setTimeout(() => {
          const list = Array.isArray(colors) ? colors : [colors];
          const removed = [];
          const blocked = [];
          list.forEach(c => {
            if (c.toLowerCase() === 'red') { blocked.push(c + ' (has real production/consumption history)'); return; }
            if (!window.__mockState.overrides['PRC-1']) window.__mockState.overrides['PRC-1'] = {};
            window.__mockState.overrides['PRC-1'][c.toLowerCase()] = { color: c, action: 'EXCLUDE' };
            removed.push(c);
          });
          const message = blocked.length === 0
            ? `Removed ${removed.length} combination(s).`
            : (removed.length > 0
              ? `Removed ${removed.length} combination(s). ${blocked.length} skipped (can't be removed): ${blocked.join('; ')}.`
              : `Nothing removed — can't be removed: ${blocked.join('; ')}.`);
          cb({ success: removed.length > 0 || blocked.length === 0, data: { removed, blocked }, message });
        }, 20);
        runner.includeWarehousePoolColor = (processId, color) => setTimeout(() => {
          if (!window.__mockState.overrides['PRC-1']) window.__mockState.overrides['PRC-1'] = {};
          window.__mockState.overrides['PRC-1'][color.toLowerCase()] = { color, action: 'INCLUDE' };
          cb({ success: true, data: { color }, message: `"${color}" added as a known combination for this process.` });
        }, 20);
        return runner;
      }
    }
  }
};
"""


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1400, "height": 900})
        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)
        page.evaluate(MOCK_JS)

        results = []

        print("[1] Open Warehouse Pool breakdown for PRC-1, initial state...")
        page.evaluate("App.Stock.loadWarehousePoolData()")
        page.wait_for_timeout(500)
        page.evaluate("App.Stock.openWarehousePoolProcessModal(encodeURIComponent('PRC-1'))")
        page.locator("#warehousePoolProcessModal").wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(300)

        rows = page.evaluate("""
            Array.from(document.querySelectorAll('#warehousePoolProcessModalBody tr')).map(tr => {
                const cells = tr.querySelectorAll('td');
                const colorBadge = cells[4]?.querySelector('.badge');
                const delBtn = cells[6]?.querySelector('button[title*="Remove"], button[disabled]');
                return {
                    color: colorBadge ? colorBadge.textContent.trim() : null,
                    deleteDisabled: delBtn ? delBtn.disabled : null
                };
            })
        """)
        print(f"  Rows: {rows}")
        red_row = next((r for r in rows if r['color'] == 'Red'), None)
        blue_row = next((r for r in rows if r['color'] == 'Blue'), None)
        ok1 = red_row is not None and red_row['deleteDisabled'] is True and blue_row is not None and blue_row['deleteDisabled'] is False
        print("  " + ("✅ PASS — Red protected, Blue removable" if ok1 else "❌ FAIL"))
        results.append(ok1)

        print("[2] Click X on 'Blue' (removable) -> confirm -> Blue disappears...")
        page.evaluate("""
            const row = Array.from(document.querySelectorAll('#warehousePoolProcessModalBody tr'))
                .find(tr => tr.querySelector('.badge')?.textContent.trim() === 'Blue');
            row.querySelector('button[title="Remove this combination"]').click();
        """)
        page.wait_for_timeout(300)
        confirm_visible = page.locator("text=Remove").count() > 0
        page.click("#confirmModal .btn-danger, #confirmModal .btn-primary, #confirmModal button:has-text('OK'), #confirmModal button:has-text('Yes'), #confirmModal button:has-text('Confirm')")
        page.wait_for_timeout(500)
        colors_after = page.evaluate("""
            Array.from(document.querySelectorAll('#warehousePoolProcessModalBody .badge.bg-info')).map(b => b.textContent.trim())
        """)
        print(f"  Colors after removing Blue: {colors_after}")
        ok2 = 'Blue' not in colors_after and 'Red' in colors_after and 'Green' in colors_after
        print("  " + ("✅ PASS" if ok2 else "❌ FAIL"))
        results.append(ok2)

        print("[3] Try to remove 'Red' (protected) -> button should be disabled, no-op...")
        red_disabled = page.evaluate("""
            (() => {
                const row = Array.from(document.querySelectorAll('#warehousePoolProcessModalBody tr'))
                    .find(tr => tr.querySelector('.badge')?.textContent.trim() === 'Red');
                const btn = Array.from(row.querySelectorAll('button')).find(b => b.title.includes('history'));
                return btn ? btn.disabled : null;
            })()
        """)
        print(f"  Red delete button disabled: {red_disabled}")
        ok3 = red_disabled is True
        print("  " + ("✅ PASS" if ok3 else "❌ FAIL"))
        results.append(ok3)

        print("[4] Add a new combination 'Purple'...")
        page.fill("#warehousePoolAddComboInput", "Purple")
        page.click("button:has-text('Add Combination')")
        page.wait_for_timeout(500)
        colors_after_add = page.evaluate("""
            Array.from(document.querySelectorAll('#warehousePoolProcessModalBody .badge.bg-info')).map(b => b.textContent.trim())
        """)
        print(f"  Colors after adding Purple: {colors_after_add}")
        ok4 = 'Purple' in colors_after_add
        print("  " + ("✅ PASS" if ok4 else "❌ FAIL"))
        results.append(ok4)

        print("[4b] Search 'purple' -> only the Purple row shows...")
        page.fill("#searchWarehousePoolCombo", "purple")
        page.dispatch_event("#searchWarehousePoolCombo", "keyup")
        page.wait_for_timeout(300)
        colors_during_search = page.evaluate("""
            Array.from(document.querySelectorAll('#warehousePoolProcessModalBody .badge.bg-info')).map(b => b.textContent.trim())
        """)
        print(f"  Colors visible while searching 'purple': {colors_during_search}")
        ok4b = colors_during_search == ['Purple']
        print("  " + ("✅ PASS" if ok4b else "❌ FAIL"))
        results.append(ok4b)

        print("[4c] Clear search -> all rows return...")
        page.fill("#searchWarehousePoolCombo", "")
        page.dispatch_event("#searchWarehousePoolCombo", "keyup")
        page.wait_for_timeout(300)
        colors_after_clear = page.evaluate("""
            Array.from(document.querySelectorAll('#warehousePoolProcessModalBody .badge.bg-info')).map(b => b.textContent.trim())
        """)
        print(f"  Colors visible after clearing search: {colors_after_clear}")
        ok4c = set(colors_after_clear) == {'Green', 'Purple', 'Red'}
        print("  " + ("✅ PASS" if ok4c else "❌ FAIL"))
        results.append(ok4c)

        print("[5] Bulk-select Green + Purple, bulk delete...")
        page.evaluate("""
            document.querySelectorAll('#warehousePoolProcessModalBody .warehousepool-combo-chk').forEach(chk => {
                chk.checked = true;
                chk.dispatchEvent(new Event('change'));
            });
        """)
        page.wait_for_timeout(200)
        bulk_btn_visible = page.locator("#btnBulkDeleteWarehousePoolCombos").is_visible()
        print(f"  Bulk delete button visible after selecting: {bulk_btn_visible}")
        page.click("#btnBulkDeleteWarehousePoolCombos")
        page.wait_for_timeout(300)
        page.click("#confirmModal .btn-danger, #confirmModal .btn-primary, #confirmModal button:has-text('OK'), #confirmModal button:has-text('Yes'), #confirmModal button:has-text('Confirm')")
        page.wait_for_timeout(500)
        colors_after_bulk = page.evaluate("""
            Array.from(document.querySelectorAll('#warehousePoolProcessModalBody .badge.bg-info')).map(b => b.textContent.trim())
        """)
        print(f"  Colors after bulk delete: {colors_after_bulk}")
        ok5 = bulk_btn_visible and colors_after_bulk == ['Red']
        print("  " + ("✅ PASS" if ok5 else "❌ FAIL"))
        results.append(ok5)

        if console_errors:
            print("\n⚠️ console errors:")
            for e in console_errors:
                print(" ", e)

        browser.close()
        return all(results) and not console_errors


if __name__ == "__main__":
    ok = run()
    print("\n" + ("ALL PASS" if ok else "SOME FAILED"))
    sys.exit(0 if ok else 1)
