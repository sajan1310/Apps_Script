"""
Playwright Verification for Notification Bell & History Panel (App.Notify)
"""
import sys
import io
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"

failures = 0
def check(cond, msg):
    global failures
    if cond:
        print(f"  PASS: {msg}")
    else:
        failures += 1
        print(f"  FAIL: {msg}")

def main():
    print("Testing Notification Bell & History Panel (App.Notify)...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(DIST_HTML.as_uri())

        # Wait for DOM and App initialization
        page.wait_for_selector("#notif-bell-btn")

        # 1. Verify App.Notify exists
        has_notify = page.evaluate("typeof App.Notify !== 'undefined'")
        check(has_notify, "App.Notify is defined on window.App")

        # Clear any startup notifications for a clean baseline test.
        #
        # App.Init awaits Notify.checkSmartAlerts() + checkBackendAlerts() on
        # boot, and on file:// the mock-data harness at the top of
        # Script_Core.html is active, so those really do produce alerts (low
        # stock, overdue POs). They resolve asynchronously: clearing the moment
        # #notif-bell-btn appears sometimes ran BEFORE that pass finished, and
        # its 2 alerts then landed on top of the 2 this test adds itself — every
        # count/message assertion below saw 4 and failed. Intermittent, and
        # nothing to do with the bell.
        #
        # So: wait for the startup pass to go quiet (item count stable across
        # consecutive polls) and only then clear.
        page.wait_for_function(
            """
            () => {
                const n = (window.App?.Notify?.items || []).length;
                const stable = window.__notifPrev === n ? (window.__notifStable || 0) + 1 : 0;
                window.__notifPrev = n;
                window.__notifStable = stable;
                return stable >= 3;
            }
            """,
            timeout=10000,
        )
        page.evaluate("App.Notify.clear()")
        baseline = page.evaluate("(App.Notify.items || []).length")
        check(baseline == 0, f"baseline is genuinely empty before the test adds its own (got {baseline})")

        # 2. Check initial state (badge hidden, empty notification list)
        badge_visible = page.evaluate("document.getElementById('notif-badge').style.display !== 'none'")
        check(not badge_visible, "Notification badge is initially hidden after clear")

        # 3. Add notifications via App.Utils.showToast
        page.evaluate("App.Utils.showToast('Item saved successfully.')")
        page.evaluate("App.Utils.showToast('Failed to load vendors', true)")

        # Verify badge count is 2 and visible
        badge_text = page.evaluate("document.getElementById('notif-badge').textContent")
        badge_display = page.evaluate("document.getElementById('notif-badge').style.display")
        check(badge_text == "2", f"Badge count is 2 (got '{badge_text}')")
        check(badge_display != "none", f"Badge is visible (display: {badge_display})")

        # 4. Click Notification Bell button to open panel
        page.click("#notif-bell-btn")
        panel_display = page.evaluate("document.getElementById('notif-panel').style.display")
        check(panel_display == "flex", f"Notification panel is open (display: {panel_display})")

        # 5. Verify badge is cleared after opening panel
        badge_display_after_open = page.evaluate("document.getElementById('notif-badge').style.display")
        check(badge_display_after_open == "none", "Badge is hidden after opening history panel")

        # 6. Verify 2 notification items rendered in history list
        items_count = page.evaluate("document.querySelectorAll('#notif-list .notif-item').length")
        check(items_count == 2, f"Rendered 2 notification items in list (got {items_count})")

        has_error_class = page.evaluate("document.querySelector('#notif-list .notif-item').classList.contains('notif-error')")
        check(has_error_class, "Most recent error notification has .notif-error class")

        msg1 = page.evaluate("document.querySelectorAll('#notif-list .notif-item-msg')[0].textContent")
        check(msg1 == "Failed to load vendors", f"Item 1 message is '{msg1}'")

        msg2 = page.evaluate("document.querySelectorAll('#notif-list .notif-item-msg')[1].textContent")
        check(msg2 == "Item saved successfully.", f"Item 2 message is '{msg2}'")

        # 7. Test single notification removal (dismiss button)
        page.click("#notif-list .notif-item:first-child .notif-item-remove")
        items_count_after_remove = page.evaluate("document.querySelectorAll('#notif-list .notif-item').length")
        check(items_count_after_remove == 1, f"Items count after single dismiss is 1 (got {items_count_after_remove})")

        remaining_msg = page.evaluate("document.querySelector('#notif-list .notif-item-msg').textContent")
        check(remaining_msg == "Item saved successfully.", f"Remaining item is '{remaining_msg}'")

        # 8. Click outside to close panel
        page.click("body", position={"x": 10, "y": 10})
        panel_display_closed = page.evaluate("document.getElementById('notif-panel').style.display")
        check(panel_display_closed == "none", "Panel closes when clicking outside")

        # 9. Re-open and test Clear All button
        page.click("#notif-bell-btn")
        page.click("#notif-clear-btn")
        empty_text = page.evaluate("document.querySelector('#notif-list .notif-empty')?.textContent")
        check(empty_text == "No notifications yet.", f"Panel list reset to empty state (got '{empty_text}')")

        clear_btn_disabled = page.evaluate("document.getElementById('notif-clear-btn').disabled")
        check(clear_btn_disabled == True, "Clear All button is disabled when history is empty")

        browser.close()

    if failures > 0:
        print(f"\nResult: {failures} assertion(s) FAILED")
        sys.exit(1)
    else:
        print("\nResult: ALL ASSERTIONS PASSED!")

if __name__ == "__main__":
    main()
