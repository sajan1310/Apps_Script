"""
Verification script: Remove Group button in Edit BOM form.
Tests three scenarios:
  1. Only one group → button blocked (toast, no removal)
  2. Two groups, second is empty → removes immediately without confirmation
  3. Two groups, second has a filled item → confirmation dialog appears
"""

import sys
import io
import time
from pathlib import Path
from playwright.sync_api import sync_playwright, expect

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False, slow_mo=200)
        ctx = browser.new_context()
        page = ctx.new_page()

        # Intercept google.script.run so the app doesn't stall waiting for the backend.
        # Real GAS calling convention chains withSuccessHandler/withFailureHandler
        # BEFORE the method name, e.g. google.script.run.withSuccessHandler(fn).methodName(...).
        page.add_init_script("""
            function makeRunner(successFn) {
                return new Proxy({}, { get: () => (...args) => successFn(null) });
            }
            window.google = {
                script: {
                    run: {
                        withSuccessHandler: function(fn) {
                            return { withFailureHandler: function() { return makeRunner(fn); } };
                        },
                        withFailureHandler: function() {
                            return { withSuccessHandler: function(fn) { return makeRunner(fn); } };
                        }
                    }
                }
            };
        """)

        url = DIST_HTML.as_uri()
        page.goto(url, wait_until="domcontentloaded")
        page.wait_for_timeout(1500)

        # ── Open the BOM create modal directly via JS (bypasses auth/lock screen) ──
        print("Opening BOM create modal via JS...")
        page.evaluate("""
            // Stub out the API.call so openCreateModal doesn't crash waiting for product ID
            window._origApiCall = window.Api && window.Api.call;
            if (window.Api) {
                window.Api.call = async (...args) => 'PRD-9999';
            }
            App.BOM.openCreateModal();
        """)

        modal = page.locator("#editBomModal")
        modal.wait_for(state="visible", timeout=TIMEOUT)
        print("  ✅ Modal opened")

        # ─────────────────────────────────────────────────────────────
        # SCENARIO 1: Remove the only group → silently reset to a fresh
        #             empty group (no toast, no confirmation needed)
        # ─────────────────────────────────────────────────────────────
        print("\n[Scenario 1] Remove the only (empty) group...")
        remove_btns = page.locator("#bomGroupsContainer .bom-group-card button:has-text('Remove Group')")
        expect(remove_btns).to_have_count(1, timeout=TIMEOUT)
        original_group_id = page.locator("#bomGroupsContainer .bom-group-card").first.get_attribute("id")

        remove_btns.first.click()
        page.wait_for_timeout(400)

        # A group is still present (auto-replaced), but it's a fresh card
        expect(remove_btns).to_have_count(1, timeout=TIMEOUT)
        new_group_id = page.locator("#bomGroupsContainer .bom-group-card").first.get_attribute("id")
        assert new_group_id != original_group_id, "Expected the lone group to be replaced with a fresh card"

        # No confirmation dialog and no toast for removing an empty lone group
        confirm_modal = page.locator("#confirmModal")
        assert not confirm_modal.is_visible(), "Confirm dialog should NOT appear for an empty lone group"
        print("  ✅ Lone empty group silently reset — no toast, no confirmation")

        # ─────────────────────────────────────────────────────────────
        # SCENARIO 2: Add second group (empty) → Remove should work
        #             without a confirmation dialog
        # ─────────────────────────────────────────────────────────────
        print("\n[Scenario 2] Add a second empty group and remove it...")
        add_group_btn = modal.locator("button:has-text('Add Process Group')")
        add_group_btn.click()
        page.wait_for_timeout(300)

        all_cards = page.locator("#bomGroupsContainer .bom-group-card")
        expect(all_cards).to_have_count(2, timeout=TIMEOUT)
        print("  Two groups present")

        # Click Remove Group on the second card
        second_remove = all_cards.nth(1).locator("button:has-text('Remove Group')")
        second_remove.click()
        page.wait_for_timeout(500)

        # Confirm modal should NOT have appeared (empty row = no confirmation)
        confirm_modal = page.locator("#confirmModal")
        confirm_visible = confirm_modal.is_visible()

        # Group count should be back to 1
        expect(all_cards).to_have_count(1, timeout=TIMEOUT)
        print(f"  Confirm dialog visible: {confirm_visible}")
        assert not confirm_visible, "Confirm dialog should NOT appear for an empty group"
        print("  ✅ Empty group removed immediately — no confirmation dialog")

        # ─────────────────────────────────────────────────────────────
        # SCENARIO 3: Add group, inject a filled select, then Remove →
        #             confirmation dialog must appear
        # ─────────────────────────────────────────────────────────────
        print("\n[Scenario 3] Add a group, fill an item, then remove it...")
        add_group_btn.click()
        page.wait_for_timeout(300)
        expect(all_cards).to_have_count(2, timeout=TIMEOUT)

        # Inject a non-empty value into the first item select of the second group
        page.evaluate("""
            const secondCard = document.querySelectorAll('#bomGroupsContainer .bom-group-card')[1];
            const sel = secondCard.querySelector('.comp-item-select');
            if (sel) {
                // Add and select a dummy option so value !== ''
                const opt = document.createElement('option');
                opt.value = '0';
                opt.text  = 'Test Item';
                opt.selected = true;
                sel.appendChild(opt);
            }
        """)

        second_remove2 = all_cards.nth(1).locator("button:has-text('Remove Group')")
        second_remove2.click()
        page.wait_for_timeout(500)

        # Confirmation modal SHOULD appear
        expect(confirm_modal).to_be_visible(timeout=TIMEOUT)
        confirm_text = page.locator("#confirmMessageText").inner_text()
        print(f"  Confirm message: '{confirm_text}'")
        assert "component" in confirm_text.lower(), f"Expected 'component' in confirm text, got: {confirm_text}"
        print("  ✅ Confirmation dialog appeared with correct message")

        # Dismiss without confirming — click the modal's close/cancel button
        # (avoids Escape propagating to the outer editBomModal keyboard listener)
        dismiss_btn = page.locator("#confirmModal [data-bs-dismiss='modal'], #confirmModal .btn-secondary").first
        dismiss_btn.click()
        page.wait_for_timeout(600)
        # Outer BOM modal must still be open
        expect(modal).to_be_visible(timeout=TIMEOUT)
        expect(all_cards).to_have_count(2, timeout=TIMEOUT)
        print("  ✅ After cancel, group still present and BOM modal still open")

        # Now confirm deletion — re-locate button fresh to avoid stale reference
        page.locator("#bomGroupsContainer .bom-group-card").nth(1).locator("button:has-text('Remove Group')").click()
        page.wait_for_timeout(500)
        expect(confirm_modal).to_be_visible(timeout=TIMEOUT)
        page.locator("#confirmActionBtn").click()
        page.wait_for_timeout(600)
        expect(all_cards).to_have_count(1, timeout=TIMEOUT)
        print("  ✅ After confirming, group removed")

        # ─────────────────────────────────────────────────────────────
        print("\n✅ ALL SCENARIOS PASSED")
        browser.close()
        return True

if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
