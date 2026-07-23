"""
Verifies the "+ Add Custom Sub-Group" color field on the Production Lot
form is now a real Select2 dropdown (initCustomColorInputSelect2,
Script_Production.html) instead of a plain text input: it offers existing
Color Master names as searchable options, still lets the operator type a
genuinely new name (tags:true + createTag), and correctly resolves a
typed name that only differs by case to the existing Color Master entry
instead of minting a visually-duplicate option - same pick-or-create
pattern already proven for Assigned To (initContractorSelect2).

Run: python .pw-test/verify_production_custom_color_select2.py
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
    "lotPrefix": "FP", "outputItemName": "12 inch Painted Frame", "isFinalStage": False,
    "active": True, "processType": "General"
}
MOCK_COLORS = [{"name": "Red", "remarks": ""}, {"name": "Blue", "remarks": ""}, {"name": "Sea Green", "remarks": ""}]
MOCK_API_RESPONSES = {
    "getProcessColorGroups": {"success": True, "data": ["Red", "Blue"]},
    "getProcessColorAxes": {"success": True, "data": {"axes": [], "primaryColorAxis": "", "primaryAxisKey": ""}},
    "getProcessComponentsData": {"success": True, "data": []},
    "getWarehousePoolData": {"success": True, "data": []},
    "getProcessWipData": {"success": True, "data": []},
    "getStockData": {"success": True, "data": []},
    "getContractorRateForProcess": {"success": True, "data": {"ratePerUnit": 0}},
    "getColors": {"success": True, "data": MOCK_COLORS},
}

failures = 0
def check(cond, msg):
    global failures
    if cond:
        print(f"  PASS: {msg}")
    else:
        failures += 1
        print(f"  FAIL: {msg}")


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1400, "height": 900})
        console_errors = []
        page.on("pageerror", lambda e: console_errors.append(str(e)))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(500)
        page.evaluate(f"""
            App.State.globalProcesses = [{json.dumps(MOCK_PROCESS)}];
            App.State.globalModels = [{{name: 'General', remarks: ''}}];
            App.State.globalItems = [];
            window.__mockResponses = {json.dumps(MOCK_API_RESPONSES)};
            window.google = {{
                script: {{ run: {{
                    withSuccessHandler(cb) {{
                        const runner = {{ withFailureHandler() {{ return runner; }} }};
                        Object.keys(window.__mockResponses).forEach(method => {{
                            runner[method] = (...args) => setTimeout(() => cb(window.__mockResponses[method]), 20);
                        }});
                        return runner;
                    }}
                }} }}
            }};
        """)

        print("[1] Open Create Production modal, drive Size->Model cascade, wait for the color checklist to render...")
        page.evaluate("App.Production.openCreateModal()")
        page.locator("#editProductionModal").wait_for(state="visible", timeout=TIMEOUT)
        page.evaluate("""
            document.getElementById('productionSize').value = '12 inch';
            App.Production.handleSizeChange('12 inch');
        """)
        page.wait_for_timeout(300)
        page.evaluate("""
            document.getElementById('productionModel').value = 'General';
            App.Production.handleModelChange('General');
        """)
        page.wait_for_timeout(600)

        print("\n[2] productionCustomColorInput is a real <select>, Select2-enhanced")
        tag_name = page.evaluate("document.getElementById('productionCustomColorInput').tagName")
        check(tag_name == "SELECT", f"underlying element is a <select>, not a text <input> (got {tag_name})")
        has_select2 = page.evaluate("!!window.jQuery(document.getElementById('productionCustomColorInput')).data('select2')")
        check(has_select2, "Select2 widget is actually initialized on it")

        print("\n[3] Opening the dropdown offers the Color Master suggestions")
        trigger = page.locator("#productionCustomColorInput + .select2 .select2-selection")
        trigger.click()
        page.wait_for_timeout(200)
        options_text = page.locator(".select2-results__option").all_inner_texts()
        check(any("Sea Green" in t for t in options_text), f"'Sea Green' (a Color Master entry never used by this process's recipe) is offered as a real option (got {options_text})")

        print("\n[4] Picking an existing color, adding it, and confirming the field visibly clears afterward")
        # Dropdown is already open from step 3 - picking the option directly
        # instead of pressing Escape first, since Escape inside a Bootstrap
        # modal closes the WHOLE MODAL (not just the Select2 dropdown),
        # which would make every subsequent locator in this test fail.
        page.locator(".select2-results__option", has_text="Sea Green").first.click()
        page.wait_for_timeout(150)
        selected_before_add = page.evaluate("document.getElementById('productionCustomColorInput').value")
        check(selected_before_add == "Sea Green", f"picking the option sets the underlying select's value (got '{selected_before_add}')")

        page.click("button:has-text('Add Custom Sub-Group')")
        page.wait_for_timeout(200)
        row_added = page.evaluate("""
            !!Array.from(document.querySelectorAll('#productionColorChecklist .production-color-row'))
                .find(r => r.dataset.color === 'Sea Green')
        """)
        check(row_added, "a checklist row for 'Sea Green' was actually added")
        cleared_value = page.evaluate("document.getElementById('productionCustomColorInput').value")
        check(cleared_value == "", f"the Select2 field's underlying value is cleared after adding (got '{cleared_value}')")
        visible_text = page.locator("#productionCustomColorInput + .select2 .select2-selection__rendered").inner_text()
        check("Sea Green" not in visible_text, f"the VISIBLE Select2 selection box no longer shows the just-added name (got '{visible_text}')")

        print("\n[5] Typing a genuinely new name (not in Color Master) still works via tags:true")
        trigger.click()
        page.wait_for_timeout(150)
        search_box = page.locator(".select2-search__field")
        search_box.fill("Sunrise Coral")
        page.wait_for_timeout(200)
        new_tag_option = page.locator(".select2-results__option", has_text="Sunrise Coral")
        check(new_tag_option.count() > 0, "a 'create new tag' option appears for a genuinely new typed name")
        new_tag_option.first.click()
        page.wait_for_timeout(150)
        new_value = page.evaluate("document.getElementById('productionCustomColorInput').value")
        check(new_value == "Sunrise Coral", f"typing+selecting a new name sets it as the value (got '{new_value}')")

        print("\n[6] Typing an existing name with different casing resolves to the SAME Color Master entry, not a duplicate")
        page.evaluate("window.jQuery('#productionCustomColorInput').val(null).trigger('change')")
        page.wait_for_timeout(150)
        trigger.click()
        page.wait_for_timeout(150)
        page.locator(".select2-search__field").fill("red")
        page.wait_for_timeout(200)
        lower_option = page.locator(".select2-results__option", has_text="red")
        lower_option.first.click()
        page.wait_for_timeout(150)
        resolved_value = page.evaluate("document.getElementById('productionCustomColorInput').value")
        check(resolved_value == "Red", f"typing lowercase 'red' resolves to the existing 'Red' Color Master entry, not a new 'red' tag (got '{resolved_value}')")

        if console_errors:
            print("\n⚠️ console errors:")
            for e in console_errors:
                print(" ", e)
            global failures
            failures += len(console_errors)

        browser.close()
        return failures == 0


if __name__ == "__main__":
    ok = run()
    print("\n" + ("ALL PASS" if ok else "SOME FAILED"))
    sys.exit(0 if ok else 1)
