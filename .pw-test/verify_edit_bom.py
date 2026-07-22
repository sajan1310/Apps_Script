"""
Verification script: Edit BOM form.
Tests:
  1. Form pre-fill: product ID, name, remarks, components, additional costs
  2. Multi-group BOM renders correct number of group cards with correct names
  3. Orphan item (not in Items Master) preserved via synthetic option
  4. Submit in edit mode sends productId in formData (confirms edit vs create)
"""

import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

# A BOM with 2 components in 2 different process groups
MOCK_BOM = {
    "productId": "PRD-1001",
    "productName": "Maharaja Super Cruiser 24 Speed",
    "remarks": "Premium variant - use only approved suppliers.",
    "components": [
        {
            "itemName": "9-PLUS-RIM-PAPER---M-COLOR",
            "size": "12X1.75",
            "narration": "Color Rim paper standard quality",
            "rate": 120.5,
            "vendor": "Avon Cycles",
            "qtyPerProduct": 2,
            "lineCost": 241,
            "processId": "PROC-FRAME",
            "processGroup": "Frame Assembly"
        },
        {
            "itemName": "BRAKE-PAD",
            "size": "Standard",
            "narration": "Front",
            "rate": 50,
            "vendor": "Hero MotoCorp",
            "qtyPerProduct": 1,
            "lineCost": 50,
            "processId": "PROC-FITTING",
            "processGroup": "Fitting & Finishing"
        }
    ],
    "totalCost": 291,
    "totalQty": 3,
    "additionalCosts": [
        {"description": "Labor - Fitting", "rate": 50},
        {"description": "Paint Cost", "rate": 75}
    ],
    "totalAdditionalCost": 125,
    "grandTotal": 416
}

# Items Master entries matching the BOM components
MOCK_ITEMS = [
    {
        "name": "9-PLUS-RIM-PAPER---M-COLOR",
        "size": "12X1.75",
        "narration": "Color Rim paper standard quality",
        "vendors": [
            {"vendor": "Avon Cycles", "rate": 120.5},
            {"vendor": "DT Swiss", "rate": 118.0}
        ]
    },
    {
        "name": "BRAKE-PAD",
        "size": "Standard",
        "narration": "Front",
        "vendors": [
            {"vendor": "Hero MotoCorp", "rate": 50}
        ]
    }
]

# Process Master entries so the group selects resolve processId -> processName
MOCK_PROCESSES = [
    {"processId": "PROC-FRAME", "processName": "Frame Assembly", "active": True},
    {"processId": "PROC-FITTING", "processName": "Fitting & Finishing", "active": True}
]


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False, slow_mo=150)
        ctx = browser.new_context()
        page = ctx.new_page()

        url = DIST_HTML.as_uri()
        page.goto(url, wait_until="domcontentloaded")
        page.wait_for_timeout(1500)

        # Inject mock state — bypasses the auth/load-data flow entirely
        page.evaluate(f"""
            App.State.globalBOMs = [{json.dumps(MOCK_BOM)}];
            App.State.globalItems = {json.dumps(MOCK_ITEMS)};
            App.State.globalProcesses = {json.dumps(MOCK_PROCESSES)};
        """)

        modal = page.locator("#editBomModal")

        # ─────────────────────────────────────────────────────────────
        # SCENARIO 1: Form pre-fill (product ID, name, remarks, button)
        # ─────────────────────────────────────────────────────────────
        print("\n[Scenario 1] Open edit modal and verify pre-fill...")
        page.evaluate("App.BOM.openEditModal(0)")
        modal.wait_for(state="visible", timeout=TIMEOUT)
        print("  ✅ Edit modal opened")

        title_text = page.locator("#bomFormTitle").inner_text()
        print(f"  Modal title: '{title_text}'")
        assert "PRD-1001" in title_text, f"Expected PRD-1001 in title, got: {title_text}"
        print("  ✅ Title contains product ID")

        visible_id = page.locator("#bomVisibleProductId").input_value()
        assert visible_id == "PRD-1001", f"Expected PRD-1001, got: {visible_id}"
        print("  ✅ Visible product ID: PRD-1001")

        hidden_id = page.locator("#bomProductId").input_value()
        assert hidden_id == "PRD-1001", f"Hidden productId should be PRD-1001, got: {hidden_id}"
        print("  ✅ Hidden product ID set (edit mode confirmed)")

        product_name = page.locator("#bomProductName").input_value()
        assert product_name == "Maharaja Super Cruiser 24 Speed", f"Product name mismatch: {product_name}"
        print("  ✅ Product name pre-filled correctly")

        remarks = page.locator("#bomRemarks").input_value()
        assert "Premium variant" in remarks, f"Remarks mismatch: {remarks}"
        print("  ✅ Remarks pre-filled correctly")

        # 2026-07-22: edit-mode footer switched from a single "Update Product
        # BOM" button to Save (stays open) + Exit (App.Nav.exit) — see
        # verify_item_bom_process_save_exit.py for that behavior's coverage.
        submit_text = page.locator("#bomSubmitBtn").inner_text()
        assert submit_text.strip() == "Save", f"Expected 'Save' button text in edit mode, got: {submit_text}"
        print("  ✅ Submit button says 'Save'")
        exit_visible = page.locator("#bomExitBtn").is_visible()
        assert exit_visible, "Expected Exit button visible in edit mode"
        print("  ✅ Exit button visible")

        # ─────────────────────────────────────────────────────────────
        # SCENARIO 2: Multi-group rendering
        # ─────────────────────────────────────────────────────────────
        print("\n[Scenario 2] Verify multi-group rendering...")
        group_cards = page.locator("#bomGroupsContainer .bom-group-card")
        count = group_cards.count()
        print(f"  Group cards found: {count}")
        assert count == 2, f"Expected 2 group cards, got: {count}"
        print("  ✅ 2 process groups rendered")

        group_names = [
            group_cards.nth(0).locator(".group-name-input option:checked").inner_text(),
            group_cards.nth(1).locator(".group-name-input option:checked").inner_text()
        ]
        print(f"  Group names: {group_names}")
        assert "Frame Assembly" in group_names, f"Frame Assembly not found: {group_names}"
        assert "Fitting & Finishing" in group_names, f"Fitting & Finishing not found: {group_names}"
        print("  ✅ Both group names rendered correctly")

        for i in range(2):
            rows = group_cards.nth(i).locator(".bom-group-items-body tr")
            row_count = rows.count()
            assert row_count == 1, f"Group {i+1}: expected 1 row, got {row_count}"
        print("  ✅ Each group has exactly 1 component row")

        # ─────────────────────────────────────────────────────────────
        # SCENARIO 3: Additional costs pre-fill
        # ─────────────────────────────────────────────────────────────
        print("\n[Scenario 3] Verify additional costs pre-fill...")
        cost_rows = page.locator("#bomCostsBody tr")
        cost_count = cost_rows.count()
        print(f"  Additional cost rows: {cost_count}")
        assert cost_count == 2, f"Expected 2 additional cost rows, got: {cost_count}"

        cost1_desc = cost_rows.nth(0).locator(".cost-description").input_value()
        cost1_rate = cost_rows.nth(0).locator(".cost-rate").input_value()
        assert cost1_desc == "Labor - Fitting", f"Cost 1 description mismatch: {cost1_desc}"
        assert float(cost1_rate) == 50, f"Cost 1 rate mismatch: {cost1_rate}"
        print(f"  ✅ Cost 1: '{cost1_desc}' @ ₹{cost1_rate}")

        cost2_desc = cost_rows.nth(1).locator(".cost-description").input_value()
        cost2_rate = cost_rows.nth(1).locator(".cost-rate").input_value()
        assert cost2_desc == "Paint Cost", f"Cost 2 description mismatch: {cost2_desc}"
        assert float(cost2_rate) == 75, f"Cost 2 rate mismatch: {cost2_rate}"
        print(f"  ✅ Cost 2: '{cost2_desc}' @ ₹{cost2_rate}")

        # ─────────────────────────────────────────────────────────────
        # SCENARIO 4: serializeForm includes productId (edit mode)
        # ─────────────────────────────────────────────────────────────
        # Api is a closure variable, not on window — so we check the
        # serialized formData directly via App.BOM.serializeForm()
        # rather than intercepting the API layer.
        print("\n[Scenario 4] Verify serializeForm includes productId (edit mode)...")

        # Change product name first so we can confirm the update is captured
        name_input = page.locator("#bomProductName")
        name_input.click(click_count=3)
        name_input.type("Maharaja Super Cruiser 26 Speed")

        captured = page.evaluate("""
            (() => {
                const { formData, componentCount } = App.BOM.serializeForm();
                return { formData, componentCount };
            })()
        """)

        form_data = captured["formData"]
        assert form_data.get("productId") == "PRD-1001", \
            f"productId should be PRD-1001 in edit mode, got: {form_data.get('productId')}"
        print(f"  ✅ formData.productId = '{form_data['productId']}' (edit mode confirmed)")

        captured_name = form_data.get("productName", "")
        assert "26 Speed" in captured_name, f"Updated name not in formData: {captured_name}"
        print(f"  ✅ Updated product name in formData: '{captured_name}'")

        captured_components = json.loads(form_data.get("components", "[]"))
        assert len(captured_components) == 2, f"Expected 2 components, got: {len(captured_components)}"
        print(f"  ✅ {len(captured_components)} components serialized")

        captured_costs = json.loads(form_data.get("additionalCosts", "[]"))
        assert len(captured_costs) == 2, f"Expected 2 additional costs, got: {len(captured_costs)}"
        print(f"  ✅ {len(captured_costs)} additional costs serialized")

        assert captured["componentCount"] == 2, \
            f"componentCount should be 2, got: {captured['componentCount']}"
        print(f"  ✅ componentCount = {captured['componentCount']}")

        # ─────────────────────────────────────────────────────────────
        # PROBE: Orphan item (deleted from Items Master) is preserved
        # ─────────────────────────────────────────────────────────────
        print("\n🔍 [Probe] Orphan item handling (item not in Items Master)...")

        # Close the current modal first
        page.evaluate("""
            bootstrap.Modal.getOrCreateInstance(document.getElementById('editBomModal')).hide();
        """)
        page.wait_for_timeout(500)

        orphan_bom = {
            "productId": "PRD-9999",
            "productName": "Orphan Test Product",
            "remarks": "",
            "components": [
                {
                    "itemName": "DELETED-ITEM",
                    "size": "Special",
                    "narration": "Was removed from Items Master",
                    "rate": 99,
                    "vendor": "Old Vendor",
                    "qtyPerProduct": 5,
                    "lineCost": 495,
                    "processId": "",
                    "processGroup": "General"
                }
            ],
            "totalCost": 495,
            "totalQty": 5,
            "additionalCosts": [],
            "totalAdditionalCost": 0,
            "grandTotal": 495
        }

        page.evaluate(f"""
            App.State.globalBOMs = [{json.dumps(orphan_bom)}];
            App.BOM.openEditModal(0);
        """)
        modal.wait_for(state="visible", timeout=TIMEOUT)
        page.wait_for_timeout(500)

        orphan_count = page.evaluate("""
            document.querySelectorAll('#bomGroupsContainer option[value="orphan"]').length
        """)
        print(f"  Orphan option count: {orphan_count}")
        assert orphan_count == 1, f"Expected 1 orphan option for deleted item, got: {orphan_count}"

        orphan_label = page.evaluate("""
            document.querySelector('#bomGroupsContainer option[value="orphan"]').textContent
        """)
        print(f"  Orphan label: '{orphan_label}'")
        assert "Not in Items Master" in orphan_label, f"Orphan label missing expected text: {orphan_label}"
        assert "DELETED-ITEM" in orphan_label, f"Orphan label missing item name: {orphan_label}"
        print("  ✅ Orphan item preserved with 'Not in Items Master' label")

        # ─────────────────────────────────────────────────────────────
        print("\n✅ ALL SCENARIOS PASSED")
        browser.close()
        return True


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
