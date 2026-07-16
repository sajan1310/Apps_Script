"""
Verification script: Item Reference Integrity dialog's "Repoint To" Select2
dropdown, scrolling behavior.

Covers TWO real bugs found in this dialog, in order:

Bug 1 (fixed first): dropdownParent was the outer `.modal` element, but this
app's global CSS puts the actual scroll container (overflow-y:auto,
max-height:82vh) on `.modal-body` — a descendant, not an ancestor Select2's
scroll tracking would see. An open dropdown stayed frozen in its original
screen position while its trigger scrolled away underneath it.

Bug 2 (found via a user screenshot, after fixing bug 1 by pointing
dropdownParent at #itemDriftModalBody): Select2's position math for a
SCROLLABLE dropdownParent is only reliable while a dropdown that was ALREADY
open tracks a scroll. Opening a dropdown for the FIRST TIME on a row that only
became visible by scrolling computed a badly wrong position (observed
rendering at y:-234, i.e. off-screen above the modal, in a table scrolled a
few hundred px down) — a known Select2 limitation with scrollable
dropdownParents. Fixed by switching dropdownParent to document.body (never
scrolled while a Bootstrap modal is open, so Select2's math is always
correct) and closing (not repositioning) any open dropdown when the row list
scrolls, via a scroll listener on #itemDriftModalBody.
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000

MOCK_ITEMS = [{"name": f"Item {i:04d}", "size": "General"} for i in range(50)]
MOCK_DRIFT = [
    {"sheet": "Wastage Log", "context": f"W-{g}", "itemName": f"Stale Item {g}", "size": "500ml"}
    for g in range(20)
]


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1400, "height": 700})

        console_errors = []
        page.on("pageerror", lambda exc: console_errors.append(f"pageerror: {exc}"))

        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(1000)
        page.evaluate(f"""
            App.State.globalItems = {json.dumps(MOCK_ITEMS)};
            Api.call = async (fn) => ({{ success: true, data: {json.dumps(MOCK_DRIFT)}, message: '' }});
        """)
        page.evaluate("App.Item.checkReferenceIntegrity()")
        page.locator("#itemDriftModal").wait_for(state="visible", timeout=TIMEOUT)
        page.locator("#itemDriftModalBody .select2-container").first.wait_for(state="visible", timeout=TIMEOUT)

        print("[1] Opening a dropdown, then scrolling, closes it (doesn't detach it)...")
        first = page.locator("#itemDriftModalBody .select2-container").first
        first.click()
        page.wait_for_timeout(250)
        assert page.locator(".select2-dropdown").count() == 1, "dropdown should be open before scrolling"
        page.evaluate("document.getElementById('itemDriftModalBody').scrollTop += 300")
        page.wait_for_timeout(250)
        assert page.locator(".select2-dropdown").count() == 0, "dropdown should auto-close on scroll, not linger detached"
        print("  ✅ Dropdown closed cleanly on scroll")

        page.evaluate("document.getElementById('itemDriftModalBody').scrollTop = 0")
        page.wait_for_timeout(200)

        print("[2] Scrolling first, then opening a row that only became visible via scroll, positions correctly...")
        page.evaluate("document.getElementById('itemDriftModalBody').scrollTop = 400")
        page.wait_for_timeout(300)

        body_box = page.evaluate("""() => {
            const r = document.getElementById('itemDriftModalBody').getBoundingClientRect();
            return { top: r.top, bottom: r.bottom };
        }""")
        rows = page.locator("#itemDriftModalBody .select2-container")
        visible_idx = None
        for i in range(rows.count()):
            box = rows.nth(i).bounding_box()
            if box and body_box["top"] <= box["y"] <= body_box["bottom"]:
                visible_idx = i
                break
        assert visible_idx is not None, "expected at least one row visible after scrolling"

        target = rows.nth(visible_idx)
        tbox = target.bounding_box()
        target.click()
        page.wait_for_timeout(300)

        dd_box = page.locator(".select2-dropdown").bounding_box()
        assert dd_box is not None, "dropdown should open for the newly-visible row"
        gap = dd_box["y"] - (tbox["y"] + tbox["height"])
        assert abs(gap) < 5, f"dropdown mispositioned relative to its trigger after scroll (gap={gap:.1f}px, expected ~0) — this is the exact bug from the user's screenshot"
        print(f"  ✅ Dropdown correctly aligned to its trigger (gap={gap:.2f}px)")

        if console_errors:
            print("\n[Console/page errors]")
            for e in console_errors:
                print("  -", e)
            raise AssertionError("page errors occurred during test")

        print("\nALL CHECKS PASSED")
        browser.close()


if __name__ == "__main__":
    run()
