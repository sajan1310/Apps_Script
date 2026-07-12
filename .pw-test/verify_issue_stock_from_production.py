"""
Verification script: "Issue Stock" button inside the Log Production Lot form.

Opens the Production modal, then clicks its "Issue Stock" button (which
calls App.Production.openIssueStockForLot() -> App.Issue.openIssueModal()),
and checks whether the resulting Issue Stock modal actually renders on top
and is interactive, or whether it's stacked BEHIND the still-open Production
modal (a known Bootstrap 5 limitation for simultaneously-open modals that
this codebase already worked around once, in App.Utils.confirmAction —
see Script.html's "does not support nested modals natively" comment).
"""

import sys
import io
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"
TIMEOUT = 8000


def run():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context()
        page = ctx.new_page()

        url = DIST_HTML.as_uri()
        page.goto(url, wait_until="domcontentloaded")
        page.wait_for_timeout(1000)

        print("\n[Step 1] Open the Log Production Lot modal...")
        page.evaluate("App.Production.openCreateModal()")
        prod_modal = page.locator("#editProductionModal")
        prod_modal.wait_for(state="visible", timeout=TIMEOUT)
        print("  Production modal is visible")

        print("\n[Step 2] Click 'Issue Stock' from inside the Production form...")
        page.evaluate("App.Production.openIssueStockForLot()")
        issue_modal = page.locator("#issueStockModal")
        issue_modal.wait_for(state="visible", timeout=TIMEOUT)
        print("  Issue Stock modal has the 'show' class / is technically visible")

        # ── The real check: is the Issue Stock modal actually ON TOP, or is
        # it stacked behind the still-open Production modal? Bootstrap gives
        # every .modal the same z-index (1055), so ties are broken by DOM
        # order, not by which was opened last.
        info = page.evaluate("""
            (() => {
                const prod = document.getElementById('editProductionModal');
                const issue = document.getElementById('issueStockModal');
                const prodZ = getComputedStyle(prod).zIndex;
                const issueZ = getComputedStyle(issue).zIndex;
                const domOrder = prod.compareDocumentPosition(issue);
                // Bit 4 (0x04) = issue FOLLOWS prod in the DOM (renders on top at equal z-index)
                const issueIsLaterInDom = !!(domOrder & Node.DOCUMENT_POSITION_FOLLOWING) === false
                    ? !!(prod.compareDocumentPosition(issue) & Node.DOCUMENT_POSITION_PRECEDING)
                    : false;
                const issueRect = issue.querySelector('.modal-dialog').getBoundingClientRect();
                const cx = issueRect.left + issueRect.width / 2;
                const cy = issueRect.top + 40; // near the header, inside the dialog
                const topEl = document.elementFromPoint(cx, cy);
                const topElIsInsideIssueModal = issue.contains(topEl);
                const topElDescription = topEl ? (topEl.id ? '#' + topEl.id : topEl.className) : null;
                return { prodZ, issueZ, issueIsLaterInDom, topElIsInsideIssueModal, topElDescription };
            })()
        """)
        print(f"  Production modal z-index: {info['prodZ']}")
        print(f"  Issue modal z-index:      {info['issueZ']}")
        print(f"  Issue modal appears AFTER Production modal in DOM: {info['issueIsLaterInDom']}")
        print(f"  Element actually on top at Issue modal's header point: {info['topElDescription']}")
        print(f"  That top element is INSIDE the Issue modal: {info['topElIsInsideIssueModal']}")

        if not info["topElIsInsideIssueModal"]:
            print("\n❌ BUG CONFIRMED: the Issue Stock modal is rendered behind the still-open")
            print("   Production modal — clicking into it actually hits the Production modal instead.")
        else:
            print("\n✅ Issue Stock modal is on top and interactive.")

        # Can we actually type into the Issue modal's reference field?
        ref_input = page.locator("#issueReferenceInput")
        clickable = False
        try:
            ref_input.click(timeout=2000)
            clickable = True
        except Exception as e:
            print(f"  Click on #issueReferenceInput failed: {e}")
        print(f"\n[Step 3] Can click into Issue Stock modal's Reference field: {clickable}")

        browser.close()
        return info["topElIsInsideIssueModal"] and clickable


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
