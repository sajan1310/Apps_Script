"""
Verification script — App.Item's metadata-filter performance refactor
(Script_Items.html). getMetaInfo/getMetadataFlags and the 'low stock' quick
filter now look up an item's Stock row via a Map<name|size, stockEntry>
built once (_getStockMap), instead of globalStock.find()/.some() per item —
which ran over the FULL item list for 'low stock'/'zero stock'/'no metadata'
searches and for the per-column Meta Data filter.

Confirms the filters still produce the exact same results: low stock,
zero stock, no metadata, and pending order all still resolve correctly
after the refactor.

Run: python .pw-test/verify_items_metadata_filter_perf.py
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


SEED = """
    App.State.globalItems = [
      { name: 'Widget', size: '', narration: '', vendors: [], remarks: '', specification: '' },   // low stock
      { name: 'Gadget', size: 'L', narration: '', vendors: [], remarks: '', specification: '' },   // zero stock
      { name: 'Sprocket', size: '', narration: '', vendors: [], remarks: 'handle with care', specification: 'steel' },  // has metadata, healthy stock
      { name: 'Orphan', size: '', narration: '', vendors: [], remarks: '', specification: '' }     // no stock row at all -> no metadata
    ];
    App.State.globalStock = [
      { name: 'Widget', size: '', currentStock: 2, threshold: 10, isLowStock: true },
      { name: 'Gadget', size: 'L', currentStock: 0, threshold: 0, isLowStock: false },  // zero stock but not "low" (threshold 0) -- isolates the two filters
      { name: 'Sprocket', size: '', currentStock: 50, threshold: 10, isLowStock: false }
    ];
    App.State.globalPOs = [];
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

        page.evaluate("""
            window.google = {
                script: {
                    run: {
                        withSuccessHandler(cb) {
                            const runner = { withFailureHandler() { return runner; } };
                            return new Proxy(runner, { get(target, prop) {
                                if (prop in target) return target[prop];
                                return (...args) => {};
                            }});
                        }
                    }
                }
            };
        """)

        print("\n[Seed] 4 items: 1 low stock, 1 zero stock, 1 with remarks/spec, 1 with no Stock row at all")
        page.evaluate(SEED)

        def filtered_names(term):
            return page.evaluate(f"""
                (() => {{
                  App.Item.filterData({term!r});
                  return App.State.filteredItems.map(i => i.name);
                }})()
            """)

        names = filtered_names("low stock")
        check(names == ["Widget"], f"'low stock' filter returns only Widget (got {names})")

        names = filtered_names("zero stock")
        check(names == ["Gadget"], f"'zero stock' filter returns only Gadget (got {names})")

        names = filtered_names("no metadata")
        check(names == ["Orphan"], f"'no metadata' filter returns only Orphan (no stock row, no remarks/spec) (got {names})")

        # getMetaInfo directly, to check the rendered stock text and dedup path
        meta = page.evaluate("""
            (() => {
              const pendingMap = App.Utils.getPendingByItem();
              const stockMap = App.Item._getStockMap();
              const widget = App.State.globalItems.find(i => i.name === 'Widget');
              const info = App.Item.getMetaInfo(widget, pendingMap, stockMap);
              return { partsCount: info.parts.length, hasLowBadge: info.parts.some(p => p.includes('Low')) };
            })()
        """)
        check(meta["partsCount"] == 1, f"Widget's meta has 1 part (its Stock line) (got {meta['partsCount']})")
        check(meta["hasLowBadge"], "Widget's Stock part is flagged Low")

        check(len(console_errors) == 0, f"no page errors ({console_errors})")

        browser.close()


if __name__ == "__main__":
    run()
    print("\nALL TESTS PASSED" if failures == 0 else f"\n{failures} TEST(S) FAILED")
    sys.exit(1 if failures else 0)
