"""
Verifies the new same-process axis-pairing capability added to the Process
editor's "Linked Processes" section (Script_Process.html) - closing the gap
where manually-tagged Color Axes (e.g. Rim Color <-> Mudguard Color) had no
explicit pairing UI, unlike pool-derived axes which already had one via
cross-process Process Color Links. Picking "This process" as the linked
"process" now shows two axis pickers instead of a flat color list, and maps
just those two axes' own colors - not the whole flat getProcessColorGroups
union, which would collide two axes sharing a literal color name.

Run: python .pw-test/verify_process_color_links_self_axis_pairing.py
"""
import sys
import io
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
DIST_HTML = Path(__file__).parent.parent / "dist" / "index.html"

MOCK_AXES = {
    "success": True,
    "data": {
        "axes": [
            {"key": "tag:rim color", "label": "Rim Color", "colors": ["Blue-White", "Red-White"], "source": "tag"},
            {"key": "tag:mudguard color", "label": "Mudguard Color", "colors": ["Blue", "Red"], "source": "tag"}
        ],
        "primaryColorAxis": "Rim Color",
        "primaryAxisKey": "tag:rim color"
    }
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
    global failures
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        page = browser.new_page()
        errors = []
        page.on("pageerror", lambda e: errors.append(str(e)))
        page.goto(DIST_HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(500)

        page.evaluate(f"""
            App.State.globalProcesses = [{{ processId: 'PRC-1', processName: 'Fitted Frame Assembly' }}];
            window.__mockResponses = {{ getProcessColorAxes: {json.dumps(MOCK_AXES)} }};
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
            document.getElementById('processFormProcessId').value = 'PRC-1';
            document.getElementById('processColorLinksContainer').innerHTML = '';
        """)

        print("[1] Add a blank Linked Process card - defaults to 'Choose a process...'")
        page.evaluate("App.Process.addColorLinkCard()")
        page.wait_for_timeout(100)
        selectedText = page.evaluate("""
            document.querySelector('#processColorLinksContainer .proc-colorlink-process-select').selectedOptions[0].textContent
        """)
        check("Choose a process" in selectedText, f"blank card defaults to 'Choose a process...' (got '{selectedText}')")
        selfWrapperHidden = page.evaluate("""
            document.querySelector('#processColorLinksContainer .proc-colorlink-self-axes-wrapper').style.display === 'none'
        """)
        check(selfWrapperHidden, "self-axes picker row stays hidden until the operator actually selects 'This process'")

        print("\n[2] Select 'This process' - self-axes picker row appears, axes load")
        page.evaluate("""
            const sel = document.querySelector('#processColorLinksContainer .proc-colorlink-process-select');
            sel.value = '__self__';
            sel.dispatchEvent(new Event('change'));
        """)
        page.wait_for_timeout(200)
        wrapperVisible = page.evaluate("""
            document.querySelector('#processColorLinksContainer .proc-colorlink-self-axes-wrapper').style.display !== 'none'
        """)
        check(wrapperVisible, "self-axes picker row becomes visible")
        axisOptionCount = page.evaluate("""
            document.querySelector('#processColorLinksContainer .proc-colorlink-my-axis').options.length
        """)
        check(axisOptionCount == 3, f"My Axis picker offers 'Choose an axis...' + 2 real axes (got {axisOptionCount} options)")

        print("\n[3] Pick Rim Color <-> Mudguard Color - mapping table renders with each axis's OWN colors")
        page.evaluate("""
            const myAxis = document.querySelector('#processColorLinksContainer .proc-colorlink-my-axis');
            const theirAxis = document.querySelector('#processColorLinksContainer .proc-colorlink-their-axis');
            myAxis.value = 'tag:rim color';
            myAxis.dispatchEvent(new Event('change'));
            theirAxis.value = 'tag:mudguard color';
            theirAxis.dispatchEvent(new Event('change'));
        """)
        page.wait_for_timeout(150)
        rowCount = page.evaluate("document.querySelectorAll('#processColorLinksContainer .proc-colorlink-mapping-body tr').length")
        check(rowCount == 2, f"mapping table has one row per Rim Color's own 2 colors, not the flat 4-color union (got {rowCount})")
        firstRowColor = page.evaluate("document.querySelector('#processColorLinksContainer .proc-colorlink-my-color').dataset.color")
        check(firstRowColor in ("Blue-White", "Red-White"), f"row's 'My Color' comes from Rim Color axis, not Mudguard (got '{firstRowColor}')")

        print("\n[4] Map Red-White -> Red and Blue-White -> Blue, then serialize")
        page.evaluate("""
            document.querySelectorAll('#processColorLinksContainer .proc-colorlink-mapping-body tr').forEach(row => {
                const myColor = row.querySelector('.proc-colorlink-my-color').dataset.color;
                const sel = row.querySelector('.proc-colorlink-their-color');
                const target = myColor.startsWith('Red') ? 'Red' : 'Blue';
                Array.from(sel.options).forEach(o => { if (o.value === target) sel.value = target; });
            });
        """)
        serialized = page.evaluate("App.Process.serializeColorLinks()")
        check(len(serialized) == 2, f"serializeColorLinks returns 2 pairs (got {len(serialized)}): {serialized}")
        check(all(l['otherProcessId'] == 'PRC-1' for l in serialized), "every pair's otherProcessId resolves to THIS process's own ID (same-process link)")
        check(all(l['myAxisKey'] == 'tag:rim color' and l['theirAxisKey'] == 'tag:mudguard color' for l in serialized),
              "every pair carries the chosen axis keys (myAxisKey='tag:rim color', theirAxisKey='tag:mudguard color')")
        redWhite = next((l for l in serialized if l['myColor'] == 'Red-White'), None)
        check(redWhite is not None and redWhite['theirColor'] == 'Red', f"Red-White -> Red mapped correctly (got {redWhite})")

        print("\n[5] Picking the SAME axis on both sides is rejected with a hint, no table shown")
        page.evaluate("""
            const theirAxis = document.querySelector('#processColorLinksContainer .proc-colorlink-their-axis');
            theirAxis.value = 'tag:rim color';
            theirAxis.dispatchEvent(new Event('change'));
        """)
        page.wait_for_timeout(100)
        tableHiddenAfterSameAxis = page.evaluate("""
            document.querySelector('#processColorLinksContainer .proc-colorlink-mapping-wrapper').style.display === 'none'
        """)
        check(tableHiddenAfterSameAxis, "mapping table hides when the same axis is picked on both sides")
        hintText = page.evaluate("document.querySelector('#processColorLinksContainer .proc-colorlink-hint').textContent")
        check("different" in hintText.lower(), f"hint explains the two axes must differ (got '{hintText}')")
        serializedAfterSameAxis = page.evaluate("App.Process.serializeColorLinks()")
        check(len(serializedAfterSameAxis) == 0, f"serializeColorLinks drops the card entirely once both axes match (got {serializedAfterSameAxis})")

        print("\n[6] renderColorLinksData restores a saved same-process link into its own card with axes pre-selected")
        page.evaluate("document.getElementById('processColorLinksContainer').innerHTML = ''")
        page.evaluate("""
            App.Process.renderColorLinksData([
                { otherProcessId: 'PRC-1', otherProcessName: 'Fitted Frame Assembly', myColor: 'Red-White', theirColor: 'Red', myAxisKey: 'tag:rim color', theirAxisKey: 'tag:mudguard color' },
                { otherProcessId: 'PRC-1', otherProcessName: 'Fitted Frame Assembly', myColor: 'Blue-White', theirColor: 'Blue', myAxisKey: 'tag:rim color', theirAxisKey: 'tag:mudguard color' }
            ]);
        """)
        page.wait_for_timeout(250)
        cardCount = page.evaluate("document.querySelectorAll('#processColorLinksContainer .proc-colorlink-card').length")
        check(cardCount == 1, f"both saved pairs (same axis combo) collapse into ONE restored card, not two (got {cardCount})")
        restoredSelectValue = page.evaluate("document.querySelector('#processColorLinksContainer .proc-colorlink-process-select').value")
        check(restoredSelectValue == '__self__', f"restored card shows 'This process' selected (got '{restoredSelectValue}')")
        restoredMyAxis = page.evaluate("document.querySelector('#processColorLinksContainer .proc-colorlink-my-axis').value")
        restoredTheirAxis = page.evaluate("document.querySelector('#processColorLinksContainer .proc-colorlink-their-axis').value")
        check(restoredMyAxis == 'tag:rim color' and restoredTheirAxis == 'tag:mudguard color',
              f"restored card's axis pickers are pre-selected to the saved pair (got my={restoredMyAxis!r}, their={restoredTheirAxis!r})")
        restoredRowCount = page.evaluate("document.querySelectorAll('#processColorLinksContainer .proc-colorlink-mapping-body tr').length")
        check(restoredRowCount == 2, f"restored mapping table shows both saved rows (got {restoredRowCount})")

        check(len(errors) == 0, f"no console/page errors (got {errors if errors else 'NONE'})")

        browser.close()


if __name__ == "__main__":
    run()
    print(f"\n{'ALL PASS' if failures == 0 else str(failures) + ' FAILURE(S)'}")
    sys.exit(0 if failures == 0 else 1)
