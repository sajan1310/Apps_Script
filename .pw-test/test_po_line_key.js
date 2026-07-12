/**
 * Standalone Node harness verifying the shared _buildPoLineKey() helper
 * (module_bill.js, next to _aggregateBilledBaseQtyByPo) — extracted so the
 * 4 previously hand-built copies of this key (module_bill.js's aggregator,
 * module_po.js's _attachPoStatus and suggestPoAllocations, module_dashboard.js's
 * _getOpenPoSummary) can never drift out of sync. Covers the edge inputs
 * called out in the fix: null size, undefined narration, mixed case, and
 * padded whitespace must all still produce a matching key.
 *
 * Run: node .pw-test/test_po_line_key.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

const sandbox = { console };
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);

// _buildPoLineKey has no dependency on getBillData()/sheets — but it lives in
// module_bill.js, so load that file (and utils.js, which it references
// elsewhere) to get the real function without needing a fake sheet at all.
['utils.js', 'module_bill.js'].forEach(f =>
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f }));

const { _buildPoLineKey } = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); }
  else console.log('PASS:', msg);
}

const baseline = _buildPoLineKey('PO-1001', 'Widget', '110mm', 'For export batch');
assert(baseline === 'po-1001|widget|110mm|for export batch', `baseline key shape (got ${JSON.stringify(baseline)})`);

// Mixed case + padded whitespace must normalize to the same key.
const padded = _buildPoLineKey('  PO-1001 ', ' Widget ', ' 110mm ', ' For Export Batch ');
assert(padded === baseline, `mixed case + padding normalizes to the same key (got ${JSON.stringify(padded)})`);

const mixedCase = _buildPoLineKey('po-1001', 'WIDGET', '110MM', 'FOR EXPORT BATCH');
assert(mixedCase === baseline, `mixed case alone normalizes to the same key (got ${JSON.stringify(mixedCase)})`);

// Missing narration (undefined) must key the same as an explicit empty string
// — the exact "undefined" literal leaking into the key that motivated this fix.
const undefinedNarration = _buildPoLineKey('PO-1002', 'Widget', '', undefined);
const emptyNarration = _buildPoLineKey('PO-1002', 'Widget', '', '');
assert(undefinedNarration === emptyNarration, `undefined narration keys the same as '' (got ${JSON.stringify(undefinedNarration)} vs ${JSON.stringify(emptyNarration)})`);
assert(!undefinedNarration.includes('undefined'), `key never contains the literal "undefined" (got ${JSON.stringify(undefinedNarration)})`);

// null size must key the same as an explicit empty string.
const nullSize = _buildPoLineKey('PO-1003', 'Widget', null, 'note');
const emptySize = _buildPoLineKey('PO-1003', 'Widget', '', 'note');
assert(nullSize === emptySize, `null size keys the same as '' (got ${JSON.stringify(nullSize)} vs ${JSON.stringify(emptySize)})`);
assert(!nullSize.includes('null'), `key never contains the literal "null" (got ${JSON.stringify(nullSize)})`);

// Every field null/undefined at once must not throw and must key identically
// to all-empty-string.
const allMissing = _buildPoLineKey(undefined, null, undefined, null);
const allEmpty = _buildPoLineKey('', '', '', '');
assert(allMissing === allEmpty, `all-missing fields key the same as all-empty (got ${JSON.stringify(allMissing)})`);
assert(allMissing === '|||', `all-empty key is the bare delimiter shape (got ${JSON.stringify(allMissing)})`);

console.log(failures === 0 ? '\nALL TESTS PASSED' : `\n${failures} TEST(S) FAILED`);
process.exitCode = failures === 0 ? 0 : 1;
