/**
 * Standalone Node harness (same mock-SpreadsheetApp pattern as
 * test_color_axes.js) covering the Phase 3 "Ask B" fix: a process
 * assembling 2+ independently-tracked pool inputs with no Process Color
 * Link between them (e.g. Fitted Frame Assembly consuming a Painted Frame
 * axis + a Fitted Rim axis) used to credit its own Warehouse Pool output as
 * TWO separate single-color buckets ("Red-White":10 and "Black":10)
 * instead of ONE real combined bucket ("Black / Red-White":10) — there was
 * no way to answer "how many Black+Red-White units are in stock" from the
 * data as stored. recalculateWarehousePool's Pass 1 (module_warehouse.js)
 * now combines a lot's colorBreakdown entries into one bucket whenever the
 * pairing is unambiguous, using the new server-side _colorNamesMatch
 * (module_process.js) to auto-exclude a redundant, name-matched axis (e.g.
 * Mudguard Color, which matches Frame Color 1:1 by name) from the combined
 * key, and falling back to today's exact per-entry crediting whenever the
 * pairing isn't clean-cut (never guessing at a quantity attribution).
 *
 * Run: node .pw-test/test_warehouse_pool_composite_bucket_credit.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet; this.row = row; this.col = col; this.numRows = numRows; this.numCols = numCols;
  }
  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowArr = [];
      for (let c = 0; c < this.numCols; c++) rowArr.push(this.sheet._get(this.row + r, this.col + c));
      out.push(rowArr);
    }
    return out;
  }
  getValue() { return this.sheet._get(this.row, this.col); }
  setValues(values) {
    values.forEach((rowArr, r) => rowArr.forEach((val, c) => this.sheet._set(this.row + r, this.col + c, val)));
    return this;
  }
  setValue(v) { this.sheet._set(this.row, this.col, v); return this; }
  clearContent() {
    for (let r = 0; r < this.numRows; r++) for (let c = 0; c < this.numCols; c++) this.sheet._set(this.row + r, this.col + c, '');
    return this;
  }
  setFontWeight() { return this; }
  setBackground() { return this; }
}

class FakeSheet {
  constructor(name) { this.name = name; this.rows = []; }
  getName() { return this.name; }
  _ensureRow(r) { while (this.rows.length < r) this.rows.push([]); }
  _get(r, c) { this._ensureRow(r); const row = this.rows[r - 1]; return row[c - 1] === undefined ? '' : row[c - 1]; }
  _set(r, c, v) { this._ensureRow(r); const row = this.rows[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; }
  getLastRow() {
    for (let r = this.rows.length; r >= 1; r--) {
      if (this.rows[r - 1].some(v => v !== '' && v !== undefined && v !== null)) return r;
    }
    return 0;
  }
  getLastColumn() {
    let max = 0;
    this.rows.forEach(row => {
      for (let c = row.length; c >= 1; c--) {
        if (row[c - 1] !== '' && row[c - 1] !== undefined && row[c - 1] !== null) { max = Math.max(max, c); break; }
      }
    });
    return max;
  }
  getRange(row, col, numRows = 1, numCols = 1) { return new FakeRange(this, row, col, numRows, numCols); }
  appendRow(arr) { const r = this.getLastRow() + 1; arr.forEach((v, i) => this._set(r, i + 1, v)); }
  deleteRow(r) { this.rows.splice(r - 1, 1); }
  deleteRows(r, n) { this.rows.splice(r - 1, n); }
  insertRows(r, n) { for (let i = 0; i < n; i++) this.rows.splice(r - 1, 0, []); }
  insertColumnsAfter(afterPosition, howMany) {
    this.rows.forEach(row => { const blanks = new Array(howMany).fill(''); row.splice(afterPosition, 0, ...blanks); });
  }
}

class FakeSpreadsheet {
  constructor() { this.sheets = {}; }
  getSheetByName(name) { return this.sheets[name] || null; }
  addSheet(name) { const s = new FakeSheet(name); this.sheets[name] = s; return s; }
  insertSheet(name) { return this.addSheet(name); }
}

const ss = new FakeSpreadsheet();

const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush: () => {} },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  console,
  Logger: { log: () => {} },
  Utilities: { getUuid: () => 'uuid-' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);

['config.js', 'utils.js', 'module_process.js', 'module_production.js', 'module_warehouse.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});

vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.COLOR_COMBO_DELIMITER = COLOR_COMBO_DELIMITER;
`, ctx, { filename: 'expose.js' });

const {
  APP_CONFIG, COLOR_COMBO_DELIMITER,
  saveProcess, saveProduction, recalculateWarehousePool, getWarehousePoolData
} = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); }
}

function bucketsFor(outputItemName) {
  const res = getWarehousePoolData();
  return (res.data || []).filter(r => r.outputItemName === outputItemName);
}

console.log('\n=== Setup: Fitted Frame Assembly process (no Process Color Link between its axes) ===');
const procRes = saveProcess({
  processName: 'Fitted Frame Assembly',
  sequence: 1,
  lotPrefix: 'FFA',
  outputItemName: 'Fitted Frame Assembled',
  isFinalStage: false,
  active: true,
  remarks: '',
  components: JSON.stringify([])
});
assert(procRes.success, 'saveProcess succeeds: ' + procRes.message);
const processId = procRes.data && procRes.data.processId;
assert(!!processId, 'processId returned');

function saveLot(colorBreakdown) {
  return saveProduction({
    processId,
    assignedTo: 'Test Contractor',
    status: 'Completed',
    colorBreakdown: JSON.stringify(colorBreakdown),
    componentsConsumed: JSON.stringify([
      { itemName: 'Assembly Hardware', sourceType: 'ITEM', qty: 1, colorGroup: 'COMMON' }
    ])
  });
}

console.log('\n=== Test 1: Frame + Rim (2 independent axes, no link) combine into ONE bucket ===');
{
  const res = saveLot([
    { color: 'Red-White', qty: 10, countsTowardTotal: true, isCustom: true },
    { color: 'Black', qty: 10, countsTowardTotal: false, isCustom: true }
  ]);
  assert(res.success, 'saveProduction succeeds: ' + res.message);
  recalculateWarehousePool();

  const buckets = bucketsFor('Fitted Frame Assembled');
  const combined = buckets.find(b => b.color === `Black${COLOR_COMBO_DELIMITER}Red-White`);
  assert(!!combined && combined.producedQty === 10, `ONE combined "Black / Red-White" bucket with producedQty 10 (got ${JSON.stringify(combined)})`);
  assert(!buckets.some(b => b.color === 'Red-White'), 'no separate standalone "Red-White" bucket exists');
  assert(!buckets.some(b => b.color === 'Black'), 'no separate standalone "Black" bucket exists');
  assert(buckets.length === 1, `exactly one bucket total for this process so far (got ${buckets.length}: ${JSON.stringify(buckets.map(b => b.color))})`);
}

console.log('\n=== Test 2: a redundant, name-matching axis (Mudguard "Red" matches "Red-White") is excluded from the key ===');
{
  const res = saveLot([
    { color: 'Red-White', qty: 5, countsTowardTotal: true, isCustom: true },
    { color: 'Black', qty: 5, countsTowardTotal: false, isCustom: true },
    { color: 'Red', qty: 5, countsTowardTotal: false, isCustom: true } // Mudguard, name-matches Red-White -> redundant
  ]);
  assert(res.success, 'saveProduction (with redundant Mudguard entry) succeeds: ' + res.message);
  recalculateWarehousePool();

  const buckets = bucketsFor('Fitted Frame Assembled');
  const combined = buckets.find(b => b.color === `Black${COLOR_COMBO_DELIMITER}Red-White`);
  assert(!!combined && combined.producedQty === 15, `combined "Black / Red-White" bucket accumulated to 15 (10 from Test 1 + 5 here) (got ${JSON.stringify(combined)})`);
  assert(!buckets.some(b => b.color === 'Red'), 'no separate "Red" (Mudguard) bucket was created - redundant entry silently folded in, not separately credited');
  assert(!buckets.some(b => (b.color || '').includes('Red-White') && (b.color || '').includes('Red') && (b.color || '').split(COLOR_COMBO_DELIMITER).length > 2),
    'no 3-way combo was created either (Mudguard must not add a 3rd dimension)');
}

console.log('\n=== Test 3: 2 independent (non-matching) otherEntries -> falls back to separate per-entry credits, no guessing ===');
{
  const res = saveLot([
    { color: 'Blue-White', qty: 8, countsTowardTotal: true, isCustom: true },
    { color: 'Black', qty: 8, countsTowardTotal: false, isCustom: true },  // independent (no name match to Blue-White)
    { color: 'Green', qty: 8, countsTowardTotal: false, isCustom: true }   // ALSO independent - ambiguous which pairs with Blue-White
  ]);
  assert(res.success, 'saveProduction (2 independent otherEntries) succeeds: ' + res.message);
  recalculateWarehousePool();

  const buckets = bucketsFor('Fitted Frame Assembled');
  const blueWhite = buckets.find(b => b.color === 'Blue-White');
  const black = buckets.find(b => b.color === 'Black');
  const green = buckets.find(b => b.color === 'Green');
  assert(!!blueWhite && blueWhite.producedQty === 8, `standalone "Blue-White" bucket credited 8 - fallback, unchanged from pre-fix behavior (got ${JSON.stringify(blueWhite)})`);
  assert(!!black && black.producedQty === 8, `standalone "Black" bucket credited 8 separately (got ${JSON.stringify(black)})`);
  assert(!!green && green.producedQty === 8, `standalone "Green" bucket credited 8 separately (got ${JSON.stringify(green)})`);
  assert(!buckets.some(b => (b.color || '').includes(COLOR_COMBO_DELIMITER) && (b.color || '').includes('Blue-White')),
    'no composite bucket was guessed for the ambiguous Blue-White/Black/Green combination');
}

console.log('\n=== Test 4: a single-entry lot (the common case) is byte-identical to pre-fix behavior ===');
{
  const res = saveLot([
    { color: 'Solo', qty: 7, isCustom: true } // no countsTowardTotal field at all - defaults true
  ]);
  assert(res.success, 'saveProduction (single entry) succeeds: ' + res.message);
  recalculateWarehousePool();

  const buckets = bucketsFor('Fitted Frame Assembled');
  const solo = buckets.find(b => b.color === 'Solo');
  assert(!!solo && solo.producedQty === 7, `standalone "Solo" bucket credited 7, unchanged single-axis behavior (got ${JSON.stringify(solo)})`);
}

console.log('\n' + (failures === 0 ? 'ALL TESTS PASSED' : failures + ' TEST(S) FAILED'));
process.exit(failures === 0 ? 0 : 1);
