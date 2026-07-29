/**
 * Verifies the generalized N-way Pass 1 combining logic in
 * recalculateWarehousePool (module_warehouse.js):
 * 1. A lot with 3 independent axes (Frame=primary, Mudguard, Rim), each
 *    contributing exactly ONE entry, now combines into ONE 3-way composite
 *    bucket "Blue-White / Black / BCP" instead of 3 separate single-axis
 *    buckets.
 * 2. A redundant axis (name-matches the primary) is still excluded from
 *    the combination, same as before.
 * 3. If any ONE axis contributes 2+ entries (ambiguous pairing), the whole
 *    lot falls back to per-entry crediting -- no guessing.
 * 4. The pre-existing 2-axis (primary + 1 independent) case is unchanged.
 *
 * Run: node .pw-test/test_nway_composite_credit.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = 'c:\\Users\\erkar\\my-app-script-project';

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) { this.sheet = sheet; this.row = row; this.col = col; this.numRows = numRows; this.numCols = numCols; }
  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) { const rowArr = []; for (let c = 0; c < this.numCols; c++) rowArr.push(this.sheet._get(this.row + r, this.col + c)); out.push(rowArr); }
    return out;
  }
  getValue() { return this.sheet._get(this.row, this.col); }
  setValues(values) { values.forEach((rowArr, r) => rowArr.forEach((val, c) => this.sheet._set(this.row + r, this.col + c, val))); return this; }
  setValue(v) { this.sheet._set(this.row, this.col, v); return this; }
  clearContent() { for (let r = 0; r < this.numRows; r++) for (let c = 0; c < this.numCols; c++) this.sheet._set(this.row + r, this.col + c, ''); return this; }
  setFontWeight() { return this; } setBackground() { return this; } setNumberFormat() { return this; }
}
class FakeSheet {
  constructor(name) { this.name = name; this.rows = []; }
  getName() { return this.name; }
  _ensureRow(r) { while (this.rows.length < r) this.rows.push([]); }
  _get(r, c) { this._ensureRow(r); const row = this.rows[r - 1]; return row[c - 1] === undefined ? '' : row[c - 1]; }
  _set(r, c, v) { this._ensureRow(r); const row = this.rows[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; }
  getLastRow() { for (let r = this.rows.length; r >= 1; r--) { if (this.rows[r - 1].some(v => v !== '' && v !== undefined && v !== null)) return r; } return 0; }
  getLastColumn() { let max = 0; this.rows.forEach(row => { for (let c = row.length; c >= 1; c--) { if (row[c - 1] !== '' && row[c - 1] !== undefined && row[c - 1] !== null) { max = Math.max(max, c); break; } } }); return max; }
  getRange(row, col, numRows = 1, numCols = 1) { return new FakeRange(this, row, col, numRows, numCols); }
  appendRow(arr) { const r = this.getLastRow() + 1; arr.forEach((v, i) => this._set(r, i + 1, v)); }
  deleteRow(r) { this.rows.splice(r - 1, 1); } deleteRows(r, n) { this.rows.splice(r - 1, n); }
  insertRows(r, n) { for (let i = 0; i < n; i++) this.rows.splice(r - 1, 0, []); }
  insertColumnsAfter(afterPosition, howMany) { this.rows.forEach(row => { const blanks = new Array(howMany).fill(''); row.splice(afterPosition, 0, ...blanks); }); }
}
class FakeSpreadsheet { constructor() { this.sheets = {}; } getSheetByName(name) { return this.sheets[name] || null; } addSheet(name) { const s = new FakeSheet(name); this.sheets[name] = s; return s; } insertSheet(name) { return this.addSheet(name); } }
const ss = new FakeSpreadsheet();
const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush: () => {} },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
  console, Logger: { log: () => {} },
  Utilities: { getUuid: () => 'uuid-' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);
['config.js', 'utils.js', 'module_units.js', 'module_items.js', 'module_process.js', 'module_production.js', 'module_warehouse.js', 'module_stock.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});
vm.runInContext('global.APP_CONFIG=APP_CONFIG; global.PRODUCTION_COL=PRODUCTION_COL;', ctx, { filename: 'expose.js' });
const { APP_CONFIG, PRODUCTION_COL, recalculateWarehousePool, getWarehousePoolData } = ctx;

let failures = 0;
function assert(cond, msg) { if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); } }

const prodSheet = ss.addSheet(APP_CONFIG.SHEETS.PRODUCTION);
function setLot(row, { qty, lotNumber, outputItemName, colorBreakdown }) {
  prodSheet._set(row, PRODUCTION_COL.DATE, '01/01/2026');
  prodSheet._set(row, PRODUCTION_COL.QTY, qty);
  prodSheet._set(row, PRODUCTION_COL.STATUS, 'Completed');
  prodSheet._set(row, PRODUCTION_COL.PROCESS_ID, 'PRC-FIT');
  prodSheet._set(row, PRODUCTION_COL.LOT_NUMBER, lotNumber);
  prodSheet._set(row, PRODUCTION_COL.OUTPUT_ITEM_NAME, outputItemName);
  prodSheet._set(row, PRODUCTION_COL.COMPONENTS_CONSUMED, '[]');
  prodSheet._set(row, PRODUCTION_COL.COLOR_BREAKDOWN, JSON.stringify(colorBreakdown));
}

console.log('=== Test 1: 3 independent axes (Frame=primary, Mudguard, Rim), 1 each -> combines into ONE 3-way bucket ===');
setLot(2, {
  qty: 25, lotNumber: 'LOT-3AXIS', outputItemName: 'Fitted Frame',
  colorBreakdown: [
    { color: 'Blue-White', qty: 25, countsTowardTotal: true, axisKey: 'pool:frame' },
    { color: 'Black', qty: 25, countsTowardTotal: false, axisKey: 'pool:mudguard' },
    { color: 'BCP', qty: 25, countsTowardTotal: false, axisKey: 'pool:rim' }
  ]
});
recalculateWarehousePool();
let pool = getWarehousePoolData().data.filter(r => r.outputItemName === 'Fitted Frame');
console.log('  buckets:', JSON.stringify(pool));
assert(pool.length === 1, `exactly 1 bucket created (got ${pool.length})`);
assert(pool[0] && pool[0].color === 'Blue-White / Black / BCP', `bucket color is the 3-way composite "Blue-White / Black / BCP" (got "${pool[0] && pool[0].color}")`);
assert(pool[0] && pool[0].producedQty === 25, `bucket producedQty is 25 (got ${pool[0] && pool[0].producedQty})`);

console.log('\n=== Test 2: redundant axis (name-matches primary) still excluded from the combination ===');
prodSheet.rows = []; // reset
setLot(2, {
  qty: 10, lotNumber: 'LOT-REDUNDANT', outputItemName: 'Fitted Frame 2',
  colorBreakdown: [
    { color: 'Red-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:frame' },
    { color: 'Red', qty: 10, countsTowardTotal: false, axisKey: 'pool:mudguard' }, // name-matches "Red-White" -> redundant, excluded
    { color: 'BCP', qty: 10, countsTowardTotal: false, axisKey: 'pool:rim' } // genuinely independent
  ]
});
recalculateWarehousePool();
pool = getWarehousePoolData().data.filter(r => r.outputItemName === 'Fitted Frame 2');
console.log('  buckets:', JSON.stringify(pool));
assert(pool.length === 1, `exactly 1 bucket created (got ${pool.length})`);
assert(pool[0] && pool[0].color === 'Red-White / BCP', `bucket color excludes the redundant "Red" (Mudguard) and only combines Frame+Rim: "Red-White / BCP" (got "${pool[0] && pool[0].color}")`);

console.log('\n=== Test 3: one axis contributes 2 entries (ambiguous) -> falls back to per-entry crediting, no guessing ===');
prodSheet.rows = [];
setLot(2, {
  qty: 10, lotNumber: 'LOT-AMBIGUOUS', outputItemName: 'Fitted Frame 3',
  colorBreakdown: [
    { color: 'Blue-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:frame' },
    { color: 'BCP', qty: 6, countsTowardTotal: false, axisKey: 'pool:rim' },
    { color: 'Black', qty: 4, countsTowardTotal: false, axisKey: 'pool:rim' } // SAME axis, 2 entries -> ambiguous which pairs with Frame
  ]
});
recalculateWarehousePool();
pool = getWarehousePoolData().data.filter(r => r.outputItemName === 'Fitted Frame 3');
console.log('  buckets:', JSON.stringify(pool));
assert(pool.length === 3, `falls back to 3 SEPARATE single-color buckets, not combined (got ${pool.length})`);
const colors3 = pool.map(p => p.color).sort();
assert(JSON.stringify(colors3) === JSON.stringify(['BCP', 'Black', 'Blue-White']), `each entry credited under its own color: ${JSON.stringify(colors3)}`);

console.log('\n=== Test 4: pre-existing 2-axis case (primary + 1 independent) still works exactly as before ===');
prodSheet.rows = [];
setLot(2, {
  qty: 8, lotNumber: 'LOT-2AXIS', outputItemName: 'Fitted Frame 4',
  colorBreakdown: [
    { color: 'Pink-White', qty: 8, countsTowardTotal: true, axisKey: 'pool:frame' },
    { color: 'Black', qty: 8, countsTowardTotal: false, axisKey: 'pool:rim' }
  ]
});
recalculateWarehousePool();
pool = getWarehousePoolData().data.filter(r => r.outputItemName === 'Fitted Frame 4');
console.log('  buckets:', JSON.stringify(pool));
assert(pool.length === 1, `exactly 1 bucket (unchanged 2-axis behavior) (got ${pool.length})`);
assert(pool[0] && pool[0].color === 'Pink-White / Black', `bucket color "Pink-White / Black" (unchanged) (got "${pool[0] && pool[0].color}")`);

console.log('\n=== Test 5a: a SINGLE independent entry with no axisKey still combines (matches original 1-independent-entry behavior) ===');
prodSheet.rows = [];
setLot(2, {
  qty: 5, lotNumber: 'LOT-NOAXISKEY-1', outputItemName: 'Fitted Frame 5',
  colorBreakdown: [
    { color: 'Orange-White', qty: 5, countsTowardTotal: true }, // no axisKey
    { color: 'Silver', qty: 5, countsTowardTotal: false } // no axisKey, but only ONE such entry
  ]
});
recalculateWarehousePool();
pool = getWarehousePoolData().data.filter(r => r.outputItemName === 'Fitted Frame 5');
console.log('  buckets:', JSON.stringify(pool));
assert(pool.length === 1, `exactly 1 bucket (single blank-axisKey entry still combines) (got ${pool.length})`);
assert(pool[0] && pool[0].color === 'Orange-White / Silver', `bucket color "Orange-White / Silver" (got "${pool[0] && pool[0].color}")`);

console.log('\n=== Test 5b: TWO independent entries with no axisKey at all -- no way to tell which pairs with which, must fall back (not combine) ===');
prodSheet.rows = [];
setLot(2, {
  qty: 5, lotNumber: 'LOT-NOAXISKEY-2', outputItemName: 'Fitted Frame 6',
  colorBreakdown: [
    { color: 'Orange-White', qty: 5, countsTowardTotal: true }, // no axisKey
    { color: 'Silver', qty: 5, countsTowardTotal: false }, // no axisKey
    { color: 'Gold', qty: 5, countsTowardTotal: false } // no axisKey -- 2 blank-axisKey entries now, ambiguous
  ]
});
recalculateWarehousePool();
pool = getWarehousePoolData().data.filter(r => r.outputItemName === 'Fitted Frame 6');
console.log('  buckets:', JSON.stringify(pool));
assert(pool.length === 3, `falls back to 3 separate buckets -- no axis info to safely combine 2 blank-axisKey entries (got ${pool.length})`);

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
