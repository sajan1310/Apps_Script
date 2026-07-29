/**
 * Regression test for the reported bug: "there are 3 axis in 2nd sequence
 * process which fails to reflect in next sequence."
 *
 * A real Edit Production Lot (Row #116) had THREE axes and MULTIPLE primary
 * colors checked at once:
 *   Painted Frame (PRIMARY): Blue-White 10, Pink-White 10, Purple-White 10, Red-White 10
 *   Broad Mudguard         : Blue 10, Pink 10, Purple 10, Red 10   (redundant -- name-matches the frame colors)
 *   Fitted Rim             : Black 40                              (genuinely independent, one color for the whole lot)
 *
 * recalculateWarehousePool's Pass 1 combining only ever fired when EXACTLY
 * ONE primary entry existed, so this lot fell all the way back to per-entry
 * crediting: 9 fragmented single-color buckets (Blue-White, Blue, Pink-White,
 * Pink, ..., Black) instead of 4 real composite ones. The downstream
 * ("next sequence") process consuming this pool item then saw a checklist of
 * loose half-colors -- Blue AND Blue-White AND Black as if they were separate
 * producible outputs -- with quantities that never lined up (the reported
 * "-2 avail." on Blue-White).
 *
 * Run: node .pw-test/test_multi_primary_composite_credit.js
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
function bucketsFor(item) {
  return getWarehousePoolData().data
    .filter(r => r.outputItemName === item)
    .map(r => ({ color: r.color, qty: r.producedQty }))
    .sort((a, b) => a.color.localeCompare(b.color));
}

console.log('=== Test 1: Row #116 -- 4 primary colors x redundant Mudguard axis x one independent Rim color ===');
setLot(2, {
  qty: 40, lotNumber: 'LOT-116', outputItemName: 'Fitted Frame 16 inch Crysta S/Rim',
  colorBreakdown: [
    { color: 'Blue-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:frame' },
    { color: 'Pink-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:frame' },
    { color: 'Purple-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:frame' },
    { color: 'Red-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:frame' },
    { color: 'Blue', qty: 10, countsTowardTotal: false, axisKey: 'pool:mudguard' },
    { color: 'Pink', qty: 10, countsTowardTotal: false, axisKey: 'pool:mudguard' },
    { color: 'Purple', qty: 10, countsTowardTotal: false, axisKey: 'pool:mudguard' },
    { color: 'Red', qty: 10, countsTowardTotal: false, axisKey: 'pool:mudguard' },
    { color: 'Black', qty: 40, countsTowardTotal: false, axisKey: 'pool:rim' }
  ]
});
recalculateWarehousePool();
let pool = bucketsFor('Fitted Frame 16 inch Crysta S/Rim');
console.log('  buckets:', JSON.stringify(pool));
assert(pool.length === 4, `one composite bucket per primary color, not 9 fragments (got ${pool.length})`);
assert(JSON.stringify(pool) === JSON.stringify([
  { color: 'Blue-White / Black', qty: 10 },
  { color: 'Pink-White / Black', qty: 10 },
  { color: 'Purple-White / Black', qty: 10 },
  { color: 'Red-White / Black', qty: 10 }
]), 'each primary color pairs with the single independent Rim color, redundant Mudguard excluded');
assert(pool.reduce((s, b) => s + b.qty, 0) === 40, `credited total still equals the lot's 40 units (got ${pool.reduce((s, b) => s + b.qty, 0)})`);

console.log('\n=== Test 2: multiple primaries, NO other axis -> unchanged per-color buckets (no phantom composites) ===');
prodSheet.rows = [];
setLot(2, {
  qty: 20, lotNumber: 'LOT-PLAIN', outputItemName: 'Painted Frame',
  colorBreakdown: [
    { color: 'Blue-White', qty: 12, countsTowardTotal: true, axisKey: 'pool:frame' },
    { color: 'Red-White', qty: 8, countsTowardTotal: true, axisKey: 'pool:frame' }
  ]
});
recalculateWarehousePool();
pool = bucketsFor('Painted Frame');
console.log('  buckets:', JSON.stringify(pool));
assert(JSON.stringify(pool) === JSON.stringify([
  { color: 'Blue-White', qty: 12 },
  { color: 'Red-White', qty: 8 }
]), 'a single-axis multi-color lot still credits one plain bucket per color');

console.log('\n=== Test 3: multiple primaries + an independent axis with 2+ colors -> ambiguous, falls back per-entry ===');
prodSheet.rows = [];
setLot(2, {
  qty: 20, lotNumber: 'LOT-AMBIG', outputItemName: 'Fitted Frame X',
  colorBreakdown: [
    { color: 'Blue-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:frame' },
    { color: 'Red-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:frame' },
    { color: 'Black', qty: 12, countsTowardTotal: false, axisKey: 'pool:rim' },
    { color: 'BCP', qty: 8, countsTowardTotal: false, axisKey: 'pool:rim' }
  ]
});
recalculateWarehousePool();
pool = bucketsFor('Fitted Frame X');
console.log('  buckets:', JSON.stringify(pool));
assert(pool.length === 4, `2 Rim colors give no way to tell which frame they pair with -> 4 separate buckets (got ${pool.length})`);

console.log('\n=== Test 4: partially-checked redundant axis (Mudguard only mirrors SOME primaries) still combines ===');
prodSheet.rows = [];
setLot(2, {
  qty: 20, lotNumber: 'LOT-PARTIAL', outputItemName: 'Fitted Frame Y',
  colorBreakdown: [
    { color: 'Blue-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:frame' },
    { color: 'Red-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:frame' },
    { color: 'Blue', qty: 10, countsTowardTotal: false, axisKey: 'pool:mudguard' }, // mirrors Blue-White only
    { color: 'Black', qty: 20, countsTowardTotal: false, axisKey: 'pool:rim' }
  ]
});
recalculateWarehousePool();
pool = bucketsFor('Fitted Frame Y');
console.log('  buckets:', JSON.stringify(pool));
assert(JSON.stringify(pool) === JSON.stringify([
  { color: 'Blue-White / Black', qty: 10 },
  { color: 'Red-White / Black', qty: 10 }
]), 'a redundant entry mirroring any primary is dropped, the independent Rim color still pairs with both');

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
