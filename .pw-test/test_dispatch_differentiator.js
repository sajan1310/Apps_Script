/**
 * Dispatch Differentiator (PROCESS_COL.DISPATCH_DIFFERENTIATOR).
 *
 * A final-stage output's pool colour is the whole chain's composite
 * ("Black / Blue-White / Brown"). Dispatch usually cares about ONE axis of
 * it, so the final-stage process names that axis and Ready to Dispatch is
 * split into one row per distinct value, labelled "<Output Item Name> / <value>".
 *
 * Run: node .pw-test/test_dispatch_differentiator.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = 'c:\\Users\\erkar\\my-app-script-project';

class FakeRange {
  constructor(s, r, c, nr, nc) { this.sheet = s; this.row = r; this.col = c; this.numRows = nr; this.numCols = nc; }
  getValues() { const o = []; for (let r = 0; r < this.numRows; r++) { const a = []; for (let c = 0; c < this.numCols; c++) a.push(this.sheet._get(this.row + r, this.col + c)); o.push(a); } return o; }
  getValue() { return this.sheet._get(this.row, this.col); }
  setValues(v) { v.forEach((ra, r) => ra.forEach((val, c) => this.sheet._set(this.row + r, this.col + c, val))); return this; }
  setValue(v) { this.sheet._set(this.row, this.col, v); return this; }
  clearContent() { for (let r = 0; r < this.numRows; r++) for (let c = 0; c < this.numCols; c++) this.sheet._set(this.row + r, this.col + c, ''); return this; }
  setFontWeight() { return this; } setBackground() { return this; } setNumberFormat() { return this; }
}
class FakeSheet {
  constructor(n) { this.name = n; this.rows = []; }
  getName() { return this.name; }
  _ensureRow(r) { while (this.rows.length < r) this.rows.push([]); }
  _get(r, c) { this._ensureRow(r); const w = this.rows[r - 1]; return w[c - 1] === undefined ? '' : w[c - 1]; }
  _set(r, c, v) { this._ensureRow(r); const w = this.rows[r - 1]; while (w.length < c) w.push(''); w[c - 1] = v; }
  getLastRow() { for (let r = this.rows.length; r >= 1; r--) { if (this.rows[r - 1].some(v => v !== '' && v !== undefined && v !== null)) return r; } return 0; }
  getLastColumn() { let m = 0; this.rows.forEach(w => { for (let c = w.length; c >= 1; c--) { if (w[c - 1] !== '' && w[c - 1] !== undefined && w[c - 1] !== null) { m = Math.max(m, c); break; } } }); return m; }
  getRange(r, c, nr = 1, nc = 1) { return new FakeRange(this, r, c, nr, nc); }
  appendRow(a) { const r = this.getLastRow() + 1; a.forEach((v, i) => this._set(r, i + 1, v)); }
  deleteRow(r) { this.rows.splice(r - 1, 1); } deleteRows(r, n) { this.rows.splice(r - 1, n); }
  insertRows(r, n) { for (let i = 0; i < n; i++) this.rows.splice(r - 1, 0, []); }
  insertColumnsAfter(a, h) { this.rows.forEach(w => { w.splice(a, 0, ...new Array(h).fill('')); }); }
}
class FakeSpreadsheet { constructor() { this.sheets = {}; } getSheetByName(n) { return this.sheets[n] || null; } addSheet(n) { const s = new FakeSheet(n); this.sheets[n] = s; return s; } insertSheet(n) { return this.addSheet(n); } }
const ss = new FakeSpreadsheet();
const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush: () => {} },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
  console, Logger: { log: () => {} },
  Utilities: { getUuid: () => 'u' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);
['config.js', 'utils.js', 'module_units.js', 'module_items.js', 'module_process.js', 'module_production.js',
 'module_warehouse.js', 'module_stock.js', 'module_clients.js', 'module_dispatch.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});
vm.runInContext('global.APP_CONFIG=APP_CONFIG; global.PRODUCTION_COL=PRODUCTION_COL; global.PROCESS_COL=PROCESS_COL;', ctx, { filename: 'expose.js' });
const C = ctx;
const { APP_CONFIG, PRODUCTION_COL, PROCESS_COL } = C;

let failures = 0;
function assert(cond, msg) { if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); } }

const prodSheet = ss.addSheet(APP_CONFIG.SHEETS.PRODUCTION);
let nextRow = 2;
function addLot(o) {
  const row = nextRow++;
  prodSheet._set(row, PRODUCTION_COL.DATE, '01/01/2026');
  prodSheet._set(row, PRODUCTION_COL.QTY, o.qty);
  prodSheet._set(row, PRODUCTION_COL.STATUS, 'Completed');
  prodSheet._set(row, PRODUCTION_COL.PROCESS_ID, o.processId);
  prodSheet._set(row, PRODUCTION_COL.LOT_NUMBER, o.lotNumber);
  prodSheet._set(row, PRODUCTION_COL.OUTPUT_ITEM_NAME, o.outputItemName);
  prodSheet._set(row, PRODUCTION_COL.COMPONENTS_CONSUMED, '[]');
  prodSheet._set(row, PRODUCTION_COL.COLOR_BREAKDOWN, JSON.stringify(o.colorBreakdown));
}
const readyRows = () => (C.getReadyToDispatchData().data || [])
  .map(r => ({ name: r.productName, diff: r.differentiator, ready: r.readyQty }))
  .sort((a, b) => a.name.localeCompare(b.name));

// Feeders: rim (2 colors) + frame (2 colors), so both are real axes.
addLot({ processId: 'D-RIM', qty: 100, lotNumber: 'r', outputItemName: 'D Rim',
  colorBreakdown: [{ color: 'BCP', qty: 50, countsTowardTotal: true, axisKey: 's' },
                   { color: 'Black', qty: 50, countsTowardTotal: true, axisKey: 's' }] });
addLot({ processId: 'D-FRM', qty: 100, lotNumber: 'f', outputItemName: 'D Frame',
  colorBreakdown: [{ color: 'Blue-White', qty: 50, countsTowardTotal: true, axisKey: 't' },
                   { color: 'Red-White', qty: 50, countsTowardTotal: true, axisKey: 't' }] });
C.recalculateWarehousePool();

console.log('=== Test 1: saveProcess persists the differentiator ===');
let res = C.saveProcess({
  processName: 'D Final', sequence: 9, lotPrefix: 'ZDF', outputItemName: 'D Packed Cycle',
  isFinalStage: true, active: true, remarks: '', dispatchDifferentiator: 'D Frame',
  components: JSON.stringify([
    { itemName: 'D Rim', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' },
    { itemName: 'D Frame', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' }])
});
assert(res.success, 'process saved: ' + (res.message || ''));
const pid = res.data.processId;
let rec = (C.getProcessData(false).data || []).find(p => p.processId === pid);
assert(rec && rec.dispatchDifferentiator === 'D Frame',
  `round-trips through the sheet (got "${rec && rec.dispatchDifferentiator}")`);
assert(rec && rec.isFinalStage === true, 'final-stage flag preserved alongside it');

const axes = C.computeColorAxesForProcess(pid, C.getProcessComponentsData(pid).data || [], C.getWarehousePoolData().data || [], []);
const keyOf = (re) => axes.find(a => re.test(a.label)).key;

console.log('\n=== Test 2: Ready to Dispatch splits one row per differentiator value ===');
addLot({ processId: pid, qty: 20, lotNumber: 'D1', outputItemName: 'D Packed Cycle',
  colorBreakdown: [
    { color: 'Blue-White', qty: 6, countsTowardTotal: true, axisKey: keyOf(/frame/i) },
    { color: 'Red-White', qty: 14, countsTowardTotal: true, axisKey: keyOf(/frame/i) },
    { color: 'Black', qty: 20, countsTowardTotal: false, axisKey: keyOf(/rim/i) }] });
C.recalculateWarehousePool();
let rows = readyRows();
console.log('  ', JSON.stringify(rows));
assert(rows.length === 2, `two rows, one per frame colour (got ${rows.length})`);
assert(JSON.stringify(rows) === JSON.stringify([
  { name: 'D Packed Cycle / Blue-White', diff: 'Blue-White', ready: 6 },
  { name: 'D Packed Cycle / Red-White', diff: 'Red-White', ready: 14 }]),
  'each row is "<Output Item Name> / <value>" and carries its own ready qty');

console.log('\n=== Test 3: the differentiator picks its OWN axis, not a fixed segment position ===');
// The composite is "Black / Blue-White" (rim first, per recipe order), so a
// positional read would have returned the rim colour. It must return the frame's.
const poolColors = C.getWarehousePoolData().data
  .filter(r => r.outputItemName === 'D Packed Cycle').map(r => r.color).sort();
console.log('   pool colours:', JSON.stringify(poolColors));
assert(poolColors.every(c => c.indexOf('Black') === 0),
  'the composite really does lead with the rim colour, so position would give the wrong answer');
assert(rows.every(r => r.diff === 'Blue-White' || r.diff === 'Red-White'),
  'the differentiator resolves by axis membership, never by segment index');

console.log('\n=== Test 4: switching the differentiator to the Rim regroups the same stock ===');
res = C.saveProcess({
  processId: pid, processName: 'D Final', sequence: 9, lotPrefix: 'ZDF', outputItemName: 'D Packed Cycle',
  isFinalStage: true, active: true, remarks: '', dispatchDifferentiator: 'D Rim',
  components: JSON.stringify([
    { itemName: 'D Rim', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' },
    { itemName: 'D Frame', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' }])
});
assert(res.success, 'process updated: ' + (res.message || ''));
rows = readyRows();
console.log('  ', JSON.stringify(rows));
assert(rows.length === 1 && rows[0].name === 'D Packed Cycle / Black' && rows[0].ready === 20,
  `all 20 units collapse into one Black row (got ${JSON.stringify(rows)})`);

console.log('\n=== Test 5: no differentiator -> one aggregate row, exactly as before ===');
res = C.saveProcess({
  processId: pid, processName: 'D Final', sequence: 9, lotPrefix: 'ZDF', outputItemName: 'D Packed Cycle',
  isFinalStage: true, active: true, remarks: '', dispatchDifferentiator: '',
  components: JSON.stringify([
    { itemName: 'D Rim', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' },
    { itemName: 'D Frame', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' }])
});
assert(res.success, 'process updated: ' + (res.message || ''));
rows = readyRows();
console.log('  ', JSON.stringify(rows));
assert(rows.length === 1 && rows[0].name === 'D Packed Cycle' && rows[0].diff === '' && rows[0].ready === 20,
  `one unlabelled aggregate row of 20 (got ${JSON.stringify(rows)})`);

console.log('\n=== Test 6: an unresolvable differentiator degrades to the aggregate row ===');
const procSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PROCESS_MASTER);
const ids = procSheet.getRange(2, PROCESS_COL.PROCESS_ID, procSheet.getLastRow() - 1, 1).getValues();
for (let i = 0; i < ids.length; i++) {
  if (String(ids[i][0]).trim() === pid) {
    procSheet.getRange(i + 2, PROCESS_COL.DISPATCH_DIFFERENTIATOR).setValue('An Axis That No Longer Exists');
  }
}
// _getAllProcessRecords reads the sheet directly, so no cache to clear here.
rows = readyRows();
console.log('  ', JSON.stringify(rows));
assert(rows.length === 1 && rows[0].name === 'D Packed Cycle',
  'a renamed/removed axis falls back to the aggregate row rather than erroring or dropping stock');

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
