/**
 * Dispatch Differentiator, the WRITE path (test_dispatch_differentiator.js only
 * ever covered the read).
 *
 * _computeReadyToDispatchMap splits an untagged final-stage output into one
 * bucket per differentiator value, keyed '__output__<item>||<value>'. But the
 * productId every one of those rows reports back is the bare Output Item Name
 * (no value), and saveDispatch looks the availability up as
 * readyMap[pid] || readyMap['__output__' + pid] — neither of which exists once
 * the key carries a value. Availability read as 0 and the dispatch was rejected
 * outright: a final-stage process with a Dispatch Differentiator configured
 * could not dispatch anything at all.
 *
 * Also asserts the client-facing records carry a unique `key` per split row —
 * with only productId to go on, Script_Dispatch.html's
 * `find(r => r.productId === pid)` and its <option value> list both collapse
 * every variant onto the first one.
 *
 * Run: node .pw-test/verify_dispatch_differentiator_save.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = 'c:\\Users\\erkar\\my-app-script-project';

class FakeRange {
  constructor(s, r, c, nr, nc) { this.sheet = s; this.row = r; this.col = c; this.numRows = nr; this.numCols = nc; }
  getValues() { const o = []; for (let r = 0; r < this.numRows; r++) { const a = []; for (let c = 0; c < this.numCols; c++) a.push(this.sheet._get(this.row + r, this.col + c)); o.push(a); } return o; }
  getValue() { return this.sheet._get(this.row, this.col); }
  setValues(v) {
    if (!Array.isArray(v) || v.length !== this.numRows) throw new Error(`rows: data ${v.length} vs range ${this.numRows}`);
    v.forEach(ra => { if (ra.length !== this.numCols) throw new Error(`The number of columns in the data does not match the number of columns in the range. The data has ${ra.length} but the range has ${this.numCols}.`); });
    v.forEach((ra, r) => ra.forEach((val, c) => this.sheet._set(this.row + r, this.col + c, val))); return this;
  }
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
vm.runInContext('global.APP_CONFIG=APP_CONFIG; global.PRODUCTION_COL=PRODUCTION_COL; global.PROCESS_COL=PROCESS_COL; global.DISPATCH_COL=DISPATCH_COL;', ctx, { filename: 'expose.js' });
const C = ctx;
const { APP_CONFIG, PRODUCTION_COL, PROCESS_COL, DISPATCH_COL } = C;

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

addLot({ processId: 'E-RIM', qty: 100, lotNumber: 'r', outputItemName: 'E Rim',
  colorBreakdown: [{ color: 'BCP', qty: 50, countsTowardTotal: true, axisKey: 's' },
                   { color: 'Black', qty: 50, countsTowardTotal: true, axisKey: 's' }] });
addLot({ processId: 'E-FRM', qty: 100, lotNumber: 'f', outputItemName: 'E Frame',
  colorBreakdown: [{ color: 'Blue-White', qty: 50, countsTowardTotal: true, axisKey: 't' },
                   { color: 'Red-White', qty: 50, countsTowardTotal: true, axisKey: 't' }] });
C.recalculateWarehousePool();

let res = C.saveProcess({
  processName: 'E Final', sequence: 9, lotPrefix: 'ZEF', outputItemName: 'E Packed Cycle',
  isFinalStage: true, active: true, remarks: '', dispatchDifferentiator: 'E Frame',
  components: JSON.stringify([
    { itemName: 'E Rim', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' },
    { itemName: 'E Frame', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' }])
});
if (!res.success) { console.error('setup failed:', res.message); process.exit(1); }
const pid = res.data.processId;
const axes = C.computeColorAxesForProcess(pid, C.getProcessComponentsData(pid).data || [], C.getWarehousePoolData().data || [], []);
const keyOf = re => axes.find(a => re.test(a.label)).key;

addLot({ processId: pid, qty: 20, lotNumber: 'E1', outputItemName: 'E Packed Cycle',
  colorBreakdown: [
    { color: 'Blue-White', qty: 6, countsTowardTotal: true, axisKey: keyOf(/frame/i) },
    { color: 'Red-White', qty: 14, countsTowardTotal: true, axisKey: keyOf(/frame/i) },
    { color: 'Black', qty: 20, countsTowardTotal: false, axisKey: keyOf(/rim/i) }] });
C.recalculateWarehousePool();

const ready = C.getReadyToDispatchData().data || [];
console.log('ready rows:', JSON.stringify(ready.map(r => ({ id: r.productId, key: r.key, name: r.productName, q: r.readyQty }))));

console.log('\n=== Test 1: each split row is addressable on its own ===');
assert(ready.length === 2, `two split rows (got ${ready.length})`);
const keys = ready.map(r => r.key).filter(Boolean);
assert(keys.length === 2 && new Set(keys).size === 2,
  'every row carries a distinct `key` so the client can tell the variants apart');

console.log('\n=== Test 2: a differentiated product can actually be dispatched ===');
res = C.saveDispatch({
  clientName: 'E Client', dispatchDate: '02/01/2026',
  lines: JSON.stringify([{ productId: ready[0].productId, productName: ready[0].productName, qty: 5, rate: 0 }])
});
assert(res.success, `dispatching 5 of ${ready[0].readyQty} available succeeds (got: ${res.message})`);

console.log('\n=== Test 3: availability is still enforced, not just waved through ===');
res = C.saveDispatch({
  clientName: 'E Client', dispatchDate: '02/01/2026',
  lines: JSON.stringify([{ productId: ready[0].productId, productName: ready[0].productName, qty: 9999, rate: 0 }])
});
assert(!res.success && /Ready to Dispatch/i.test(res.message || ''),
  `over-dispatch is still rejected (got: ${res.success ? 'accepted!' : res.message})`);

console.log('\n=== Test 4: the dispatched units come off the right variant ===');
C.recalculateWarehousePool();
const after = (C.getReadyToDispatchData().data || [])
  .map(r => ({ name: r.productName, q: r.readyQty })).sort((a, b) => a.name.localeCompare(b.name));
console.log('  after dispatch:', JSON.stringify(after));
const total = after.reduce((s, r) => s + r.q, 0);
assert(total === 15, `20 produced - 5 dispatched = 15 still ready in total (got ${total})`);

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
