/**
 * Round-trip invariants across Production -> Warehouse Pool -> Dispatch.
 *
 * Rather than asserting one known bug, this checks properties that must hold
 * for ANY sequence of operations, which is where arithmetic drift hides:
 *
 *   I1. Pool credit for a process's output == the sum of its Completed lots'
 *       quantities. No more (double-count), no less (dropped bucket).
 *   I2. Pool debit of an upstream POOL component == qtyPerUnit x downstream
 *       lot qty.
 *   I3. Deleting a lot returns the pool exactly to its pre-lot state
 *       (delete/re-add is a no-op on every bucket).
 *   I4. Un-completing a lot withdraws its credit; re-completing restores it.
 *   I5. Ready to Dispatch == pool credit for a final-stage output, and a
 *       dispatch reduces it by exactly the dispatched qty.
 *   I6. recalculateWarehousePool is idempotent — running it twice in a row
 *       changes nothing.
 *
 * Run: node .pw-test/verify_production_pool_dispatch_invariants.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

class FakeRange {
  constructor(s, r, c, nr, nc) { this.sheet = s; this.row = r; this.col = c; this.numRows = nr; this.numCols = nc; }
  getValues() { const o = []; for (let r = 0; r < this.numRows; r++) { const a = []; for (let c = 0; c < this.numCols; c++) a.push(this.sheet._get(this.row + r, this.col + c)); o.push(a); } return o; }
  getValue() { return this.sheet._get(this.row, this.col); }
  setValues(v) {
    if (!Array.isArray(v) || v.length !== this.numRows) throw new Error(`The number of rows in the data does not match the number of rows in the range. The data has ${Array.isArray(v) ? v.length : 0} but the range has ${this.numRows}.`);
    v.forEach(ra => { if (!Array.isArray(ra) || ra.length !== this.numCols) throw new Error(`The number of columns in the data does not match the number of columns in the range. The data has ${Array.isArray(ra) ? ra.length : 0} but the range has ${this.numCols}.`); });
    v.forEach((ra, r) => ra.forEach((val, c) => this.sheet._set(this.row + r, this.col + c, val))); return this;
  }
  setValue(v) { if (this.numRows !== 1 || this.numCols !== 1) { for (let r = 0; r < this.numRows; r++) for (let c = 0; c < this.numCols; c++) this.sheet._set(this.row + r, this.col + c, v); return this; } this.sheet._set(this.row, this.col, v); return this; }
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
['config.js', 'utils.js', 'module_units.js', 'module_tags.js', 'module_items.js', 'module_po.js',
 'module_bom.js', 'module_process.js', 'module_production.js', 'module_warehouse.js',
 'module_stock.js', 'module_clients.js', 'module_contractors.js', 'module_dispatch.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});
vm.runInContext('global.APP_CONFIG=APP_CONFIG; global.PRODUCTION_COL=PRODUCTION_COL; global.WAREHOUSE_POOL_COL=WAREHOUSE_POOL_COL;', ctx, { filename: 'expose.js' });
const C = ctx;
const { APP_CONFIG, PRODUCTION_COL } = C;

let failures = 0;
function assert(cond, msg) { if (!cond) { failures++; console.error('  FAIL:', msg); } else { console.log('  PASS:', msg); } }
const near = (a, b) => Math.abs(a - b) < 0.0001;

// Pool snapshot as a comparable, order-independent object.
function poolSnapshot() {
  const out = {};
  (C.getWarehousePoolData().data || []).forEach(r => {
    const k = `${r.outputItemName}|${r.productTag}|${r.color}`;
    out[k] = { produced: r.producedQty, consumed: r.consumedQty, available: r.availableQty };
  });
  return out;
}
const availableFor = (item) => (C.getWarehousePoolData().data || [])
  .filter(r => r.outputItemName === item)
  .reduce((s, r) => s + r.availableQty, 0);
const producedFor = (item) => (C.getWarehousePoolData().data || [])
  .filter(r => r.outputItemName === item)
  .reduce((s, r) => s + r.producedQty, 0);

// ── Setup: a 2-stage chain, no colors (isolates the arithmetic from the
// composite-color machinery, which its own tests already cover).
C.initProductionSheet();
let res = C.saveProcess({
  processName: 'INV Painting', sequence: 1, lotPrefix: 'IVP', outputItemName: 'INV Painted Frame',
  isFinalStage: false, active: true, components: JSON.stringify([])
});
if (!res.success) { console.error('setup stage1 failed:', res.message); process.exit(1); }
const P1 = res.data.processId;

res = C.saveProcess({
  processName: 'INV Packing', sequence: 2, lotPrefix: 'IVK', outputItemName: 'INV Packed Cycle',
  isFinalStage: true, active: true,
  components: JSON.stringify([{ itemName: 'INV Painted Frame', sourceType: 'POOL', qtyPerUnit: 2, colorGroup: 'COMMON' }])
});
if (!res.success) { console.error('setup stage2 failed:', res.message); process.exit(1); }
const P2 = res.data.processId;

// saveProduction requires at least one consumed component, so every lot
// carries a token ITEM-sourced one. ITEM components hit Stock, not the pool,
// so they don't perturb any of the pool invariants below.
const lot = (processId, qty, status) => C.saveProduction({
  processId, assignedTo: 'INV Contractor', qty, date: '01/01/2026', status,
  componentsConsumed: JSON.stringify([
    { itemName: 'INV Raw Bar', sourceType: 'ITEM', qty: qty, colorGroup: 'COMMON' }])
});
const lotConsuming = (processId, qty, status, perUnit) => C.saveProduction({
  processId, assignedTo: 'INV Contractor', qty, date: '01/01/2026', status,
  componentsConsumed: JSON.stringify([
    { itemName: 'INV Painted Frame', sourceType: 'POOL', qty: perUnit * qty, colorGroup: 'COMMON' }])
});

console.log('\n=== I1: pool credit == sum of Completed lot quantities ===');
res = lot(P1, 100, 'Completed'); assert(res.success, 'stage-1 lot of 100 saved: ' + res.message);
res = lot(P1, 40, 'Completed'); assert(res.success, 'stage-1 lot of 40 saved: ' + res.message);
res = lot(P1, 25, 'Pending');   assert(res.success, 'stage-1 lot of 25 left Pending: ' + res.message);
assert(near(producedFor('INV Painted Frame'), 140),
  `credited 140 (100+40), the Pending 25 excluded (got ${producedFor('INV Painted Frame')})`);

console.log('\n=== I6: recalculateWarehousePool is idempotent ===');
const before = JSON.stringify(poolSnapshot());
C.recalculateWarehousePool();
C.recalculateWarehousePool();
assert(before === JSON.stringify(poolSnapshot()), 'two extra recalcs changed nothing');

console.log('\n=== I2: pool debit == qtyPerUnit x downstream lot qty ===');
const beforeDownstream = availableFor('INV Painted Frame');
res = lotConsuming(P2, 30, 'Completed', 2);
assert(res.success, 'stage-2 lot of 30 consuming 2/unit saved: ' + res.message);
assert(near(availableFor('INV Painted Frame'), beforeDownstream - 60),
  `debited exactly 60 (30 x 2) (was ${beforeDownstream}, now ${availableFor('INV Painted Frame')})`);
assert(near(producedFor('INV Packed Cycle'), 30),
  `stage-2 output credited 30 (got ${producedFor('INV Packed Cycle')})`);

console.log('\n=== I5: Ready to Dispatch tracks the final-stage pool ===');
let ready = (C.getReadyToDispatchData().data || []).filter(r => r.productId === 'INV Packed Cycle');
assert(ready.length === 1 && near(ready[0].readyQty, 30),
  `30 ready (got ${JSON.stringify(ready.map(r => r.readyQty))})`);
res = C.saveDispatch({
  clientName: 'INV Client', dispatchDate: '02/01/2026',
  lines: JSON.stringify([{ productId: 'INV Packed Cycle', productName: 'INV Packed Cycle', qty: 12, rate: 0 }])
});
assert(res.success, 'dispatching 12 of 30 succeeds: ' + res.message);
const dispatchNumber = res.data.dispatchNumber;
ready = (C.getReadyToDispatchData().data || []).filter(r => r.productId === 'INV Packed Cycle');
assert(ready.length === 1 && near(ready[0].readyQty, 18),
  `18 ready after dispatching 12 (got ${JSON.stringify(ready.map(r => r.readyQty))})`);
res = C.saveDispatch({
  clientName: 'INV Client', dispatchDate: '02/01/2026',
  lines: JSON.stringify([{ productId: 'INV Packed Cycle', productName: 'INV Packed Cycle', qty: 19, rate: 0 }])
});
assert(!res.success, `over-dispatching 19 of 18 is rejected (got: ${res.message})`);

console.log('\n=== I3: delete a lot -> pool returns exactly to its pre-lot state ===');
const snapBeforeExtraLot = JSON.stringify(poolSnapshot());
res = lot(P1, 77, 'Completed');
assert(res.success, 'extra stage-1 lot of 77 saved: ' + res.message);
assert(!near(producedFor('INV Painted Frame'), 140), 'the extra lot did change the pool (sanity)');
// Locate its sheet row, then delete it with the same guard the client sends.
const prodRows = C.getProductionData().data || [];
const extra = prodRows.find(r => r.qty === 77);
assert(!!extra, 'the extra lot is findable in getProductionData');
res = C.deleteProduction(extra.rowIdx, extra.productId, extra.qty);
assert(res.success, 'deleting it succeeds: ' + res.message);
assert(JSON.stringify(poolSnapshot()) === snapBeforeExtraLot,
  'every pool bucket is byte-for-byte back to its pre-lot state');

console.log('\n=== I4: un-completing withdraws the credit, re-completing restores it ===');
const snapAllCompleted = JSON.stringify(poolSnapshot());
const target = (C.getProductionData().data || []).find(r => r.qty === 40);
assert(!!target, 'the 40-unit lot is findable');
res = C.updateProductionStatus(target.rowIdx, target.qty, 'Pending');
assert(res.success, 'setting it Pending succeeds: ' + res.message);
assert(near(producedFor('INV Painted Frame'), 100),
  `credit drops to 100 with the 40 withdrawn (got ${producedFor('INV Painted Frame')})`);
res = C.updateProductionStatus(target.rowIdx, target.qty, 'Completed');
assert(res.success, 'setting it back to Completed succeeds: ' + res.message);
assert(JSON.stringify(poolSnapshot()) === snapAllCompleted,
  'the pool is byte-for-byte back to its all-Completed state');

console.log('\n=== I3b: deleting the dispatch releases exactly what it claimed ===');
res = C.deleteDispatch(dispatchNumber, undefined, undefined);
assert(res.success, 'deleting the 12-unit dispatch succeeds: ' + res.message);
ready = (C.getReadyToDispatchData().data || []).filter(r => r.productId === 'INV Packed Cycle');
assert(ready.length === 1 && near(ready[0].readyQty, 30),
  `back to 30 ready (got ${JSON.stringify(ready.map(r => r.readyQty))})`);

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
