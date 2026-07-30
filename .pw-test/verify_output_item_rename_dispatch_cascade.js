/**
 * Renaming a final-stage process's Output Item Name must cascade to the
 * Dispatch ledger.
 *
 * An UNTAGGED final-stage lot is dispatched under its Output Item Name — that
 * is literally what saveDispatch writes into DISPATCH_COL.PRODUCT_ID (see
 * _computeReadyToDispatchMap: `productId: isTagged ? r.productTag :
 * r.outputItemName`), and recalculateWarehousePool's Pass 3 debits the pool by
 * matching that productId back against `bucket.outputItemName`.
 *
 * _renamePoolOutputItemNameEverywhere rewrites the name on the Production
 * sheet, Process Components and Warehouse Pool Opening — but not on Dispatch.
 * So after a rename, Pass 3 finds no bucket for the old name, the dispatch
 * debit silently disappears, and every already-shipped unit comes back as
 * Ready to Dispatch.
 *
 * Run: node .pw-test/verify_output_item_rename_dispatch_cascade.js
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
vm.runInContext('global.APP_CONFIG=APP_CONFIG; global.DISPATCH_COL=DISPATCH_COL;', ctx, { filename: 'expose.js' });
const C = ctx;
const { APP_CONFIG, DISPATCH_COL } = C;

let failures = 0;
function assert(cond, msg) { if (!cond) { failures++; console.error('  FAIL:', msg); } else { console.log('  PASS:', msg); } }
const near = (a, b) => Math.abs(a - b) < 0.0001;

C.initProductionSheet();
let res = C.saveProcess({
  processName: 'RN Packing', sequence: 1, lotPrefix: 'RNP', outputItemName: 'RN Packed Cycle',
  isFinalStage: true, active: true, components: JSON.stringify([])
});
if (!res.success) { console.error('setup failed:', res.message); process.exit(1); }
const PID = res.data.processId;

res = C.saveProduction({
  processId: PID, assignedTo: 'RN C', qty: 100, date: '01/01/2026', status: 'Completed',
  componentsConsumed: JSON.stringify([{ itemName: 'RN Raw', sourceType: 'ITEM', qty: 100, colorGroup: 'COMMON' }])
});
assert(res.success, '100-unit untagged final-stage lot created: ' + res.message);

const readyFor = name => {
  const rows = (C.getReadyToDispatchData().data || []).filter(r => r.productId === name);
  return rows.reduce((s, r) => s + r.readyQty, 0);
};
assert(near(readyFor('RN Packed Cycle'), 100), `100 ready (got ${readyFor('RN Packed Cycle')})`);

res = C.saveDispatch({
  clientName: 'RN Client', dispatchDate: '02/01/2026',
  lines: JSON.stringify([{ productId: 'RN Packed Cycle', productName: 'RN Packed Cycle', qty: 60, rate: 0 }])
});
assert(res.success, 'dispatching 60 succeeds: ' + res.message);
assert(near(readyFor('RN Packed Cycle'), 40), `40 ready after the dispatch (got ${readyFor('RN Packed Cycle')})`);

console.log('\n=== Rename the process\'s Output Item Name ===');
res = C.saveProcess({
  processId: PID, processName: 'RN Packing', sequence: 1, lotPrefix: 'RNP',
  outputItemName: 'RN Packed Bicycle',   // renamed
  isFinalStage: true, active: true, components: JSON.stringify([])
});
assert(res.success, 'rename saved: ' + res.message);
C.recalculateWarehousePool();

console.log('\n=== Test 1: the dispatch still counts against the renamed product ===');
const readyOld = readyFor('RN Packed Cycle');
const readyNew = readyFor('RN Packed Bicycle');
console.log(`   ready under old name: ${readyOld}, under new name: ${readyNew}`);
assert(near(readyOld, 0), `nothing is left stranded under the old name (got ${readyOld})`);
assert(near(readyNew, 40),
  `still 40 ready under the new name — the 60 dispatched must NOT come back (got ${readyNew})`);

console.log('\n=== Test 2: the Dispatch ledger itself was rekeyed ===');
const dispatchRows = (C.getDispatchData().data || []);
console.log('   dispatch productIds:', JSON.stringify(dispatchRows.map(d => d.productId)));
assert(dispatchRows.every(d => d.productId !== 'RN Packed Cycle'),
  'no Dispatch row still references the old Output Item Name');
assert(dispatchRows.some(d => d.productId === 'RN Packed Bicycle'),
  'the Dispatch row now references the new Output Item Name');

console.log('\n=== Test 3: total pool consumption still reflects the 60 shipped ===');
const consumed = (C.getWarehousePoolData().data || [])
  .filter(r => r.outputItemName === 'RN Packed Bicycle')
  .reduce((s, r) => s + r.consumedQty, 0);
assert(near(consumed, 60), `the pool still shows 60 consumed by Dispatch (got ${consumed})`);

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
