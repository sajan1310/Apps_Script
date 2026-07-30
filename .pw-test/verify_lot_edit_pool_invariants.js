/**
 * Invariants for EDITING an already-Completed lot (as opposed to creating or
 * deleting one, which verify_production_pool_dispatch_invariants.js covers).
 * Editing is the path with the most moving parts — the pool has already been
 * credited/debited under the OLD values and has to end up exactly as if the
 * new values had been the original ones.
 *
 *   E1. Editing a lot's qty moves the pool by exactly the delta.
 *   E2. Editing a lot is equivalent to deleting it and creating the new one:
 *       the resulting pool must be byte-for-byte identical either way.
 *   E3. Editing a colour lot's breakdown moves credit BETWEEN colour buckets
 *       with no total leakage, and empties the bucket it moved away from.
 *   E4. Editing a downstream lot's qty re-debits the upstream pool by the new
 *       amount only (no residue from the old debit).
 *
 * Run: node .pw-test/verify_lot_edit_pool_invariants.js
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

function freshCtx() {
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
  vm.runInContext('global.APP_CONFIG=APP_CONFIG; global.PRODUCTION_COL=PRODUCTION_COL;', ctx, { filename: 'expose.js' });
  ctx.__ss = ss;
  return ctx;
}

let failures = 0;
function assert(cond, msg) { if (!cond) { failures++; console.error('  FAIL:', msg); } else { console.log('  PASS:', msg); } }
const near = (a, b) => Math.abs(a - b) < 0.0001;

function poolSnapshot(C) {
  const out = {};
  (C.getWarehousePoolData().data || []).forEach(r => {
    out[`${r.outputItemName}|${r.productTag}|${r.color}`] =
      [r.producedQty, r.consumedQty, r.availableQty];
  });
  return out;
}
const producedFor = (C, item) => (C.getWarehousePoolData().data || [])
  .filter(r => r.outputItemName === item).reduce((s, r) => s + r.producedQty, 0);
const availableFor = (C, item) => (C.getWarehousePoolData().data || [])
  .filter(r => r.outputItemName === item).reduce((s, r) => s + r.availableQty, 0);
const bucketsFor = (C, item) => (C.getWarehousePoolData().data || [])
  .filter(r => r.outputItemName === item)
  .map(r => `${r.color}=${r.producedQty}`).sort().join(', ');

// Builds the 2-stage chain in a given context. Returns { P1, P2 }.
function setupChain(C) {
  C.initProductionSheet();
  let res = C.saveProcess({
    processName: 'EDT Painting', sequence: 1, lotPrefix: 'EDP', outputItemName: 'EDT Painted Frame',
    isFinalStage: false, active: true, components: JSON.stringify([])
  });
  if (!res.success) throw new Error('stage1: ' + res.message);
  const P1 = res.data.processId;
  res = C.saveProcess({
    processName: 'EDT Packing', sequence: 2, lotPrefix: 'EDK', outputItemName: 'EDT Packed Cycle',
    isFinalStage: true, active: true,
    components: JSON.stringify([{ itemName: 'EDT Painted Frame', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' }])
  });
  if (!res.success) throw new Error('stage2: ' + res.message);
  return { P1, P2: res.data.processId };
}
const tokenComp = qty => JSON.stringify([{ itemName: 'EDT Raw Bar', sourceType: 'ITEM', qty, colorGroup: 'COMMON' }]);
const findLot = (C, qty) => (C.getProductionData().data || []).find(r => r.qty === qty);

console.log('\n=== E1: editing a lot\'s qty moves the pool by exactly the delta ===');
{
  const C = freshCtx();
  const { P1 } = setupChain(C);
  let res = C.saveProduction({ processId: P1, assignedTo: 'EDT C', qty: 50, date: '01/01/2026',
    status: 'Completed', componentsConsumed: tokenComp(50) });
  assert(res.success, 'lot of 50 created: ' + res.message);
  assert(near(producedFor(C, 'EDT Painted Frame'), 50), `pool at 50 (got ${producedFor(C, 'EDT Painted Frame')})`);

  const lot = findLot(C, 50);
  res = C.saveProduction({ rowIdx: lot.rowIdx, processId: P1, assignedTo: 'EDT C', qty: 80,
    date: '01/01/2026', status: 'Completed', componentsConsumed: tokenComp(80) });
  assert(res.success, 'edited 50 -> 80: ' + res.message);
  assert(near(producedFor(C, 'EDT Painted Frame'), 80),
    `pool moved to exactly 80, not 130 (got ${producedFor(C, 'EDT Painted Frame')})`);

  const lot2 = findLot(C, 80);
  res = C.saveProduction({ rowIdx: lot2.rowIdx, processId: P1, assignedTo: 'EDT C', qty: 20,
    date: '01/01/2026', status: 'Completed', componentsConsumed: tokenComp(20) });
  assert(res.success, 'edited 80 -> 20 (downward): ' + res.message);
  assert(near(producedFor(C, 'EDT Painted Frame'), 20),
    `pool moved down to exactly 20 (got ${producedFor(C, 'EDT Painted Frame')})`);
}

console.log('\n=== E2: edit == delete + recreate (identical resulting pool) ===');
{
  // Path A: create 50, edit to 80.
  const A = freshCtx();
  let a = setupChain(A);
  A.saveProduction({ processId: a.P1, assignedTo: 'EDT C', qty: 50, date: '01/01/2026',
    status: 'Completed', componentsConsumed: tokenComp(50) });
  const lotA = findLot(A, 50);
  A.saveProduction({ rowIdx: lotA.rowIdx, processId: a.P1, assignedTo: 'EDT C', qty: 80,
    date: '01/01/2026', status: 'Completed', componentsConsumed: tokenComp(80) });

  // Path B: create 80 outright.
  const B = freshCtx();
  let b = setupChain(B);
  B.saveProduction({ processId: b.P1, assignedTo: 'EDT C', qty: 80, date: '01/01/2026',
    status: 'Completed', componentsConsumed: tokenComp(80) });

  assert(JSON.stringify(poolSnapshot(A)) === JSON.stringify(poolSnapshot(B)),
    'editing to 80 leaves the same pool as creating 80 directly\n' +
    `      edited:  ${JSON.stringify(poolSnapshot(A))}\n      created: ${JSON.stringify(poolSnapshot(B))}`);
}

console.log('\n=== E3: editing a colour breakdown moves credit between buckets, no leakage ===');
{
  const C = freshCtx();
  C.initProductionSheet();
  // A colour-tracked process: two colour sub-groups on its recipe.
  let res = C.saveProcess({
    processName: 'EDT Colour Painting', sequence: 1, lotPrefix: 'EDC', outputItemName: 'EDT Colour Frame',
    isFinalStage: false, active: true,
    components: JSON.stringify([
      { itemName: 'EDT Red Paint', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Red', colorAxis: 'Frame Colour' },
      { itemName: 'EDT Blue Paint', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Blue', colorAxis: 'Frame Colour' }])
  });
  assert(res.success, 'colour process created: ' + res.message);
  const PC = res.data.processId;

  res = C.saveProduction({
    processId: PC, assignedTo: 'EDT C', date: '01/01/2026', status: 'Completed',
    colorBreakdown: JSON.stringify([
      { color: 'Red', qty: 30, countsTowardTotal: true },
      { color: 'Blue', qty: 20, countsTowardTotal: true }]),
    componentsConsumed: JSON.stringify([
      { itemName: 'EDT Red Paint', sourceType: 'ITEM', qty: 30, colorGroup: 'Red' },
      { itemName: 'EDT Blue Paint', sourceType: 'ITEM', qty: 20, colorGroup: 'Blue' }])
  });
  assert(res.success, 'Red 30 + Blue 20 lot created: ' + res.message);
  assert(near(producedFor(C, 'EDT Colour Frame'), 50),
    `total credit 50 (got ${producedFor(C, 'EDT Colour Frame')}) — buckets: ${bucketsFor(C, 'EDT Colour Frame')}`);

  // Move it all to Blue: the Red bucket must end at 0, not keep its 30.
  const lot = findLot(C, 50);
  res = C.saveProduction({
    rowIdx: lot.rowIdx, processId: PC, assignedTo: 'EDT C', date: '01/01/2026', status: 'Completed',
    colorBreakdown: JSON.stringify([{ color: 'Blue', qty: 50, countsTowardTotal: true }]),
    componentsConsumed: JSON.stringify([
      { itemName: 'EDT Blue Paint', sourceType: 'ITEM', qty: 50, colorGroup: 'Blue' }])
  });
  assert(res.success, 'edited to Blue 50 only: ' + res.message);
  assert(near(producedFor(C, 'EDT Colour Frame'), 50),
    `total still 50, no leakage (got ${producedFor(C, 'EDT Colour Frame')}) — buckets: ${bucketsFor(C, 'EDT Colour Frame')}`);
  const red = (C.getWarehousePoolData().data || [])
    .filter(r => r.outputItemName === 'EDT Colour Frame' && r.color.toLowerCase() === 'red');
  assert(red.every(r => near(r.producedQty, 0)),
    `the Red bucket is emptied, not left holding its old 30 (buckets: ${bucketsFor(C, 'EDT Colour Frame')})`);
}

console.log('\n=== E4: editing a downstream lot\'s qty re-debits upstream by the new amount only ===');
{
  const C = freshCtx();
  const { P1, P2 } = setupChain(C);
  C.saveProduction({ processId: P1, assignedTo: 'EDT C', qty: 100, date: '01/01/2026',
    status: 'Completed', componentsConsumed: tokenComp(100) });
  assert(near(availableFor(C, 'EDT Painted Frame'), 100), 'upstream at 100 to start');

  let res = C.saveProduction({ processId: P2, assignedTo: 'EDT C', qty: 30, date: '01/01/2026',
    status: 'Completed',
    componentsConsumed: JSON.stringify([{ itemName: 'EDT Painted Frame', sourceType: 'POOL', qty: 30, colorGroup: 'COMMON' }]) });
  assert(res.success, 'downstream lot of 30 created: ' + res.message);
  assert(near(availableFor(C, 'EDT Painted Frame'), 70),
    `upstream down to 70 (got ${availableFor(C, 'EDT Painted Frame')})`);

  const lot = findLot(C, 30);
  res = C.saveProduction({ rowIdx: lot.rowIdx, processId: P2, assignedTo: 'EDT C', qty: 10,
    date: '01/01/2026', status: 'Completed',
    componentsConsumed: JSON.stringify([{ itemName: 'EDT Painted Frame', sourceType: 'POOL', qty: 10, colorGroup: 'COMMON' }]) });
  assert(res.success, 'downstream edited 30 -> 10: ' + res.message);
  assert(near(availableFor(C, 'EDT Painted Frame'), 90),
    `upstream back up to 90, no residue of the old 30-unit debit (got ${availableFor(C, 'EDT Painted Frame')})`);
  assert(near(producedFor(C, 'EDT Packed Cycle'), 10),
    `downstream output now credits 10, not 40 (got ${producedFor(C, 'EDT Packed Cycle')})`);
}

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
