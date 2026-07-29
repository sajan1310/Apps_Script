/**
 * A composite color's axes appear in the order THIS PROCESS'S OWN RECIPE
 * declares them -- not alphabetically, and not in Warehouse Pool row order.
 *
 * A POOL recipe row is the association with the upstream process that
 * produces that item, so recipe row order is exactly "this process's
 * inputs, in the sequence the operator arranged them". The same order
 * drives the Production checklist, so the combination shown while logging
 * is the one the lot is credited under.
 *
 * Proven by building the SAME two processes twice with the recipe rows in
 * opposite orders: the composite must follow the recipe each time.
 *
 * Run: node .pw-test/test_composite_follows_recipe_order.js
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
['config.js', 'utils.js', 'module_units.js', 'module_items.js', 'module_process.js', 'module_production.js', 'module_warehouse.js', 'module_stock.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});
vm.runInContext('global.APP_CONFIG=APP_CONFIG; global.PRODUCTION_COL=PRODUCTION_COL;', ctx, { filename: 'expose.js' });
const C = ctx;
const { APP_CONFIG, PRODUCTION_COL } = C;

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
  prodSheet._set(row, PRODUCTION_COL.COMPONENTS_CONSUMED, JSON.stringify(o.componentsConsumed || []));
  prodSheet._set(row, PRODUCTION_COL.COLOR_BREAKDOWN, JSON.stringify(o.colorBreakdown));
}

// Two upstream feeders, each with 2+ pool colors so each is a real axis.
function seedFeeders() {
  addLot({ processId: 'PRC-RIM', qty: 20, lotNumber: 'R1', outputItemName: 'Fitted Rim',
    colorBreakdown: [{ color: 'Black', qty: 10, countsTowardTotal: true, axisKey: 'a' },
                     { color: 'BCP', qty: 10, countsTowardTotal: true, axisKey: 'a' }] });
  addLot({ processId: 'PRC-MUD', qty: 20, lotNumber: 'M1', outputItemName: 'Painted Mudguard',
    colorBreakdown: [{ color: 'Grey', qty: 10, countsTowardTotal: true, axisKey: 'b' },
                     { color: 'Silver', qty: 10, countsTowardTotal: true, axisKey: 'b' }] });
  addLot({ processId: 'PRC-FRM', qty: 20, lotNumber: 'F1', outputItemName: 'Painted Frame',
    colorBreakdown: [{ color: 'Blue-White', qty: 10, countsTowardTotal: true, axisKey: 'c' },
                     { color: 'Red-White', qty: 10, countsTowardTotal: true, axisKey: 'c' }] });
}

// Builds the assembly process with its POOL rows in the given order and
// returns the composite color its lot gets credited under.
function compositeForRecipeOrder(orderedPoolItems, label) {
  ss.sheets[APP_CONFIG.SHEETS.PROCESS_MASTER] = undefined;
  ss.sheets[APP_CONFIG.SHEETS.PROCESS_COMPONENTS] = undefined;
  delete ss.sheets[APP_CONFIG.SHEETS.PROCESS_MASTER];
  delete ss.sheets[APP_CONFIG.SHEETS.PROCESS_COMPONENTS];
  prodSheet.rows = []; nextRow = 2;
  seedFeeders();
  C.recalculateWarehousePool();

  const res = C.saveProcess({
    processName: 'Assembly ' + label, sequence: 4, lotPrefix: 'ZA' + label,
    outputItemName: 'Assembled ' + label, isFinalStage: false, active: true, remarks: '',
    components: JSON.stringify(orderedPoolItems.map(n => ({
      itemName: n, sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON'
    })))
  });
  if (!res.success) throw new Error('saveProcess failed: ' + res.message);
  const pid = res.data.processId;

  const axes = C.computeColorAxesForProcess(pid,
    C.getProcessComponentsData(pid).data || [],
    C.getWarehousePoolData().data || [],
    []);
  const axisKeys = axes.map(a => a.key);

  // Log a lot: Frame is primary, the other two axes contribute one color each.
  const keyOf = (item) => axes.find(a => a.label.toLowerCase() === item.toLowerCase()).key;
  addLot({ processId: pid, qty: 6, lotNumber: 'A' + label, outputItemName: 'Assembled ' + label,
    colorBreakdown: [
      { color: 'Blue-White', qty: 6, countsTowardTotal: true, axisKey: keyOf('Painted Frame') },
      { color: 'Black', qty: 6, countsTowardTotal: false, axisKey: keyOf('Fitted Rim') },
      { color: 'Grey', qty: 6, countsTowardTotal: false, axisKey: keyOf('Painted Mudguard') }
    ] });
  C.recalculateWarehousePool();
  const bucket = C.getWarehousePoolData().data.find(r => r.outputItemName === 'Assembled ' + label);
  return { composite: bucket && bucket.color, axisKeys: axisKeys, axisLabels: axes.map(a => a.label) };
}

console.log('=== Test 1: recipe lists Rim before Mudguard ===');
const A = compositeForRecipeOrder(['Painted Frame', 'Fitted Rim', 'Painted Mudguard'], 'A');
console.log('   axes in checklist order:', JSON.stringify(A.axisLabels));
console.log('   composite:', A.composite);
assert(JSON.stringify(A.axisLabels) === JSON.stringify(['Painted Frame', 'Fitted Rim', 'Painted Mudguard']),
  'the checklist renders axes in recipe row order');
assert(A.composite === 'Blue-White / Black / Grey',
  `composite follows the recipe: primary first, then Rim, then Mudguard (got "${A.composite}")`);

console.log('\n=== Test 2: same three inputs, recipe lists Mudguard before Rim ===');
const B = compositeForRecipeOrder(['Painted Frame', 'Painted Mudguard', 'Fitted Rim'], 'B');
console.log('   axes in checklist order:', JSON.stringify(B.axisLabels));
console.log('   composite:', B.composite);
assert(JSON.stringify(B.axisLabels) === JSON.stringify(['Painted Frame', 'Painted Mudguard', 'Fitted Rim']),
  'the checklist follows the reordered recipe');
assert(B.composite === 'Blue-White / Grey / Black',
  `composite follows the recipe the other way round (got "${B.composite}")`);

console.log('\n=== Test 3: the two orders really are different, i.e. recipe order is what decides ===');
assert(A.composite !== B.composite,
  'reordering the recipe reorders the composite — order is authored, not alphabetical');
assert(A.composite.split(' / ').sort().join() === B.composite.split(' / ').sort().join(),
  'both describe the same three colors, only the sequence differs');

console.log('\n=== Test 4: order is stable across repeated recalculations ===');
const first = C.getWarehousePoolData().data.find(r => r.outputItemName === 'Assembled B').color;
C.recalculateWarehousePool();
C.recalculateWarehousePool();
const again = C.getWarehousePoolData().data.find(r => r.outputItemName === 'Assembled B').color;
assert(first === again, `repeated rebuilds keep the same composite (got "${first}" then "${again}")`);

console.log('\n=== Test 5: getAxisOrderByProcess exposes the same order in bulk ===');
const orderMap = C.getAxisOrderByProcess();
const anyProcess = Object.keys(orderMap).find(k => Object.keys(orderMap[k]).length === 3);
assert(!!anyProcess, 'the 3-axis assembly process appears in the bulk order map');
const positions = orderMap[anyProcess];
assert(Object.values(positions).sort().join() === '0,1,2',
  `positions are a dense 0..n-1 sequence (got ${JSON.stringify(positions)})`);

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
