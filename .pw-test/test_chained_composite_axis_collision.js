/**
 * Chained multi-stage colour identity: Process A produces composite pool
 * colors from 2 axes; Process B consumes A's output AND adds its own axis.
 *
 * Covers the segment-collision defect: recalculateWarehousePool's Pass 1
 * decides an axis is "redundant" (a mirror of the primary, e.g. a Mudguard
 * row auto-synced to the Frame color) via _colorNamesMatch, which treats
 * "/" as a segment boundary. A composite primary color IS a run of
 * segments, so a downstream axis whose color equalled an INHERITED segment
 * was silently swallowed:
 *
 *   primary "Blue-White / Black" (frame + upstream rim)
 *   + Seat Color axis = "Black"   -> dropped as "redundant"
 *   -> credited "Blue-White / Black", losing the seat entirely, merging
 *      black-seat and brown-seat output into one bucket.
 *
 * Fix: redundancy is judged against the primary's OWN color (first segment)
 * only. Segments after the first came from an upstream process's
 * independent axes, so no downstream axis can be a mirror of one. Plain
 * (non-composite) primaries are unaffected, so every pre-existing
 * single-stage case behaves byte-for-byte as before.
 *
 * Run: node .pw-test/test_chained_composite_axis_collision.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = 'c:\\Users\\erkar\\my-app-script-project';

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) { this.sheet = sheet; this.row = row; this.col = col; this.numRows = numRows; this.numCols = numCols; }
  getValues() { const out = []; for (let r = 0; r < this.numRows; r++) { const a = []; for (let c = 0; c < this.numCols; c++) a.push(this.sheet._get(this.row + r, this.col + c)); out.push(a); } return out; }
  getValue() { return this.sheet._get(this.row, this.col); }
  setValues(v) { v.forEach((ra, r) => ra.forEach((val, c) => this.sheet._set(this.row + r, this.col + c, val))); return this; }
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
  getLastColumn() { let m = 0; this.rows.forEach(row => { for (let c = row.length; c >= 1; c--) { if (row[c - 1] !== '' && row[c - 1] !== undefined && row[c - 1] !== null) { m = Math.max(m, c); break; } } }); return m; }
  getRange(row, col, numRows = 1, numCols = 1) { return new FakeRange(this, row, col, numRows, numCols); }
  appendRow(arr) { const r = this.getLastRow() + 1; arr.forEach((v, i) => this._set(r, i + 1, v)); }
  deleteRow(r) { this.rows.splice(r - 1, 1); } deleteRows(r, n) { this.rows.splice(r - 1, n); }
  insertRows(r, n) { for (let i = 0; i < n; i++) this.rows.splice(r - 1, 0, []); }
  insertColumnsAfter(a, h) { this.rows.forEach(row => { row.splice(a, 0, ...new Array(h).fill('')); }); }
}
class FakeSpreadsheet { constructor() { this.sheets = {}; } getSheetByName(n) { return this.sheets[n] || null; } addSheet(n) { const s = new FakeSheet(n); this.sheets[n] = s; return s; } insertSheet(n) { return this.addSheet(n); } }
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
const { APP_CONFIG, PRODUCTION_COL, recalculateWarehousePool, getWarehousePoolData, saveProcess, getProcessColorAxes, getProcessColorGroups } = ctx;

let failures = 0;
function assert(cond, msg) { if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); } }

const prodSheet = ss.addSheet(APP_CONFIG.SHEETS.PRODUCTION);
let nextRow = 2;
function addLot({ processId, qty, lotNumber, outputItemName, colorBreakdown, componentsConsumed }) {
  const row = nextRow++;
  prodSheet._set(row, PRODUCTION_COL.DATE, '01/01/2026');
  prodSheet._set(row, PRODUCTION_COL.QTY, qty);
  prodSheet._set(row, PRODUCTION_COL.STATUS, 'Completed');
  prodSheet._set(row, PRODUCTION_COL.PROCESS_ID, processId);
  prodSheet._set(row, PRODUCTION_COL.LOT_NUMBER, lotNumber);
  prodSheet._set(row, PRODUCTION_COL.OUTPUT_ITEM_NAME, outputItemName);
  prodSheet._set(row, PRODUCTION_COL.COMPONENTS_CONSUMED, JSON.stringify(componentsConsumed || []));
  prodSheet._set(row, PRODUCTION_COL.COLOR_BREAKDOWN, JSON.stringify(colorBreakdown));
}
function bucketsFor(item) {
  return getWarehousePoolData().data
    .filter(r => r.outputItemName === item)
    .map(r => ({ color: r.color, produced: r.producedQty, consumed: r.consumedQty }))
    .sort((a, b) => a.color.localeCompare(b.color));
}
// Stage A: 3 frame colors x one Black rim -> 3 composite pool buckets.
function seedStageA() {
  addLot({
    processId: 'PRC-A', qty: 30, lotNumber: 'A-1', outputItemName: 'Fitted Frame',
    colorBreakdown: [
      { color: 'Blue-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:painted frame' },
      { color: 'Red-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:painted frame' },
      { color: 'Green-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:painted frame' },
      { color: 'Black', qty: 30, countsTowardTotal: false, axisKey: 'pool:fitted rim' }
    ]
  });
}

console.log('=== Test 1: Stage A credits composite buckets ===');
seedStageA();
recalculateWarehousePool();
let pool = bucketsFor('Fitted Frame');
console.log('  ', JSON.stringify(pool));
assert(JSON.stringify(pool.map(b => b.color)) === JSON.stringify(['Blue-White / Black', 'Green-White / Black', 'Red-White / Black']),
  'A produces one composite bucket per frame color');

console.log('\n=== Test 2: Process B sees A\'s composites as one axis, its own colors as another ===');
const bRes = saveProcess({
  processName: 'Seat Fitting', sequence: 2, lotPrefix: 'ZSF',
  outputItemName: 'Seated Cycle', isFinalStage: false, active: true, remarks: '',
  components: JSON.stringify([
    { itemName: 'Fitted Frame', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' },
    { itemName: 'Black Seat', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Black', colorAxis: 'Seat Color' },
    { itemName: 'Brown Seat', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Brown', colorAxis: 'Seat Color' }
  ])
});
assert(bRes.success, 'Process B saved: ' + (bRes.message || ''));
const bId = bRes.data && bRes.data.processId;
const axes = (getProcessColorAxes(bId).data || {}).axes || [];
console.log('  axes:', JSON.stringify(axes.map(a => ({ key: a.key, colors: a.colors }))));
assert(axes.length === 2, `B renders 2 independent axes (got ${axes.length})`);
const poolAxis = axes.find(a => a.source === 'pool');
const tagAxis = axes.find(a => a.source === 'tag');
assert(poolAxis && JSON.stringify(poolAxis.colors) === JSON.stringify(['Blue-White / Black', 'Green-White / Black', 'Red-White / Black']),
  "B's pool axis offers A's composite colors verbatim");
assert(tagAxis && JSON.stringify(tagAxis.colors) === JSON.stringify(['Black', 'Brown']), "B's own Seat Color axis offers Black/Brown");
const groups = getProcessColorGroups(bId).data || [];
assert(groups.indexOf('Blue-White / Black') !== -1 && groups.indexOf('Brown') !== -1,
  'saveProduction validation accepts both an inherited composite and B\'s own color');

console.log('\n=== Test 3: B adds a NON-colliding color -> identity deepens by one segment ===');
addLot({
  processId: bId, qty: 10, lotNumber: 'B-1', outputItemName: 'Seated Cycle',
  colorBreakdown: [
    { color: 'Blue-White / Black', qty: 5, countsTowardTotal: true, axisKey: 'pool:fitted frame' },
    { color: 'Red-White / Black', qty: 5, countsTowardTotal: true, axisKey: 'pool:fitted frame' },
    { color: 'Brown', qty: 10, countsTowardTotal: false, axisKey: 'tag:seat color' }
  ],
  componentsConsumed: [
    { itemName: 'Fitted Frame', sourceType: 'POOL', qty: 5, colorGroup: 'Blue-White / Black' },
    { itemName: 'Fitted Frame', sourceType: 'POOL', qty: 5, colorGroup: 'Red-White / Black' }
  ]
});
recalculateWarehousePool();
pool = bucketsFor('Seated Cycle');
console.log('  ', JSON.stringify(pool));
assert(JSON.stringify(pool.map(b => b.color)) === JSON.stringify(['Blue-White / Black / Brown', 'Red-White / Black / Brown']),
  "B's own axis is appended to each inherited composite");
assert(pool.every(b => b.produced === 5), 'each composite carries its own primary qty (5)');
const frameAfter = bucketsFor('Fitted Frame');
console.log('  upstream:', JSON.stringify(frameAfter));
assert(frameAfter.find(b => b.color === 'Blue-White / Black').consumed === 5
  && frameAfter.find(b => b.color === 'Green-White / Black').consumed === 0,
  'consumption debits exactly the composite buckets B actually consumed');

console.log('\n=== Test 4: SEGMENT COLLISION -- B\'s own color equals an INHERITED segment ===');
prodSheet.rows = []; nextRow = 2;
seedStageA();
addLot({
  processId: bId, qty: 10, lotNumber: 'B-2', outputItemName: 'Seated Cycle',
  colorBreakdown: [
    { color: 'Blue-White / Black', qty: 5, countsTowardTotal: true, axisKey: 'pool:fitted frame' },
    { color: 'Red-White / Black', qty: 5, countsTowardTotal: true, axisKey: 'pool:fitted frame' },
    // "Black" is a real, independent Seat Color -- it only LOOKS redundant
    // because the frame composite inherited a "Black" rim segment upstream.
    { color: 'Black', qty: 10, countsTowardTotal: false, axisKey: 'tag:seat color' }
  ]
});
recalculateWarehousePool();
pool = bucketsFor('Seated Cycle');
console.log('  ', JSON.stringify(pool));
assert(JSON.stringify(pool.map(b => b.color)) === JSON.stringify(['Blue-White / Black / Black', 'Red-White / Black / Black']),
  'the seat color survives instead of being swallowed by the inherited "Black" segment');

console.log('\n=== Test 5: single-color lot, same collision (the pre-existing shape) ===');
prodSheet.rows = []; nextRow = 2;
seedStageA();
addLot({
  processId: bId, qty: 5, lotNumber: 'B-3', outputItemName: 'Seated Cycle Solo',
  colorBreakdown: [
    { color: 'Blue-White / Black', qty: 5, countsTowardTotal: true, axisKey: 'pool:fitted frame' },
    { color: 'Black', qty: 5, countsTowardTotal: false, axisKey: 'tag:seat color' }
  ]
});
recalculateWarehousePool();
pool = bucketsFor('Seated Cycle Solo');
console.log('  ', JSON.stringify(pool));
assert(JSON.stringify(pool.map(b => b.color)) === JSON.stringify(['Blue-White / Black / Black']),
  'a one-color lot keeps its seat color too (collision is not multi-color-specific)');

console.log('\n=== Test 6: a genuine mirror of the primary\'s OWN color is still dropped ===');
prodSheet.rows = [];
nextRow = 2;
addLot({
  processId: 'PRC-A', qty: 20, lotNumber: 'A-MIRROR', outputItemName: 'Fitted Frame 2',
  colorBreakdown: [
    { color: 'Red-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:frame' },
    { color: 'Blue-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:frame' },
    { color: 'Red', qty: 10, countsTowardTotal: false, axisKey: 'pool:mudguard' },  // mirrors Red-White
    { color: 'Blue', qty: 10, countsTowardTotal: false, axisKey: 'pool:mudguard' }, // mirrors Blue-White
    { color: 'BCP', qty: 20, countsTowardTotal: false, axisKey: 'pool:rim' }
  ]
});
recalculateWarehousePool();
pool = bucketsFor('Fitted Frame 2');
console.log('  ', JSON.stringify(pool));
assert(JSON.stringify(pool.map(b => b.color)) === JSON.stringify(['Blue-White / BCP', 'Red-White / BCP']),
  'mudguard mirror still excluded, rim still combined (unchanged behavior)');

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
