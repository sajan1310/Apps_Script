/**
 * Canonical composite-color ordering + verifyProductionColorChain.
 *
 * Segment order used to come straight from the Color Breakdown array, i.e.
 * from checklist DOM order, which follows Warehouse Pool sheet row order --
 * and that sheet is itself rebuilt on every recalculation. Two lots of the
 * SAME product could therefore be credited as "Blue-White / Black / Grey"
 * and "Blue-White / Grey / Black", splitting one product's stock across two
 * buckets. Needs 3+ axes (2+ independent ones) to bite, which is exactly
 * the shape of the reported 3-axis process.
 *
 * _composeLotColorKey now anchors the primary segment (it encodes the
 * process chain) and orders the independent segments by AXIS identity, so
 * the key depends only on which axes contributed, never on input order.
 *
 * Run: node .pw-test/test_color_chain_order_and_verify.js
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
const logLines = [];
const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush: () => {} },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
  console, Logger: { log: (m) => logLines.push(String(m)) },
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
  prodSheet._set(row, PRODUCTION_COL.PROCESS_ID, o.processId || 'PRC-A');
  prodSheet._set(row, PRODUCTION_COL.LOT_NUMBER, o.lotNumber);
  prodSheet._set(row, PRODUCTION_COL.OUTPUT_ITEM_NAME, o.outputItemName);
  prodSheet._set(row, PRODUCTION_COL.COMPONENTS_CONSUMED, JSON.stringify(o.componentsConsumed || []));
  prodSheet._set(row, PRODUCTION_COL.COLOR_BREAKDOWN, JSON.stringify(o.colorBreakdown));
}
const colorsFor = (item) => C.getWarehousePoolData().data
  .filter(r => r.outputItemName === item)
  .map(r => ({ color: r.color, produced: r.producedQty })).sort((a, b) => a.color.localeCompare(b.color));

const FRAME = { color: 'Blue-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:frame' };
const RIM = { color: 'Black', qty: 10, countsTowardTotal: false, axisKey: 'pool:rim' };
const MUD = { color: 'Grey', qty: 10, countsTowardTotal: false, axisKey: 'pool:mudguard' };

console.log('=== Test 1: segment helpers ===');
assert(JSON.stringify(C._colorSegments('Blue-White / Black / Grey')) === JSON.stringify(['Blue-White', 'Black', 'Grey']),
  '_colorSegments splits a composite into its axis segments');
assert(JSON.stringify(C._colorSegments('Blue')) === JSON.stringify(['Blue']), 'a plain color is a single segment');
assert(C._colorOrderKey('Blue-White / Black / Grey') === C._colorOrderKey('Blue-White / Grey / Black'),
  '_colorOrderKey is order-independent');
assert(C._colorOrderKey('Blue-White / Black') !== C._colorOrderKey('Blue-White / Grey'),
  '_colorOrderKey still distinguishes genuinely different combinations');

console.log('\n=== Test 2: the SAME combination keys ONE bucket regardless of entry order ===');
prodSheet.rows = []; nextRow = 2;
addLot({ qty: 10, lotNumber: 'L1', outputItemName: 'Cycle', colorBreakdown: [FRAME, RIM, MUD] });
addLot({ qty: 10, lotNumber: 'L2', outputItemName: 'Cycle', colorBreakdown: [FRAME, MUD, RIM] });
addLot({ qty: 10, lotNumber: 'L3', outputItemName: 'Cycle', colorBreakdown: [MUD, FRAME, RIM] });
C.recalculateWarehousePool();
let pool = colorsFor('Cycle');
console.log('  ', JSON.stringify(pool));
assert(pool.length === 1, `three lots, three entry orders, ONE bucket (got ${pool.length})`);
assert(pool[0].produced === 30, `all 30 units land in it (got ${pool[0] && pool[0].produced})`);
// Independents are ordered by AXIS key, not by color: "pool:mudguard"
// (Grey) sorts before "pool:rim" (Black), so Grey leads regardless of the
// order the entries arrived in. What matters is that it is always the same.
assert(pool[0].color === 'Blue-White / Grey / Black',
  `canonical order is primary first, then independents by axis key (got "${pool[0] && pool[0].color}")`);

console.log('\n=== Test 3: the primary segment stays anchored (it encodes the chain) ===');
prodSheet.rows = []; nextRow = 2;
// A primary color that sorts LAST alphabetically must still lead.
addLot({ qty: 5, lotNumber: 'L4', outputItemName: 'Cycle2',
  colorBreakdown: [{ color: 'Zinc', qty: 5, countsTowardTotal: true, axisKey: 'pool:frame' },
                   { color: 'Amber', qty: 5, countsTowardTotal: false, axisKey: 'pool:rim' }] });
C.recalculateWarehousePool();
pool = colorsFor('Cycle2');
console.log('  ', JSON.stringify(pool));
assert(pool[0] && pool[0].color === 'Zinc / Amber',
  `primary "Zinc" leads even though "Amber" sorts first (got "${pool[0] && pool[0].color}")`);

console.log('\n=== Test 4: an inherited composite primary is not re-ordered ===');
prodSheet.rows = []; nextRow = 2;
addLot({ qty: 5, lotNumber: 'L5', outputItemName: 'Cycle3',
  colorBreakdown: [{ color: 'Blue-White / Black', qty: 5, countsTowardTotal: true, axisKey: 'pool:fitted' },
                   { color: 'Brown', qty: 5, countsTowardTotal: false, axisKey: 'tag:seat' }] });
C.recalculateWarehousePool();
pool = colorsFor('Cycle3');
console.log('  ', JSON.stringify(pool));
assert(pool[0] && pool[0].color === 'Blue-White / Black / Brown',
  `upstream segments keep their chain order, this stage appends (got "${pool[0] && pool[0].color}")`);

console.log('\n=== Test 5: a consumption recorded in the OLD order still debits the right bucket ===');
prodSheet.rows = []; nextRow = 2;
addLot({ qty: 10, lotNumber: 'L6', outputItemName: 'Cycle', colorBreakdown: [FRAME, RIM, MUD] });
addLot({ qty: 4, lotNumber: 'L7', outputItemName: 'Packed', processId: 'PRC-B',
  colorBreakdown: [{ color: 'Blue-White / Black / Grey', qty: 4, countsTowardTotal: true, axisKey: 'pool:cycle' }],
  // Stored before canonical ordering existed -> segments in the old order.
  componentsConsumed: [{ itemName: 'Cycle', sourceType: 'POOL', qty: 4, colorGroup: 'Blue-White / Grey / Black' }] });
C.recalculateWarehousePool();
const cycle = C.getWarehousePoolData().data.filter(r => r.outputItemName === 'Cycle');
console.log('  ', JSON.stringify(cycle.map(r => ({ color: r.color, prod: r.producedQty, cons: r.consumedQty }))));
assert(cycle.length === 1, `no phantom bucket opened for the old-order name (got ${cycle.length})`);
assert(cycle[0].consumedQty === 4, `the real bucket is debited 4 (got ${cycle[0] && cycle[0].consumedQty})`);

console.log('\n=== Test 6: verifyProductionColorChain reports a genuine unresolved debit ===');
prodSheet.rows = []; nextRow = 2;
addLot({ qty: 10, lotNumber: 'L8', outputItemName: 'Cycle', colorBreakdown: [FRAME, RIM, MUD] });
addLot({ qty: 2, lotNumber: 'L9', outputItemName: 'Packed', processId: 'PRC-B',
  colorBreakdown: [{ color: 'Blue-White / Black / Grey', qty: 2, countsTowardTotal: true, axisKey: 'pool:cycle' }],
  componentsConsumed: [{ itemName: 'Cycle', sourceType: 'POOL', qty: 2, colorGroup: 'Magenta / Teal' }] });
C.recalculateWarehousePool();
let res = C.verifyProductionColorChain();
console.log('  counts:', JSON.stringify(res.data.countsByType));
assert(res.success, 'verify runs and returns a response');
assert((res.data.countsByType['unresolved-debit'] || 0) === 1,
  `the unmatchable "Magenta / Teal" consumption is reported once (got ${res.data.countsByType['unresolved-debit'] || 0})`);
const ud = res.data.findings.find(f => f.type === 'unresolved-debit');
assert(ud && ud.detail.outputItemName === 'Cycle' && ud.detail.color === 'Magenta / Teal'
  && ud.detail.lotNumbers.indexOf('L9') !== -1,
  'the finding names the item, the phantom color and the consuming lot');

console.log('\n=== Test 7: a clean chain produces no findings ===');
prodSheet.rows = []; nextRow = 2;
addLot({ qty: 10, lotNumber: 'L10', outputItemName: 'Cycle', colorBreakdown: [FRAME, RIM, MUD] });
addLot({ qty: 4, lotNumber: 'L11', outputItemName: 'Packed', processId: 'PRC-B',
  colorBreakdown: [{ color: 'Blue-White / Black / Grey', qty: 4, countsTowardTotal: true, axisKey: 'pool:cycle' }],
  componentsConsumed: [{ itemName: 'Cycle', sourceType: 'POOL', qty: 4, colorGroup: 'Blue-White / Black / Grey' }] });
C.recalculateWarehousePool();
res = C.verifyProductionColorChain();
console.log('  counts:', JSON.stringify(res.data.countsByType), '| lots:', res.data.lotsChecked, '| buckets:', res.data.bucketsChecked);
assert(res.data.findings.length === 0, `no findings on a consistent chain (got ${JSON.stringify(res.data.findings.map(f => f.type))})`);
assert(res.data.lotsChecked === 2, `both color lots were examined (got ${res.data.lotsChecked})`);
assert(/No color-chain problems found/.test(res.message), `message reports a clean bill (got "${res.message}")`);

console.log('\n=== Test 8: verify is read-only ===');
const before = JSON.stringify(prodSheet.rows);
const poolBefore = JSON.stringify(C.getWarehousePoolData().data);
C.verifyProductionColorChain();
assert(JSON.stringify(prodSheet.rows) === before, 'Production sheet is untouched');
assert(JSON.stringify(C.getWarehousePoolData().data) === poolBefore, 'Warehouse Pool is untouched');

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
