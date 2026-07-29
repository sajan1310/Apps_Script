/**
 * Every added process layer extends the colour identity with another
 * "/"-joined segment, and the next layer sees those combinations on its own
 * Colors to Produce checklist.
 *
 *   Fitting  (Rim + Frame)          -> "Black / Blue-White"
 *   Seating  (+ Seat)               -> "Black / Blue-White / Brown"
 *   Packing  (+ Carton)             -> "Black / Blue-White / Brown / Kraft"
 *
 * Also covers the truncation gap this exposed: a POOL item was only treated
 * as a colour axis when it had 2+ pool colours, on the reasoning that a
 * single-coloured input is a fixed input rather than a per-output choice.
 * That is right for a raw input (a rim that is always Black) but wrong for a
 * CHAINED one -- a single colour that is itself a composite carries the
 * accumulated identity of every upstream stage, so excluding it silently
 * dropped the whole "/"-joined history. See _poolItemIsColorAxis.
 *
 * Run: node .pw-test/test_multi_layer_chain_composition.js
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
const bucketsFor = (item) => C.getWarehousePoolData().data
  .filter(r => r.outputItemName === item)
  .map(r => ({ color: r.color, qty: r.producedQty }))
  .sort((a, b) => a.color.localeCompare(b.color));
function mkProcess(name, seq, out, poolItems) {
  const r = C.saveProcess({
    processName: name, sequence: seq, lotPrefix: 'Z' + seq, outputItemName: out,
    isFinalStage: false, active: true, remarks: '',
    components: JSON.stringify(poolItems.map(i => ({ itemName: i, sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' })))
  });
  if (!r.success) throw new Error(name + ': ' + r.message);
  return r.data.processId;
}
const axesOf = (pid) => C.computeColorAxesForProcess(pid,
  C.getProcessComponentsData(pid).data || [], C.getWarehousePoolData().data || [], []);
const axisKey = (pid, re) => axesOf(pid).find(a => re.test(a.label)).key;

// Raw feeders, each with 2 colors so each is a genuine axis.
[['PF-RIM', 'Chain Rim', ['BCP', 'Black']],
 ['PF-FRM', 'Chain Frame', ['Blue-White', 'Red-White']],
 ['PF-SEAT', 'Chain Seat', ['Brown', 'Tan']],
 ['PF-BOX', 'Chain Carton', ['Kraft', 'White Box']]].forEach(([pid, item, colors]) => {
  addLot({ processId: pid, qty: 100, lotNumber: 'seed-' + item, outputItemName: item,
    colorBreakdown: colors.map(c => ({ color: c, qty: 50, countsTowardTotal: true, axisKey: 'seed' })) });
});
C.recalculateWarehousePool();

console.log('=== Layer 1: Rim + Frame ===');
const p1 = mkProcess('Chain Fitting', 2, 'Chain Fitted Frame', ['Chain Rim', 'Chain Frame']);
addLot({ processId: p1, qty: 20, lotNumber: 'C1', outputItemName: 'Chain Fitted Frame',
  colorBreakdown: [
    { color: 'Blue-White', qty: 10, countsTowardTotal: true, axisKey: axisKey(p1, /frame/i) },
    { color: 'Red-White', qty: 10, countsTowardTotal: true, axisKey: axisKey(p1, /frame/i) },
    { color: 'Black', qty: 20, countsTowardTotal: false, axisKey: axisKey(p1, /rim/i) }] });
C.recalculateWarehousePool();
let pool = bucketsFor('Chain Fitted Frame');
console.log('  ', JSON.stringify(pool));
assert(JSON.stringify(pool) === JSON.stringify([
  { color: 'Black / Blue-White', qty: 10 }, { color: 'Black / Red-White', qty: 10 }]),
  'layer 1 joins rim and frame with " / " in recipe order');

console.log('\n=== Layer 2: + Seat ===');
const p2 = mkProcess('Chain Seating', 3, 'Chain Seated Cycle', ['Chain Fitted Frame', 'Chain Seat']);
const p2Colors = C.getProcessColorGroups(p2).data || [];
console.log('   checklist offers:', JSON.stringify(p2Colors));
assert(p2Colors.indexOf('Black / Blue-White') !== -1 && p2Colors.indexOf('Black / Red-White') !== -1,
  "layer 2's checklist offers layer 1's combinations verbatim");
addLot({ processId: p2, qty: 20, lotNumber: 'C2', outputItemName: 'Chain Seated Cycle',
  colorBreakdown: [
    { color: 'Black / Blue-White', qty: 10, countsTowardTotal: true, axisKey: axisKey(p2, /fitted frame/i) },
    { color: 'Black / Red-White', qty: 10, countsTowardTotal: true, axisKey: axisKey(p2, /fitted frame/i) },
    { color: 'Brown', qty: 20, countsTowardTotal: false, axisKey: axisKey(p2, /seat/i) }] });
C.recalculateWarehousePool();
pool = bucketsFor('Chain Seated Cycle');
console.log('  ', JSON.stringify(pool));
assert(JSON.stringify(pool) === JSON.stringify([
  { color: 'Black / Blue-White / Brown', qty: 10 }, { color: 'Black / Red-White / Brown', qty: 10 }]),
  'layer 2 appends the seat colour as a third segment');

console.log('\n=== Layer 3: + Carton ===');
const p3 = mkProcess('Chain Packing', 4, 'Chain Packed Cycle', ['Chain Seated Cycle', 'Chain Carton']);
const p3Colors = C.getProcessColorGroups(p3).data || [];
console.log('   checklist offers:', JSON.stringify(p3Colors));
assert(p3Colors.indexOf('Black / Blue-White / Brown') !== -1,
  "layer 3's checklist offers layer 2's 3-segment combinations");
addLot({ processId: p3, qty: 20, lotNumber: 'C3', outputItemName: 'Chain Packed Cycle',
  colorBreakdown: [
    { color: 'Black / Blue-White / Brown', qty: 10, countsTowardTotal: true, axisKey: axisKey(p3, /seated/i) },
    { color: 'Black / Red-White / Brown', qty: 10, countsTowardTotal: true, axisKey: axisKey(p3, /seated/i) },
    { color: 'Kraft', qty: 20, countsTowardTotal: false, axisKey: axisKey(p3, /carton/i) }] });
C.recalculateWarehousePool();
pool = bucketsFor('Chain Packed Cycle');
console.log('  ', JSON.stringify(pool));
assert(JSON.stringify(pool) === JSON.stringify([
  { color: 'Black / Blue-White / Brown / Kraft', qty: 10 },
  { color: 'Black / Red-White / Brown / Kraft', qty: 10 }]),
  'layer 3 appends a fourth segment; the chain keeps growing');
assert(pool.every(b => C._colorSegments(b.color).length === 4),
  'each bucket carries exactly 4 "/"-joined segments');
assert(pool.reduce((s, b) => s + b.qty, 0) === 20, 'quantities carry through every layer (20 total)');

console.log('\n=== Layer 4: a pass-through layer adds no segment ===');
const p4 = mkProcess('Chain Labelling', 5, 'Chain Labelled Cycle', ['Chain Packed Cycle']);
addLot({ processId: p4, qty: 10, lotNumber: 'C4', outputItemName: 'Chain Labelled Cycle',
  colorBreakdown: [
    { color: 'Black / Blue-White / Brown / Kraft', qty: 10, countsTowardTotal: true, axisKey: axisKey(p4, /packed/i) }] });
C.recalculateWarehousePool();
pool = bucketsFor('Chain Labelled Cycle');
console.log('  ', JSON.stringify(pool));
assert(JSON.stringify(pool) === JSON.stringify([{ color: 'Black / Blue-White / Brown / Kraft', qty: 10 }]),
  'a layer with no colour axis of its own passes the identity through unchanged');

console.log('\n=== Layer on top of an upstream that made only ONE combination ===');
addLot({ processId: 'PF-SOLO', qty: 10, lotNumber: 'SOLO', outputItemName: 'Chain Solo Part',
  colorBreakdown: [{ color: 'Black / Blue-White', qty: 10, countsTowardTotal: true, axisKey: 'seed' }] });
C.recalculateWarehousePool();
const p5 = mkProcess('Chain Solo Next', 6, 'Chain Solo Out', ['Chain Solo Part', 'Chain Carton']);
const p5Axes = axesOf(p5);
console.log('   axes:', JSON.stringify(p5Axes.map(a => ({ label: a.label, colors: a.colors }))));
assert(p5Axes.some(a => /solo part/i.test(a.label)),
  'a single-coloured upstream item is STILL an axis when its colour is a composite — the chain must not truncate');
addLot({ processId: p5, qty: 6, lotNumber: 'C5', outputItemName: 'Chain Solo Out',
  colorBreakdown: [
    { color: 'Black / Blue-White', qty: 6, countsTowardTotal: true, axisKey: axisKey(p5, /solo part/i) },
    { color: 'Kraft', qty: 6, countsTowardTotal: false, axisKey: axisKey(p5, /carton/i) }] });
C.recalculateWarehousePool();
pool = bucketsFor('Chain Solo Out');
console.log('  ', JSON.stringify(pool));
assert(JSON.stringify(pool) === JSON.stringify([{ color: 'Black / Blue-White / Kraft', qty: 6 }]),
  'the inherited history survives and the new layer appends to it');

console.log('\n=== A single-coloured RAW input is still NOT an axis (unchanged) ===');
addLot({ processId: 'PF-FIXED', qty: 50, lotNumber: 'FIX', outputItemName: 'Chain Fixed Bolt',
  colorBreakdown: [{ color: 'Zinc', qty: 50, countsTowardTotal: true, axisKey: 'seed' }] });
C.recalculateWarehousePool();
const p6 = mkProcess('Chain Bolted', 7, 'Chain Bolted Out', ['Chain Fixed Bolt', 'Chain Carton']);
const p6Axes = axesOf(p6);
console.log('   axes:', JSON.stringify(p6Axes.map(a => a.label)));
assert(!p6Axes.some(a => /fixed bolt/i.test(a.label)),
  'a fixed single-colour raw input stays out of the checklist, exactly as before');

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
