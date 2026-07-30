/**
 * "Every process should have a primary axis."
 *
 * A blank Primary Color Axis cell used to mean "legacy: sum every checked
 * color across every axis" — which double-counts a multi-axis lot's
 * non-primary rows (a 10-unit lot with a Frame row and a Rim row saved as 20)
 * unless the operator noticed the warning on the Production form and picked a
 * group by hand. Now every process that HAS a color axis resolves and stores
 * one: its first axis in recipe order (_defaultPrimaryColorAxisLabel,
 * module_process.js), applied by saveProcess, saveProduction,
 * getProcessColorAxes and backfillProcessPrimaryColorAxes alike.
 *
 * The risky half of that change is the SINGLE-axis process (Test 3). Its
 * checklist renders the legacy pool-signature grouping, not the per-axis one,
 * so its rows submit group ids like "g1"/"other" in the same axisKey field a
 * real axis key travels in. Comparing those straight to the primary axis key
 * makes every row read as "not primary" -> quantity 0 -> the "At least one
 * ... color is required" guard rejects the save, i.e. giving such a process a
 * primary axis would have made it unable to log any lot at all.
 *
 * Uses the strict FakeRange from verify_process_save_row_width.js (setValues
 * enforces the range's declared width, as Apps Script does).
 *
 * Run: node .pw-test/verify_every_process_has_primary_axis.js
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
    if (!Array.isArray(v) || v.length !== this.numRows) {
      throw new Error(`The number of rows in the data does not match the number of rows in the range. The data has ${Array.isArray(v) ? v.length : 0} but the range has ${this.numRows}.`);
    }
    v.forEach(ra => {
      if (!Array.isArray(ra) || ra.length !== this.numCols) {
        throw new Error(`The number of columns in the data does not match the number of columns in the range. The data has ${Array.isArray(ra) ? ra.length : 0} but the range has ${this.numCols}.`);
      }
    });
    v.forEach((ra, r) => ra.forEach((val, c) => this.sheet._set(this.row + r, this.col + c, val)));
    return this;
  }
  setValue(v) {
    if (this.numRows !== 1 || this.numCols !== 1) { for (let r = 0; r < this.numRows; r++) for (let c = 0; c < this.numCols; c++) this.sheet._set(this.row + r, this.col + c, v); return this; }
    this.sheet._set(this.row, this.col, v); return this;
  }
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
vm.runInContext('global.APP_CONFIG=APP_CONFIG; global.PROCESS_COL=PROCESS_COL; global.MASTER_DATA_CACHE_KEYS=MASTER_DATA_CACHE_KEYS;', ctx, { filename: 'expose.js' });
const C = ctx;
const { APP_CONFIG, PROCESS_COL } = C;

let failures = 0;
function assert(cond, msg) { if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); } }

const FRAME = 'Painted Frame';
const RIM = 'Fitted Rim';
const SEAT = 'Fitted Seat';

// ── Seed the Warehouse Pool ────────────────────────────────────────────────
// FRAME and RIM each carry 2+ colors, so each is its own Color Axis. SEAT has
// exactly one, which is never an axis (a component that's always Black is not
// a per-output-color choice) — the ingredient for a genuine 1-axis process.
const poolSheet = ss.addSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
// Re-seeded before every axis-dependent step: saveProduction recalculates the
// whole Warehouse Pool from the Production sheet, which drops rows that no
// logged lot produced — i.e. all of these, since the upstream processes that
// would have produced them don't exist in this fixture. Without the re-seed,
// every process silently loses its axes after the first lot is logged.
function seedPool() {
  const sh = ss.getSheetByName(APP_CONFIG.SHEETS.WAREHOUSE_POOL) || ss.addSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
  sh.rows = [];
  sh.appendRow(['Output Item Name', 'Process ID', 'Product Tag', 'Produced Qty', 'Consumed Qty', 'Available Qty', 'Color']);
  [
    [FRAME, 'PRC-F', 'Blue-White'], [FRAME, 'PRC-F', 'Orange-White'],
    [RIM, 'PRC-R', 'BCP'], [RIM, 'PRC-R', 'Black'],
    [SEAT, 'PRC-S', 'Black']
  ].forEach(r => sh.appendRow([r[0], r[1], '', 100, 0, 100, r[2]]));
}
seedPool();

const procSheet = () => ss.getSheetByName(APP_CONFIG.SHEETS.PROCESS_MASTER);
const rowOf = pid => (C.getProcessData(false).data || []).find(p => p.processId === pid);
const lotOf = res => (C.getProductionData().data || []).find(l => l.lotNumber === (res.data && res.data.lotNumber));

console.log('=== Test 1: a 2-axis process saved with NO primary axis stores one anyway ===');
const twoAxis = C.saveProcess({
  processName: 'Fitting Frame', sequence: 5, lotPrefix: 'FF2', outputItemName: 'Fitted Frame Assembly',
  isFinalStage: false, active: true, remarks: '',
  // Frame rows come first, so Frame is the first axis in recipe order.
  components: JSON.stringify([
    { itemName: FRAME, sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' },
    { itemName: RIM, sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' }
  ])
});
assert(twoAxis.success, 'process saved: ' + twoAxis.message);
const twoAxisId = twoAxis.data && twoAxis.data.processId;
assert(rowOf(twoAxisId) && rowOf(twoAxisId).primaryColorAxis === FRAME,
  `blank primary defaulted to the first recipe axis "${FRAME}" (got "${rowOf(twoAxisId) && rowOf(twoAxisId).primaryColorAxis}")`);
{
  const r = C.getProcessColorAxes(twoAxisId);
  assert(r.success && r.data.primaryColorAxis === FRAME,
    `getProcessColorAxes reports it (got "${r.data && r.data.primaryColorAxis}")`);
  assert(r.data && r.data.primaryAxisKey === 'pool:' + FRAME.toLowerCase(),
    `primaryAxisKey resolves (got "${r.data && r.data.primaryAxisKey}")`);
  assert(r.data && r.data.primaryIsDefault === false,
    'primaryIsDefault false once the value is actually stored');
}

console.log('\n=== Test 2: that lot counts ONLY the primary axis, not every checked row ===');
{
  const res = C.saveProduction({
    processId: twoAxisId, assignedTo: 'Sanjay', status: 'Pending',
    colorBreakdown: JSON.stringify([
      { color: 'Blue-White', qty: 10, axisKey: 'pool:' + FRAME.toLowerCase(), countsTowardTotal: true },
      { color: 'BCP', qty: 10, axisKey: 'pool:' + RIM.toLowerCase(), countsTowardTotal: false }
    ]),
    componentsConsumed: JSON.stringify([
      { itemName: FRAME, sourceType: 'POOL', qty: 10, colorGroup: 'Blue-White' },
      { itemName: RIM, sourceType: 'POOL', qty: 10, colorGroup: 'BCP' }
    ])
  });
  assert(res.success, 'lot saves: ' + res.message);
  const lot = lotOf(res);
  assert(!!lot && lot.qty === 10, `quantity is 10 (Frame only), not 20 (got ${lot && lot.qty})`);
}

seedPool();
console.log('\n=== Test 3: a SINGLE-axis process can still log lots (the regression this could cause) ===');
const oneAxis = C.saveProcess({
  processName: 'Seat Fitting', sequence: 6, lotPrefix: 'SF2', outputItemName: 'Seated Frame',
  isFinalStage: false, active: true, remarks: '',
  components: JSON.stringify([
    { itemName: FRAME, sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' },
    { itemName: SEAT, sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' }
  ])
});
assert(oneAxis.success, 'process saved: ' + oneAxis.message);
const oneAxisId = oneAxis.data && oneAxis.data.processId;
assert(rowOf(oneAxisId) && rowOf(oneAxisId).primaryColorAxis === FRAME,
  `single-axis process also gets a primary (got "${rowOf(oneAxisId) && rowOf(oneAxisId).primaryColorAxis}")`);
{
  // A 1-axis process renders the LEGACY pool-signature checklist, whose rows
  // carry group ids ("g1"), not axis keys — the shape that used to be read as
  // "belongs to no primary axis" and rejected outright.
  const res = C.saveProduction({
    processId: oneAxisId, assignedTo: 'Sanjay', status: 'Pending',
    colorBreakdown: JSON.stringify([
      { color: 'Blue-White', qty: 7, axisKey: 'g1', countsTowardTotal: true },
      { color: 'Orange-White', qty: 3, axisKey: 'g1', countsTowardTotal: true }
    ]),
    componentsConsumed: JSON.stringify([
      { itemName: FRAME, sourceType: 'POOL', qty: 10, colorGroup: 'Blue-White' }
    ])
  });
  assert(res.success, 'legacy-grouped lot still saves: ' + res.message);
  const lot = lotOf(res);
  assert(!!lot && lot.qty === 10, `quantity is the full 7+3=10, not 0 (got ${lot && lot.qty})`);
}

console.log('\n=== Test 4: a process with NO axes at all is untouched (still blank, still sums everything) ===');
const noAxis = C.saveProcess({
  processName: 'Plain Packing', sequence: 7, lotPrefix: 'PP2', outputItemName: 'Packed Cycle',
  isFinalStage: false, active: true, remarks: '',
  components: JSON.stringify([
    { itemName: 'Red Sticker', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Red' },
    { itemName: 'Blue Sticker', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Blue' }
  ])
});
const noAxisId = noAxis.data && noAxis.data.processId;
assert(rowOf(noAxisId) && rowOf(noAxisId).primaryColorAxis === '',
  `no axes -> nothing to be primary, stays blank (got "${rowOf(noAxisId) && rowOf(noAxisId).primaryColorAxis}")`);
{
  const res = C.saveProduction({
    processId: noAxisId, assignedTo: 'Sanjay', status: 'Pending',
    colorBreakdown: JSON.stringify([{ color: 'Red', qty: 6 }, { color: 'Blue', qty: 4 }]),
    componentsConsumed: JSON.stringify([{ itemName: 'Red Sticker', sourceType: 'ITEM', qty: 6, colorGroup: 'Red' }])
  });
  assert(res.success, 'legacy lot saves: ' + res.message);
  const lot = lotOf(res);
  assert(!!lot && lot.qty === 10, `legacy sum-every-color unchanged (6+4=10, got ${lot && lot.qty})`);
}

seedPool();
console.log('\n=== Test 5: a pre-existing blank row reports (and backfills to) the default ===');
{
  // Blank the cell straight on the sheet, exactly as a row created before this
  // feature looks.
  const ids = procSheet().getRange(2, PROCESS_COL.PROCESS_ID, procSheet().getLastRow() - 1, 1).getValues();
  const rowIdx = ids.findIndex(r => String(r[0]).trim() === twoAxisId) + 2;
  procSheet().getRange(rowIdx, PROCESS_COL.PRIMARY_COLOR_AXIS).setValue('');
  C.invalidateListCache(C.MASTER_DATA_CACHE_KEYS.PROCESS_ALL, C.MASTER_DATA_CACHE_KEYS.PROCESS_ACTIVE);

  const r = C.getProcessColorAxes(twoAxisId);
  assert(r.success && r.data.primaryColorAxis === FRAME,
    `blank row still reports a primary axis (got "${r.data && r.data.primaryColorAxis}")`);
  assert(r.data && r.data.primaryIsDefault === true, 'flagged primaryIsDefault so the UI can say it was auto-picked');
  assert(r.data && r.data.savedPrimaryColorAxis === '', 'savedPrimaryColorAxis still reports the raw (blank) cell');

  const back = C.backfillProcessPrimaryColorAxes();
  assert(back.success && back.data.filled === 1, `backfill filled the one blank row (got ${back.data && back.data.filled})`);
  assert(rowOf(twoAxisId) && rowOf(twoAxisId).primaryColorAxis === FRAME,
    `cell now holds "${FRAME}" (got "${rowOf(twoAxisId) && rowOf(twoAxisId).primaryColorAxis}")`);

  const again = C.backfillProcessPrimaryColorAxes();
  assert(again.success && again.data.filled === 0, `re-running fills nothing (got ${again.data && again.data.filled})`);
  assert(rowOf(noAxisId) && rowOf(noAxisId).primaryColorAxis === '',
    'the no-axis process was skipped, not blanked or guessed at');
}

seedPool();
console.log('\n=== Test 6: an explicit choice still wins over the default ===');
{
  const res = C.saveProcess({
    processId: twoAxisId, processName: 'Fitting Frame', sequence: 5, lotPrefix: 'FF2',
    outputItemName: 'Fitted Frame Assembly', isFinalStage: false, active: true, remarks: '',
    primaryColorAxis: RIM,
    components: JSON.stringify([
      { itemName: FRAME, sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' },
      { itemName: RIM, sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' }
    ])
  });
  assert(res.success, 'process re-saved: ' + res.message);
  assert(rowOf(twoAxisId) && rowOf(twoAxisId).primaryColorAxis === RIM,
    `explicit "${RIM}" kept, not overwritten by the first-axis default (got "${rowOf(twoAxisId) && rowOf(twoAxisId).primaryColorAxis}")`);

  const lot = C.saveProduction({
    processId: twoAxisId, assignedTo: 'Sanjay', status: 'Pending',
    colorBreakdown: JSON.stringify([
      { color: 'Blue-White', qty: 10, axisKey: 'pool:' + FRAME.toLowerCase(), countsTowardTotal: false },
      { color: 'BCP', qty: 4, axisKey: 'pool:' + RIM.toLowerCase(), countsTowardTotal: true }
    ]),
    componentsConsumed: JSON.stringify([{ itemName: RIM, sourceType: 'POOL', qty: 4, colorGroup: 'BCP' }])
  });
  assert(lot.success, 'lot saves against the explicit primary: ' + lot.message);
  assert(lotOf(lot) && lotOf(lot).qty === 4, `quantity follows the Rim axis (got ${lotOf(lot) && lotOf(lot).qty})`);
}

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
