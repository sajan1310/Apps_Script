/**
 * Standalone Node harness (same mock-SpreadsheetApp pattern as
 * test_bulk_delete_row_guard.js) covering two more of the 9 architectural
 * gaps verified+fixed on 2026-07-13 (see
 * verification_2026_07_13_architectural_gaps in project memory):
 *
 * Fix #9 — deleteProduction now warns (non-blocking) when removing a
 * Completed lot's own Warehouse Pool credit would leave a bucket negative
 * because a downstream lot already consumed it — previously silent.
 *
 * Fix #3 (scoped) — a Process Components recipe row can now carry an
 * explicit Unit (PROCESS_COMPONENTS_COL.UNIT); when a Production lot's
 * componentsConsumed entry for an ITEM-sourced component carries that same
 * non-blank unit, module_stock.js#_getBilledAndConsumedQtyMaps converts it
 * to the item's own Base Unit via convertQtyToBaseUnit before netting into
 * Stock — same convertQtyToBaseUnit path Bill/PO/Return/Wastage/Issue
 * already use. A blank unit (every pre-existing recipe row) still means
 * "already in Base Unit," preserving old behavior exactly.
 *
 * Run: node .pw-test/test_production_pool_and_unit_fixes.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet; this.row = row; this.col = col; this.numRows = numRows; this.numCols = numCols;
  }
  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowArr = [];
      for (let c = 0; c < this.numCols; c++) rowArr.push(this.sheet._get(this.row + r, this.col + c));
      out.push(rowArr);
    }
    return out;
  }
  getValue() { return this.sheet._get(this.row, this.col); }
  setValues(values) {
    values.forEach((rowArr, r) => rowArr.forEach((val, c) => this.sheet._set(this.row + r, this.col + c, val)));
    return this;
  }
  setValue(v) { this.sheet._set(this.row, this.col, v); return this; }
  clearContent() {
    for (let r = 0; r < this.numRows; r++) for (let c = 0; c < this.numCols; c++) this.sheet._set(this.row + r, this.col + c, '');
    return this;
  }
  setFontWeight() { return this; }
  setBackground() { return this; }
}

class FakeSheet {
  constructor(name) { this.name = name; this.rows = []; }
  getName() { return this.name; }
  _ensureRow(r) { while (this.rows.length < r) this.rows.push([]); }
  _get(r, c) { this._ensureRow(r); const row = this.rows[r - 1]; return row[c - 1] === undefined ? '' : row[c - 1]; }
  _set(r, c, v) { this._ensureRow(r); const row = this.rows[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; }
  getLastRow() {
    for (let r = this.rows.length; r >= 1; r--) {
      if (this.rows[r - 1].some(v => v !== '' && v !== undefined && v !== null)) return r;
    }
    return 0;
  }
  getLastColumn() {
    let max = 0;
    this.rows.forEach(row => {
      for (let c = row.length; c >= 1; c--) {
        if (row[c - 1] !== '' && row[c - 1] !== undefined && row[c - 1] !== null) { max = Math.max(max, c); break; }
      }
    });
    return max;
  }
  getRange(row, col, numRows = 1, numCols = 1) { return new FakeRange(this, row, col, numRows, numCols); }
  appendRow(arr) { const r = this.getLastRow() + 1; arr.forEach((v, i) => this._set(r, i + 1, v)); }
  deleteRow(r) { this.rows.splice(r - 1, 1); }
  deleteRows(r, n) { this.rows.splice(r - 1, n); }
  insertRows(r, n) { for (let i = 0; i < n; i++) this.rows.splice(r - 1, 0, []); }
  insertColumnsAfter(afterPosition, howMany) {
    this.rows.forEach(row => { const blanks = new Array(howMany).fill(''); row.splice(afterPosition, 0, ...blanks); });
  }
}

class FakeSpreadsheet {
  constructor() { this.sheets = {}; }
  getSheetByName(name) { return this.sheets[name] || null; }
  addSheet(name) { const s = new FakeSheet(name); this.sheets[name] = s; return s; }
  insertSheet(name) { return this.addSheet(name); }
}

const ss = new FakeSpreadsheet();

const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush: () => {} },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {} }) },
  console,
  Logger: { log: () => {} },
  Utilities: { getUuid: () => 'uuid-' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);

['config.js', 'utils.js', 'module_units.js', 'module_items.js', 'module_process.js', 'module_production.js', 'module_warehouse.js', 'module_stock.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});

vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.PRODUCTION_COL = PRODUCTION_COL;
  global.UNITS_COL = UNITS_COL;
  global.ITEMS_COL = ITEMS_COL;
  global.STOCK_COL = STOCK_COL;
`, ctx, { filename: 'expose.js' });

const {
  APP_CONFIG, PRODUCTION_COL, UNITS_COL, ITEMS_COL, STOCK_COL,
  deleteProduction, recalculateWarehousePool, recalculateStock
} = ctx;
// getPoolAvailableQtyMap isn't destructured above (added after that binding
// list was written) — read it off ctx directly where needed instead.

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); }
}

console.log('\n=== Fix #9: deleteProduction warns when removing a credit leaves the pool negative ===');
{
  const prodSheet = ss.addSheet(APP_CONFIG.SHEETS.PRODUCTION);
  // Row 2: intermediate (untagged) lot crediting "Painted Frame" x10.
  prodSheet._set(2, PRODUCTION_COL.DATE, '01/01/2026');
  prodSheet._set(2, PRODUCTION_COL.QTY, 10);
  prodSheet._set(2, PRODUCTION_COL.STATUS, 'Completed');
  prodSheet._set(2, PRODUCTION_COL.PROCESS_ID, 'PRC-1');
  prodSheet._set(2, PRODUCTION_COL.LOT_NUMBER, 'LOT-1');
  prodSheet._set(2, PRODUCTION_COL.OUTPUT_ITEM_NAME, 'Painted Frame');
  prodSheet._set(2, PRODUCTION_COL.COMPONENTS_CONSUMED, '[]');
  prodSheet._set(2, PRODUCTION_COL.COLOR_BREAKDOWN, '');

  // Row 3: downstream lot consuming 8 of that credit (POOL-sourced).
  prodSheet._set(3, PRODUCTION_COL.DATE, '02/01/2026');
  prodSheet._set(3, PRODUCTION_COL.QTY, 8);
  prodSheet._set(3, PRODUCTION_COL.STATUS, 'Completed');
  prodSheet._set(3, PRODUCTION_COL.PROCESS_ID, 'PRC-2');
  prodSheet._set(3, PRODUCTION_COL.LOT_NUMBER, 'LOT-2');
  prodSheet._set(3, PRODUCTION_COL.OUTPUT_ITEM_NAME, 'Fitted Frame');
  prodSheet._set(3, PRODUCTION_COL.COMPONENTS_CONSUMED, JSON.stringify([
    { itemName: 'Painted Frame', sourceType: 'POOL', qty: 8, colorGroup: 'COMMON' }
  ]));
  prodSheet._set(3, PRODUCTION_COL.COLOR_BREAKDOWN, '');

  const recalc1 = recalculateWarehousePool();
  assert(recalc1.success, 'initial recalculateWarehousePool succeeds: ' + recalc1.message);

  // Available should now be 10 - 8 = 2. Deleting row 2 (the x10 credit)
  // would leave it at 2 - 10 = -8 -> the warning must fire. No expected*
  // guard args passed -- that optional check is exercised separately.
  const res = deleteProduction(2);
  assert(res.success, 'delete itself still succeeds (non-blocking warning): ' + res.message);
  assert(/leaves the warehouse pool negative/i.test(res.message || ''), `message carries the pool-negative warning (got "${res.message}")`);
  assert(/painted frame/i.test(res.message || ''), `warning names the affected item (got "${res.message}")`);
}

console.log('\n=== Fix #9: deleteProduction stays silent when nothing goes negative ===');
{
  const prodSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PRODUCTION);
  // Row 3 (LOT-2, still present) already consumed 8 -- re-seed a fresh
  // untagged credit lot with plenty of headroom (20) and delete it; 20 was
  // never at risk of going negative on its own since nothing else draws on
  // this new bucket.
  const row = prodSheet.getLastRow() + 1;
  prodSheet._set(row, PRODUCTION_COL.DATE, '03/01/2026');
  prodSheet._set(row, PRODUCTION_COL.QTY, 20);
  prodSheet._set(row, PRODUCTION_COL.STATUS, 'Completed');
  prodSheet._set(row, PRODUCTION_COL.PROCESS_ID, 'PRC-3');
  prodSheet._set(row, PRODUCTION_COL.LOT_NUMBER, 'LOT-3');
  prodSheet._set(row, PRODUCTION_COL.OUTPUT_ITEM_NAME, 'Untouched Item');
  prodSheet._set(row, PRODUCTION_COL.COMPONENTS_CONSUMED, '[]');
  prodSheet._set(row, PRODUCTION_COL.COLOR_BREAKDOWN, '');
  recalculateWarehousePool();

  const res = deleteProduction(row);
  assert(res.success, 'delete succeeds: ' + res.message);
  assert(!/leaves the warehouse pool negative/i.test(res.message || ''), `no warning when nothing goes negative (got "${res.message}")`);
}

console.log('\n=== Fix #3: componentsConsumed with an explicit Unit converts to the item\'s Base Unit before netting into Stock ===');
{
  const unitsSheet = ss.addSheet(APP_CONFIG.SHEETS.UNITS);
  unitsSheet._set(2, UNITS_COL.UNIT_NAME, 'Pcs');
  unitsSheet._set(2, UNITS_COL.FAMILY, 'Count');
  unitsSheet._set(2, UNITS_COL.FACTOR_TO_BASE, 1);
  unitsSheet._set(3, UNITS_COL.UNIT_NAME, 'Dozen');
  unitsSheet._set(3, UNITS_COL.FAMILY, 'Count');
  unitsSheet._set(3, UNITS_COL.FACTOR_TO_BASE, 12);

  const itemsSheet = ss.addSheet(APP_CONFIG.SHEETS.ITEMS);
  itemsSheet._set(2, ITEMS_COL.ITEM_NAME, 'Screws');
  itemsSheet._set(2, ITEMS_COL.SIZE, '');
  itemsSheet._set(2, ITEMS_COL.BASE_UNIT, 'Dozen');
  itemsSheet._set(2, ITEMS_COL.PURCHASE_UNIT, 'Pcs');
  itemsSheet._set(2, ITEMS_COL.WEIGHT_PER_BASE_UNIT, 0);
  // Second item with NO explicit recipe unit later -- must behave exactly
  // as before (qty used as-is, no conversion).
  itemsSheet._set(3, ITEMS_COL.ITEM_NAME, 'Bolts');
  itemsSheet._set(3, ITEMS_COL.SIZE, '');
  itemsSheet._set(3, ITEMS_COL.BASE_UNIT, 'Pcs');
  itemsSheet._set(3, ITEMS_COL.PURCHASE_UNIT, 'Pcs');
  itemsSheet._set(3, ITEMS_COL.WEIGHT_PER_BASE_UNIT, 0);

  const stockSheet = ss.addSheet(APP_CONFIG.SHEETS.STOCK);
  stockSheet._set(2, STOCK_COL.ITEM_NAME, 'Screws');
  stockSheet._set(2, STOCK_COL.SIZE, '');
  stockSheet._set(2, STOCK_COL.INITIAL_STOCK, 10); // 10 Dozen
  stockSheet._set(3, STOCK_COL.ITEM_NAME, 'Bolts');
  stockSheet._set(3, STOCK_COL.SIZE, '');
  stockSheet._set(3, STOCK_COL.INITIAL_STOCK, 50); // 50 Pcs

  const prodSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PRODUCTION);
  const row = prodSheet.getLastRow() + 1;
  prodSheet._set(row, PRODUCTION_COL.DATE, '04/01/2026');
  prodSheet._set(row, PRODUCTION_COL.QTY, 1);
  prodSheet._set(row, PRODUCTION_COL.STATUS, 'Completed');
  prodSheet._set(row, PRODUCTION_COL.PROCESS_ID, 'PRC-4');
  prodSheet._set(row, PRODUCTION_COL.LOT_NUMBER, 'LOT-4');
  prodSheet._set(row, PRODUCTION_COL.OUTPUT_ITEM_NAME, 'Some Output');
  prodSheet._set(row, PRODUCTION_COL.COMPONENTS_CONSUMED, JSON.stringify([
    // 24 Pcs of Screws -> converts to 2 Dozen (the item's own Base Unit).
    { itemName: 'Screws', size: '', sourceType: 'ITEM', qty: 24, unit: 'Pcs', colorGroup: 'COMMON' },
    // 3 Bolts, no unit at all -> legacy behavior, used as-is (3, no conversion).
    { itemName: 'Bolts', size: '', sourceType: 'ITEM', qty: 3, colorGroup: 'COMMON' }
  ]));
  prodSheet._set(row, PRODUCTION_COL.COLOR_BREAKDOWN, '');

  recalculateStock();

  const screwsCurrent = Number(stockSheet._get(2, STOCK_COL.CURRENT_STOCK));
  assert(screwsCurrent === 8, `Screws: 10 Dozen initial - (24 Pcs converted to 2 Dozen) = 8 Dozen (got ${screwsCurrent})`);

  const boltsCurrent = Number(stockSheet._get(3, STOCK_COL.CURRENT_STOCK));
  assert(boltsCurrent === 47, `Bolts: 50 Pcs initial - 3 Pcs (no unit -> no conversion, legacy behavior) = 47 (got ${boltsCurrent})`);
}

console.log('\n=== Fix: a POOL-sourced componentsConsumed entry with an explicit Unit now converts before debiting the Warehouse Pool (previously silently ignored) ===');
{
  // Units/Items sheets already seeded above: Dozen -> factor 12, Pcs -> factor 1,
  // and an Items Master row isn't required for the pool item itself (WIP output
  // items fall back to {baseUnit: 'Pcs', ...} same as an unregistered item).
  const prodSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PRODUCTION);

  // Upstream credit: 10 units of "Painted Widget" WIP output.
  let row = prodSheet.getLastRow() + 1;
  prodSheet._set(row, PRODUCTION_COL.DATE, '05/01/2026');
  prodSheet._set(row, PRODUCTION_COL.QTY, 10);
  prodSheet._set(row, PRODUCTION_COL.STATUS, 'Completed');
  prodSheet._set(row, PRODUCTION_COL.PROCESS_ID, 'PRC-5');
  prodSheet._set(row, PRODUCTION_COL.LOT_NUMBER, 'LOT-5');
  prodSheet._set(row, PRODUCTION_COL.OUTPUT_ITEM_NAME, 'Painted Widget');
  prodSheet._set(row, PRODUCTION_COL.COMPONENTS_CONSUMED, '[]');
  prodSheet._set(row, PRODUCTION_COL.COLOR_BREAKDOWN, '');

  // Downstream lot draws 1 Dozen of "Painted Widget" from the pool -> must
  // debit 12 (Dozen -> Pcs, factor 12), not 1.
  row = prodSheet.getLastRow() + 1;
  prodSheet._set(row, PRODUCTION_COL.DATE, '06/01/2026');
  prodSheet._set(row, PRODUCTION_COL.QTY, 1);
  prodSheet._set(row, PRODUCTION_COL.STATUS, 'Completed');
  prodSheet._set(row, PRODUCTION_COL.PROCESS_ID, 'PRC-6');
  prodSheet._set(row, PRODUCTION_COL.LOT_NUMBER, 'LOT-6');
  prodSheet._set(row, PRODUCTION_COL.OUTPUT_ITEM_NAME, 'Assembled Widget');
  prodSheet._set(row, PRODUCTION_COL.COMPONENTS_CONSUMED, JSON.stringify([
    { itemName: 'Painted Widget', sourceType: 'POOL', qty: 1, unit: 'Dozen', colorGroup: 'COMMON' }
  ]));
  prodSheet._set(row, PRODUCTION_COL.COLOR_BREAKDOWN, '');

  const recalc = recalculateWarehousePool();
  assert(recalc.success, 'recalculateWarehousePool succeeds: ' + recalc.message);

  const poolMap = ctx.getPoolAvailableQtyMap();
  const entry = poolMap['painted widget'];
  assert(!!entry, 'Painted Widget bucket exists in pool map');
  assert(entry && entry.total === -2, `Painted Widget: 10 produced - 12 (1 Dozen converted to Pcs) consumed = -2 available (got ${entry && entry.total})`);
}

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
