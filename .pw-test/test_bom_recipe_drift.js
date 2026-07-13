/**
 * Standalone Node harness (same mock-SpreadsheetApp pattern as
 * test_bulk_delete_row_guard.js) covering Fix #4 — the read-only BOM vs
 * Process Components reconciliation diagnostic
 * (module_bom.js#getBomProcessComponentsDrift), one of the 9 architectural
 * gaps verified+fixed on 2026-07-13 (see
 * verification_2026_07_13_architectural_gaps in project memory).
 *
 * BOM.qtyPerProduct (costing) and Process Components.qtyPerUnit (actual
 * shop-floor consumption) are two independent, legitimately-editable-apart
 * sources of truth — this diagnostic surfaces drift between them for a
 * human to judge, never blocks a save or auto-fixes anything.
 *
 * Run: node .pw-test/test_bom_recipe_drift.js
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

['config.js', 'utils.js', 'module_process.js', 'module_production.js', 'module_bom.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});

vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.PROCESS_COL = PROCESS_COL;
  global.PROCESS_COMPONENTS_COL = PROCESS_COMPONENTS_COL;
  global.BOM_COL = BOM_COL;
`, ctx, { filename: 'expose.js' });

const { APP_CONFIG, PROCESS_COL, PROCESS_COMPONENTS_COL, BOM_COL, getBomProcessComponentsDrift } = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); }
}

// Seed one final-stage Process (PRC-1) that Product PRD-1's BOM references
// via PROCESS_GROUP, so getBOMFinalStageProcessMap can resolve it.
const procSheet = ss.addSheet(APP_CONFIG.SHEETS.PROCESS_MASTER);
procSheet._set(2, PROCESS_COL.PROCESS_ID, 'PRC-1');
procSheet._set(2, PROCESS_COL.PROCESS_NAME, 'Final Packing');
procSheet._set(2, PROCESS_COL.SEQUENCE, 1);
procSheet._set(2, PROCESS_COL.LOT_PREFIX, 'PK');
procSheet._set(2, PROCESS_COL.IS_FINAL_STAGE, true);
procSheet._set(2, PROCESS_COL.ACTIVE, true);
procSheet._set(2, PROCESS_COL.OUTPUT_ITEM_NAME, 'Packed Bike');

// Process Components recipe for PRC-1: Paint qtyPerUnit=2, Screws qtyPerUnit=3.
const compSheet = ss.addSheet(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);
compSheet._set(2, PROCESS_COMPONENTS_COL.PROCESS_ID, 'PRC-1');
compSheet._set(2, PROCESS_COMPONENTS_COL.ITEM_NAME, 'Paint');
compSheet._set(2, PROCESS_COMPONENTS_COL.QTY_PER_UNIT, 2);
compSheet._set(2, PROCESS_COMPONENTS_COL.SOURCE_TYPE, 'ITEM');
compSheet._set(2, PROCESS_COMPONENTS_COL.COLOR_GROUP, 'COMMON');
compSheet._set(3, PROCESS_COMPONENTS_COL.PROCESS_ID, 'PRC-1');
compSheet._set(3, PROCESS_COMPONENTS_COL.ITEM_NAME, 'Screws');
compSheet._set(3, PROCESS_COMPONENTS_COL.QTY_PER_UNIT, 3);
compSheet._set(3, PROCESS_COMPONENTS_COL.SOURCE_TYPE, 'ITEM');
compSheet._set(3, PROCESS_COMPONENTS_COL.COLOR_GROUP, 'COMMON');
// A POOL-sourced recipe row -- must be excluded from comparison (BOM has
// no sourceType concept; only ITEM-sourced recipe rows are comparable).
compSheet._set(4, PROCESS_COMPONENTS_COL.PROCESS_ID, 'PRC-1');
compSheet._set(4, PROCESS_COMPONENTS_COL.ITEM_NAME, 'Painted Frame');
compSheet._set(4, PROCESS_COMPONENTS_COL.QTY_PER_UNIT, 1);
compSheet._set(4, PROCESS_COMPONENTS_COL.SOURCE_TYPE, 'POOL');
compSheet._set(4, PROCESS_COMPONENTS_COL.COLOR_GROUP, 'COMMON');

// BOM for PRD-1: Paint qtyPerProduct=2 (MATCHES recipe -> no drift),
// Screws qtyPerProduct=5 (recipe says 3 -> DRIFT), Glue (not in the
// recipe at all -> not this diagnostic's concern, no entry).
const bomSheet = ss.addSheet(APP_CONFIG.SHEETS.BOM);
bomSheet._set(2, BOM_COL.PRODUCT_ID, 'PRD-1');
bomSheet._set(2, BOM_COL.PRODUCT_NAME, 'Test Bike');
bomSheet._set(2, BOM_COL.ITEM_NAME, 'Paint');
bomSheet._set(2, BOM_COL.QTY_PER_PRODUCT, 2);
bomSheet._set(2, BOM_COL.PROCESS_GROUP, 'PRC-1');
bomSheet._set(3, BOM_COL.PRODUCT_ID, 'PRD-1');
bomSheet._set(3, BOM_COL.PRODUCT_NAME, 'Test Bike');
bomSheet._set(3, BOM_COL.ITEM_NAME, 'Screws');
bomSheet._set(3, BOM_COL.QTY_PER_PRODUCT, 5);
bomSheet._set(3, BOM_COL.PROCESS_GROUP, 'PRC-1');
bomSheet._set(4, BOM_COL.PRODUCT_ID, 'PRD-1');
bomSheet._set(4, BOM_COL.PRODUCT_NAME, 'Test Bike');
bomSheet._set(4, BOM_COL.ITEM_NAME, 'Glue');
bomSheet._set(4, BOM_COL.QTY_PER_PRODUCT, 1);
bomSheet._set(4, BOM_COL.PROCESS_GROUP, 'PRC-1');

console.log('\n=== Fix #4: getBomProcessComponentsDrift ===');
{
  const res = getBomProcessComponentsDrift();
  assert(res.success, 'call succeeds: ' + res.message);

  const drift = res.data || [];
  assert(drift.length === 1, `exactly 1 drifted item found (Screws) (got ${drift.length}: ${JSON.stringify(drift)})`);

  const entry = drift[0];
  if (entry) {
    assert(entry.itemName === 'Screws', `drift entry is for Screws (got ${entry.itemName})`);
    assert(entry.productId === 'PRD-1', `drift entry names the product (got ${entry.productId})`);
    assert(entry.bomQtyPerProduct === 5, `BOM qty reported as 5 (got ${entry.bomQtyPerProduct})`);
    assert(entry.recipeQtyPerUnit === 3, `recipe qty reported as 3 (got ${entry.recipeQtyPerUnit})`);
  }

  assert(!drift.some(d => d.itemName === 'Paint'), 'Paint (BOM 2 == recipe 2) is NOT reported -- matching quantities are not drift');
  assert(!drift.some(d => d.itemName === 'Glue'), "Glue (not in the process recipe at all) is NOT reported -- outside this diagnostic's scope");
  assert(!drift.some(d => d.itemName === 'Painted Frame'), 'the POOL-sourced recipe row is never compared against BOM (no sourceType concept there)');
}

console.log('\n=== Fix #4: no drift at all reports success with an empty list ===');
{
  bomSheet._set(3, BOM_COL.QTY_PER_PRODUCT, 3); // fix the Screws drift
  const res = getBomProcessComponentsDrift();
  assert(res.success, 'call succeeds: ' + res.message);
  assert((res.data || []).length === 0, `no drift once quantities match (got ${(res.data || []).length})`);
  assert(/no drift found/i.test(res.message || ''), `message confirms no drift (got "${res.message}")`);
}

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
