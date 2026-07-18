/**
 * Standalone Node harness (same mock-SpreadsheetApp pattern as
 * test_bom_recipe_drift.js) covering a debugging-session fix to
 * module_bom.js#getBOMData: a non-blank Color on a BOM component row is an
 * ALTERNATIVE that only applies when that specific color is produced (see
 * ensureBOMColorColumn's header comment) — a unit is ever only one color,
 * so two different colors' rows are mutually exclusive and must never both
 * be summed into the same total. The old code summed every row regardless
 * of color, inflating a multi-color product's displayed Material Cost by
 * every extra color present.
 *
 * Fix: totalCost/totalQty = common (blank-color) rows + only the FIRST
 * color's rows (a single representative headline number); a new
 * `colorCosts` array carries the full per-color breakdown.
 *
 * Run: node .pw-test/test_bom_color_cost.js
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

const fakeCache = {
  _store: {},
  get(k) { return Object.prototype.hasOwnProperty.call(this._store, k) ? this._store[k] : null; },
  put(k, v) { this._store[k] = v; },
  remove(k) { delete this._store[k]; }
};

const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush: () => {} },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  CacheService: { getScriptCache: () => fakeCache },
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
  global.BOM_COL = BOM_COL;
  global.BOM_AUTH_CACHE_PREFIX = BOM_AUTH_CACHE_PREFIX;
`, ctx, { filename: 'expose.js' });

const { APP_CONFIG, BOM_COL, BOM_AUTH_CACHE_PREFIX, getBOMData } = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); }
}

// Bypass the password gate directly via the cache token _requireBOMAccess checks.
const TOKEN = 'test-token';
fakeCache.put(BOM_AUTH_CACHE_PREFIX + TOKEN, '1');

// Product PRD-1: one common Frame row (rate 500, qty 1 -> 500), plus a
// Red-specific Paint row (rate 50, qty 1) and a Blue-specific Paint row
// (rate 55, qty 1) -- Red and Blue are mutually exclusive alternatives, a
// unit is only ever produced in one of them.
const bomSheet = ss.addSheet(APP_CONFIG.SHEETS.BOM);
bomSheet._set(2, BOM_COL.PRODUCT_ID, 'PRD-1');
bomSheet._set(2, BOM_COL.PRODUCT_NAME, 'Test Bike');
bomSheet._set(2, BOM_COL.ITEM_NAME, 'Frame');
bomSheet._set(2, BOM_COL.RATE, 500);
bomSheet._set(2, BOM_COL.QTY_PER_PRODUCT, 1);
bomSheet._set(2, BOM_COL.COLOR, '');

bomSheet._set(3, BOM_COL.PRODUCT_ID, 'PRD-1');
bomSheet._set(3, BOM_COL.PRODUCT_NAME, 'Test Bike');
bomSheet._set(3, BOM_COL.ITEM_NAME, 'Paint');
bomSheet._set(3, BOM_COL.RATE, 50);
bomSheet._set(3, BOM_COL.QTY_PER_PRODUCT, 1);
bomSheet._set(3, BOM_COL.COLOR, 'Red');

bomSheet._set(4, BOM_COL.PRODUCT_ID, 'PRD-1');
bomSheet._set(4, BOM_COL.PRODUCT_NAME, 'Test Bike');
bomSheet._set(4, BOM_COL.ITEM_NAME, 'Paint');
bomSheet._set(4, BOM_COL.RATE, 55);
bomSheet._set(4, BOM_COL.QTY_PER_PRODUCT, 1);
bomSheet._set(4, BOM_COL.COLOR, 'Blue');

const res = getBOMData(TOKEN);
assert(res.success, 'getBOMData succeeds: ' + res.message);

const bom = (res.data || []).find(p => p.productId === 'PRD-1');
assert(!!bom, 'PRD-1 found in getBOMData results');
assert(bom && bom.components.length === 3, `all 3 rows still listed individually (got ${bom && bom.components.length})`);

console.log('\n=== Fix: totalCost no longer sums both colors\' rows together ===');
assert(bom && bom.totalCost === 550, `totalCost is common(500) + first color Red(50) = 550, NOT 500+50+55=605 (got ${bom && bom.totalCost})`);
assert(bom && bom.totalQty === 2, `totalQty is common(1) + Red(1) = 2, NOT 3 (got ${bom && bom.totalQty})`);
assert(bom && bom.grandTotal === 550, `grandTotal matches totalCost when no additional costs (got ${bom && bom.grandTotal})`);

console.log('\n=== Fix: per-color breakdown is available so no color\'s true cost is lost ===');
assert(bom && Array.isArray(bom.colorCosts) && bom.colorCosts.length === 2, `colorCosts has one entry per color (got ${bom && JSON.stringify(bom.colorCosts)})`);
const redCost = bom && bom.colorCosts.find(c => c.color === 'Red');
const blueCost = bom && bom.colorCosts.find(c => c.color === 'Blue');
assert(!!redCost && redCost.totalCost === 550, `Red total = common(500) + Red(50) = 550 (got ${redCost && redCost.totalCost})`);
assert(!!blueCost && blueCost.totalCost === 555, `Blue total = common(500) + Blue(55) = 555 (got ${blueCost && blueCost.totalCost})`);

console.log('\n=== Regression: a product with NO color rows behaves exactly as before ===');
bomSheet._set(5, BOM_COL.PRODUCT_ID, 'PRD-2');
bomSheet._set(5, BOM_COL.PRODUCT_NAME, 'Colorless Widget');
bomSheet._set(5, BOM_COL.ITEM_NAME, 'Bolt');
bomSheet._set(5, BOM_COL.RATE, 10);
bomSheet._set(5, BOM_COL.QTY_PER_PRODUCT, 4);
bomSheet._set(5, BOM_COL.COLOR, '');
const res2 = getBOMData(TOKEN);
const bom2 = (res2.data || []).find(p => p.productId === 'PRD-2');
assert(!!bom2 && bom2.totalCost === 40, `no-color product totalCost unaffected (got ${bom2 && bom2.totalCost})`);
assert(!!bom2 && Array.isArray(bom2.colorCosts) && bom2.colorCosts.length === 0, `no-color product has empty colorCosts (got ${bom2 && JSON.stringify(bom2.colorCosts)})`);

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
