/**
 * Standalone Node harness covering a debugging-session fix: merging two
 * Item Master rows (mergeItemEdit / _mergeItemIdentities, used by
 * mergeSelectedItems and autoFixTruncatedDuplicateItems) used to copy a
 * source item's vendor rate and Stock quantities straight into the target
 * row with no check that both used the same unit. If the source was
 * quoted/tracked in a different Purchase/Base Unit than the target, the
 * merged numbers were silently wrong (e.g. a rate of 1200 quoted per Dozen
 * copied onto a Pcs-tracked row reads back as 1200/Pcs — 12x too high).
 *
 * Fix: _convertVendorRateAcrossItems() (module_items.js) reconciles a
 * vendor rate across differing Purchase Units before merging; module_stock
 * .js#syncStockForItem's 'merge' action reconciles Current/Initial Stock
 * across differing Base Units the same way.
 *
 * Run: node .pw-test/test_item_merge_unit_conversion.js
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
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
  console,
  Logger: { log: () => {} },
  Utilities: { getUuid: () => 'uuid-' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);

['config.js', 'utils.js', 'module_units.js', 'module_items.js', 'module_stock.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});

vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.ITEMS_COL = ITEMS_COL;
  global.STOCK_COL = STOCK_COL;
  global.UNITS_COL = UNITS_COL;
`, ctx, { filename: 'expose.js' });

const {
  APP_CONFIG, UNITS_COL, STOCK_COL, ITEMS_COL,
  _convertVendorRateAcrossItems, _getUnitsMap, syncStockForItem
} = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); }
}

const unitsSheet = ss.addSheet(APP_CONFIG.SHEETS.UNITS);
unitsSheet._set(2, UNITS_COL.UNIT_NAME, 'Pcs'); unitsSheet._set(2, UNITS_COL.FAMILY, 'Count'); unitsSheet._set(2, UNITS_COL.FACTOR_TO_BASE, 1);
unitsSheet._set(3, UNITS_COL.UNIT_NAME, 'Dozen'); unitsSheet._set(3, UNITS_COL.FAMILY, 'Count'); unitsSheet._set(3, UNITS_COL.FACTOR_TO_BASE, 12);

console.log('\n=== Fix: _convertVendorRateAcrossItems reconciles a rate across differing Purchase Units ===');
{
  const unitsMap = _getUnitsMap();
  // Source item quoted at 1200 per Dozen (i.e. 100/Pcs); target item is
  // tracked/purchased in Pcs. Merging must NOT copy 1200 straight across
  // (that would read back as 1200/Pcs -- 12x too high).
  const sourceInfo = { baseUnit: 'Dozen', purchaseUnit: 'Dozen', weightPerBaseUnit: 0 };
  const targetInfo = { baseUnit: 'Pcs', purchaseUnit: 'Pcs', weightPerBaseUnit: 0 };
  const converted = _convertVendorRateAcrossItems(1200, sourceInfo, targetInfo, unitsMap);
  assert(Math.abs(converted - 100) < 0.001, `1200/Dozen converts to 100/Pcs on the target row (got ${converted})`);
}

console.log('\n=== Regression: same Purchase Unit on both sides is a pure no-op ===');
{
  const unitsMap = _getUnitsMap();
  const sourceInfo = { baseUnit: 'Pcs', purchaseUnit: 'Pcs', weightPerBaseUnit: 0 };
  const targetInfo = { baseUnit: 'Pcs', purchaseUnit: 'Pcs', weightPerBaseUnit: 0 };
  const converted = _convertVendorRateAcrossItems(75, sourceInfo, targetInfo, unitsMap);
  assert(converted === 75, `identical units leave the rate untouched (got ${converted})`);
}

console.log('\n=== Fix: syncStockForItem merge reconciles Current/Initial Stock across differing Base Units ===');
{
  const itemsSheet = ss.addSheet(APP_CONFIG.SHEETS.ITEMS);
  itemsSheet._set(2, ITEMS_COL.ITEM_NAME, 'Old Widget Name');
  itemsSheet._set(2, ITEMS_COL.SIZE, '');
  itemsSheet._set(2, ITEMS_COL.BASE_UNIT, 'Dozen');
  itemsSheet._set(2, ITEMS_COL.PURCHASE_UNIT, 'Dozen');
  itemsSheet._set(3, ITEMS_COL.ITEM_NAME, 'Widget');
  itemsSheet._set(3, ITEMS_COL.SIZE, '');
  itemsSheet._set(3, ITEMS_COL.BASE_UNIT, 'Pcs');
  itemsSheet._set(3, ITEMS_COL.PURCHASE_UNIT, 'Pcs');

  const stockSheet = ss.addSheet(APP_CONFIG.SHEETS.STOCK);
  stockSheet._set(2, STOCK_COL.ITEM_NAME, 'Old Widget Name');
  stockSheet._set(2, STOCK_COL.SIZE, '');
  stockSheet._set(2, STOCK_COL.INITIAL_STOCK, 5); // 5 Dozen
  stockSheet._set(2, STOCK_COL.CURRENT_STOCK, 3); // 3 Dozen
  stockSheet._set(3, STOCK_COL.ITEM_NAME, 'Widget');
  stockSheet._set(3, STOCK_COL.SIZE, '');
  stockSheet._set(3, STOCK_COL.INITIAL_STOCK, 50); // 50 Pcs
  stockSheet._set(3, STOCK_COL.CURRENT_STOCK, 20); // 20 Pcs

  syncStockForItem('merge', { oldName: 'Old Widget Name', oldSize: '', newName: 'Widget', newSize: '' });

  const mergedRow = [2, 3].find(r => String(stockSheet._get(r, STOCK_COL.ITEM_NAME)) === 'Widget');
  const initial = Number(stockSheet._get(mergedRow, STOCK_COL.INITIAL_STOCK));
  const current = Number(stockSheet._get(mergedRow, STOCK_COL.CURRENT_STOCK));
  // 5 Dozen -> 60 Pcs, + 50 Pcs already on target = 110. 3 Dozen -> 36 Pcs, + 20 = 56.
  assert(initial === 110, `Initial Stock: (5 Dozen -> 60 Pcs) + 50 Pcs = 110, NOT 5+50=55 (got ${initial})`);
  assert(current === 56, `Current Stock: (3 Dozen -> 36 Pcs) + 20 Pcs = 56, NOT 3+20=23 (got ${current})`);

  const oldRowStillThere = [2, 3].some(r => String(stockSheet._get(r, STOCK_COL.ITEM_NAME)) === 'Old Widget Name');
  assert(!oldRowStillThere, 'old row was deleted after merging');
}

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
