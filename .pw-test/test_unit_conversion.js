/**
 * Standalone Node harness (same pattern as test_merge_and_backfill.js) that
 * loads the REAL server-side files and exercises the Unit Master / Base Unit
 * conversion pipeline end-to-end:
 *   - convertQtyToBaseUnit / convertRateToBaseUnit math (Gross -> Pcs)
 *   - getItemsData() exposing ratePerBaseUnit
 *   - saveBill() / savePO() computing baseQty/baseRate on the saved row
 *   - recalculateStock() using BASE_QTY (not the as-entered QTY) for stock math
 *   - autoExtractFromPoOrBill() feeding baseRate (not the as-entered price)
 *     into Items Master vendor rates
 *
 * Run: node .pw-test/test_unit_conversion.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet; this.row = row; this.col = col;
    this.numRows = numRows; this.numCols = numCols;
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
  insertColumnsBefore(col, count) {
    this.rows.forEach(row => { for (let i = 0; i < count; i++) row.splice(col - 1, 0, ''); });
  }
  insertColumnBefore(col) { this.insertColumnsBefore(col, 1); }
}

class FakeSpreadsheet {
  constructor() { this.sheets = {}; }
  getSheetByName(name) { return this.sheets[name] || null; }
  addSheet(name) { const s = new FakeSheet(name); this.sheets[name] = s; return s; }
  getSpreadsheetTimeZone() { return 'Asia/Kolkata'; }
}

const ss = new FakeSpreadsheet();

const fakeCache = {
  _store: {},
  get(k) { return this._store[k]; },
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
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }), getScriptTimeZone: () => 'Asia/Kolkata' }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);

const files = [
  'config.js', 'utils.js', 'module_units.js', 'module_items.js',
  'module_stock.js', 'module_bill.js', 'module_po.js', 'module_vendors.js'
];
files.forEach(f => vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f }));

vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.ITEMS_COL = ITEMS_COL;
  global.STOCK_COL = STOCK_COL;
  global.BILL_COL = BILL_COL;
  global.PO_COL = PO_COL;
  global.UNITS_COL = UNITS_COL;
`, ctx, { filename: 'expose.js' });

const {
  APP_CONFIG, ITEMS_COL, STOCK_COL, BILL_COL, PO_COL,
  convertQtyToBaseUnit, convertRateToBaseUnit,
  getItemsData, saveItem, saveBill, savePO,
  _getBilledAndConsumedQtyMaps, autoExtractFromPoOrBill
} = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); }
  else console.log('PASS:', msg);
}
function approx(a, b, eps = 0.001) { return Math.abs(a - b) < eps; }

// ── Seed Unit Master ────────────────────────────────────────────────────
const unitsSheet = ss.addSheet(APP_CONFIG.SHEETS.UNITS);
unitsSheet._set(1, 1, 'Unit Name'); unitsSheet._set(1, 2, 'Family'); unitsSheet._set(1, 3, 'Factor'); unitsSheet._set(1, 4, 'Remarks');
[['Pcs', 'Count', 1, ''], ['Gross', 'Count', 144, ''], ['Kg', 'Weight', 1000, ''], ['Gram', 'Weight', 1, '']]
  .forEach((row, i) => row.forEach((v, c) => unitsSheet._set(i + 2, c + 1, v)));

// ── Seed Items Master: a spoke bought by Gross @ 99, sold/used by Pcs ───
const itemsSheet = ss.addSheet(APP_CONFIG.SHEETS.ITEMS);
itemsSheet._set(1, 1, 'Item Name');
itemsSheet._set(2, ITEMS_COL.ITEM_NAME, 'Spoke');
itemsSheet._set(2, ITEMS_COL.SIZE, '110mm');
itemsSheet._set(2, ITEMS_COL.BASE_UNIT, 'Pcs');
itemsSheet._set(2, ITEMS_COL.PURCHASE_UNIT, 'Gross');
itemsSheet._set(2, ITEMS_COL.VENDOR_DATA, 'Avon Cycles');
itemsSheet._set(2, ITEMS_COL.VENDOR_DATA + 1, 99);

// === Test 1: conversion helper math ===
assert(approx(convertQtyToBaseUnit(2, 'Gross', { baseUnit: 'Pcs' }), 288), 'convertQtyToBaseUnit: 2 Gross -> 288 Pcs');
assert(approx(convertRateToBaseUnit(99, 'Gross', { baseUnit: 'Pcs' }), 99 / 144), 'convertRateToBaseUnit: 99/Gross -> 0.6875/Pcs');
assert(approx(convertQtyToBaseUnit(500, 'Gram', { baseUnit: 'Pcs', weightPerBaseUnit: 25 }), 20), '500 Gram / 25 g-per-Pc -> 20 Pcs');

// === Test 2: getItemsData exposes ratePerBaseUnit ===
const itemsRes = getItemsData();
const spoke = itemsRes.data.find(i => i.name === 'Spoke');
assert(!!spoke, 'getItemsData returns the seeded Spoke item');
assert(spoke.purchaseUnit === 'Gross', 'Spoke purchaseUnit read back as Gross');
assert(approx(spoke.vendors[0].rate, 99), 'Spoke vendor rate stays as-quoted (99)');
assert(approx(spoke.vendors[0].ratePerBaseUnit, 99 / 144), 'Spoke vendor ratePerBaseUnit computed as 99/144');

// === Test 3: saveBill computes baseQty/baseRate ===
const billSheetSeed = ss.addSheet(APP_CONFIG.SHEETS.BILL);
billSheetSeed._set(1, 1, 'PO Number'); // minimal header so _ensureBillSheetColumns has something to inspect

const billRes = saveBill({
  billNumber: 'B-100',
  billDate: '2026-01-01',
  vendor: 'Avon Cycles',
  contact: '',
  remarks: '',
  poNumbers: '["DIRECT"]',
  items: JSON.stringify([{ name: 'Spoke', size: '110mm', narration: '', qty: 2, unit: 'Gross', price: 99, gst: 0 }])
});
assert(billRes.success, 'saveBill succeeds for a Gross-unit line: ' + billRes.message);

const billSheet = ss.getSheetByName(APP_CONFIG.SHEETS.BILL);
const savedBaseQty = Number(billSheet.getRange(2, BILL_COL.BASE_QTY).getValue());
const savedBaseRate = Number(billSheet.getRange(2, BILL_COL.BASE_RATE).getValue());
assert(approx(savedBaseQty, 288), 'Saved Bill row BASE_QTY = 288 (got ' + savedBaseQty + ')');
assert(approx(savedBaseRate, 99 / 144), 'Saved Bill row BASE_RATE = 99/144 (got ' + savedBaseRate + ')');

// === Test 4: recalculateStock uses BASE_QTY, not the as-entered QTY ===
const stockSheet = ss.addSheet(APP_CONFIG.SHEETS.STOCK);
stockSheet._set(1, 1, 'Item Name');
stockSheet._set(2, STOCK_COL.ITEM_NAME, 'Spoke');
stockSheet._set(2, STOCK_COL.SIZE, '110mm');
stockSheet._set(2, STOCK_COL.INITIAL_STOCK, 10);
stockSheet._set(2, STOCK_COL.CURRENT_STOCK, 10);
stockSheet._set(2, STOCK_COL.THRESHOLD, 0);

const { billQtyMap } = _getBilledAndConsumedQtyMaps(ss);
const stockDelta = billQtyMap['spoke|110mm'] || 0;
assert(approx(stockDelta, 288), 'Stock billed-qty map credits 288 Pcs for a 2-Gross bill line (got ' + stockDelta + ')');

// === Test 5: savePO also computes baseQty/baseRate ===
ss.addSheet(APP_CONFIG.SHEETS.VENDORS);
const poSheet = ss.addSheet(APP_CONFIG.SHEETS.PO);
// PO_SETTINGS.DATA_START_ROW = 3 (2 header rows) — seed a blank header row 2
// so the real sheet's row layout (data starts at row 3) is reproduced here.
poSheet._set(2, PO_COL.PO_NUMBER, 'PO Number');
const poRes = savePO({
  poNumber: '',
  poDate: '2026-01-01',
  vendor: 'Avon Cycles',
  contact: '',
  poDescription: '', poRemarks: '', supplierRemarks: '',
  items: JSON.stringify([{ name: 'Spoke', size: '110mm', narration: '', qty: 3, unit: 'Gross', price: 99 }])
});
assert(poRes.success, 'savePO succeeds for a Gross-unit line: ' + poRes.message);
const poBaseQty = Number(poSheet.getRange(3, PO_COL.BASE_QTY).getValue());
assert(approx(poBaseQty, 432), 'Saved PO row BASE_QTY = 432 (3 Gross, got ' + poBaseQty + ')');

// === Test 6: autoExtractFromPoOrBill writes baseRate into Items Master, not raw price ===
autoExtractFromPoOrBill('Avon Cycles', '', [{ name: 'Spoke', size: '110mm', narration: '', price: 99, baseRate: 99 / 144 }]);
const refreshed = getItemsData().data.find(i => i.name === 'Spoke');
const avonRate = refreshed.vendors.find(v => v.vendor === 'Avon Cycles').rate;
assert(approx(avonRate, 99 / 144), 'autoExtractFromPoOrBill stored baseRate (0.6875), not raw price 99 (got ' + avonRate + ')');

console.log(failures === 0 ? '\nALL TESTS PASSED' : `\n${failures} TEST(S) FAILED`);
process.exitCode = failures === 0 ? 0 : 1;
