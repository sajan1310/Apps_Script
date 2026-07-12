/**
 * Standalone Node harness (same pattern as test_po_status.js) proving the
 * perf fix: suggestPoAllocations() and _getOpenPoSummary() used to call
 * getPOData() (which internally aggregates the whole Bill Ledger via
 * _aggregateBilledBaseQtyByPo() for status) and THEN separately re-aggregate
 * it themselves — two full Bill-sheet reads per call on interactive paths
 * (bill entry, dashboard load). getPOData() now accepts an optional
 * preloaded billedMap so callers compute it once and share it. This test
 * counts actual getBillData() invocations (what _aggregateBilledBaseQtyByPo
 * calls under the hood) to prove exactly one Bill-sheet read happens per
 * call, not two.
 *
 * Run: node .pw-test/test_po_bill_read_count.js
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
  deleteRows(r, count) { for (let i = 0; i < count; i++) this.rows.splice(r - 1, 1); }
  insertRows(r, count) {
    const blank = [];
    for (let i = 0; i < count; i++) blank.push([]);
    this.rows.splice(r - 1, 0, ...blank);
  }
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
  Utilities: {
    getUuid: () => 'uuid-' + Math.random().toString(36).slice(2),
    formatDate: (date, tz, fmt) => {
      const y = date.getFullYear(), m = String(date.getMonth() + 1).padStart(2, '0'), d = String(date.getDate()).padStart(2, '0');
      return fmt === 'dd/MM/yyyy' ? `${d}/${m}/${y}` : `${y}-${m}-${d}`;
    }
  },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }), getScriptTimeZone: () => 'Asia/Kolkata' }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);

const files = [
  'config.js', 'utils.js', 'module_units.js', 'module_items.js',
  'module_stock.js', 'module_bill.js', 'module_po.js', 'module_vendors.js', 'module_dashboard.js'
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
  APP_CONFIG, ITEMS_COL, PO_COL,
  saveBill, savePO, suggestPoAllocations, _getOpenPoSummary
} = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); }
  else console.log('PASS:', msg);
}

// ── Seed Unit Master, Items Master, and a PO + Bill (same shape as test_po_status.js) ──
const unitsSheet = ss.addSheet(APP_CONFIG.SHEETS.UNITS);
unitsSheet._set(1, 1, 'Unit Name'); unitsSheet._set(1, 2, 'Family'); unitsSheet._set(1, 3, 'Factor'); unitsSheet._set(1, 4, 'Remarks');
unitsSheet._set(2, 1, 'Pcs'); unitsSheet._set(2, 2, 'Count'); unitsSheet._set(2, 3, 1); unitsSheet._set(2, 4, '');

const itemsSheet = ss.addSheet(APP_CONFIG.SHEETS.ITEMS);
itemsSheet._set(1, 1, 'Item Name');
itemsSheet._set(2, ITEMS_COL.ITEM_NAME, 'Widget');
itemsSheet._set(2, ITEMS_COL.SIZE, '');
itemsSheet._set(2, ITEMS_COL.BASE_UNIT, 'Pcs');
itemsSheet._set(2, ITEMS_COL.PURCHASE_UNIT, 'Pcs');

ss.addSheet(APP_CONFIG.SHEETS.VENDORS);
const poSheet = ss.addSheet(APP_CONFIG.SHEETS.PO);
poSheet._set(2, PO_COL.PO_NUMBER, 'PO Number');
const billSheet = ss.addSheet(APP_CONFIG.SHEETS.BILL);
billSheet._set(1, 1, 'PO Number');

const poRes = savePO({
  poNumber: '', poDate: '2026-01-01', vendor: 'Acme Vendor', contact: '',
  poDescription: '', poRemarks: '', supplierRemarks: '',
  items: JSON.stringify([{ name: 'Widget', size: '', narration: '', qty: 10, unit: 'Pcs', price: 50 }])
});
assert(poRes.success, 'savePO succeeds: ' + poRes.message);
const poNumber = poRes.data.poNumber;

const billRes = saveBill({
  billNumber: 'BILL-1', billDate: '2026-01-05', vendor: 'Acme Vendor', contact: '', remarks: '',
  poNumbers: JSON.stringify([poNumber]),
  items: JSON.stringify([{ name: 'Widget', size: '', narration: '', qty: 4, unit: 'Pcs', price: 50, po: poNumber }])
});
assert(billRes.success, 'saveBill succeeds: ' + billRes.message);

// ── Instrument getBillData() (what _aggregateBilledBaseQtyByPo calls under
// the hood) to count actual Bill-sheet reads per call. ──
let getBillDataCalls = 0;
const originalGetBillData = ctx.getBillData;
ctx.getBillData = function (...args) {
  getBillDataCalls++;
  return originalGetBillData.apply(this, args);
};

console.log('\n=== Test: suggestPoAllocations reads the Bill sheet exactly once ===');
getBillDataCalls = 0;
const suggestRes = suggestPoAllocations('Acme Vendor', [
  { rowIndex: 0, name: 'Widget', size: '', narration: '', qty: 5, unit: 'Pcs', price: 50 }
], '2026-01-06');
assert(suggestRes.success, 'suggestPoAllocations succeeds: ' + suggestRes.message);
assert(getBillDataCalls === 1, `getBillData() called exactly once (got ${getBillDataCalls})`);
assert(Array.isArray(suggestRes.data) && suggestRes.data[0] && suggestRes.data[0].allocations.length > 0,
  'still correctly suggests an allocation against the open PO line (behavior preserved)');
assert(suggestRes.data[0].allocations[0].qty === 5 || suggestRes.data[0].unmatchedQty !== undefined,
  'allocation result shape is intact');

console.log('\n=== Test: _getOpenPoSummary reads the Bill sheet exactly once ===');
getBillDataCalls = 0;
const summary = _getOpenPoSummary();
assert(getBillDataCalls === 1, `getBillData() called exactly once (got ${getBillDataCalls})`);
assert(summary.count === 1, `open PO summary still counts the 1 open PO (got ${summary.count})`);
assert(Math.abs(summary.value - (6 * 50)) < 0.01, `open PO summary value reflects the 6 remaining units (got ${summary.value})`);

console.log(failures === 0 ? '\nALL TESTS PASSED' : `\n${failures} TEST(S) FAILED`);
process.exitCode = failures === 0 ? 0 : 1;
