/**
 * Standalone Node harness (same mock-SpreadsheetApp pattern as
 * test_regression_sweep_13_modules.js) covering the new Labor Job Bill
 * feature: saveBill(formData.billType === 'LABOR') resolves a chosen
 * Process against Process Master, forces ITEM_NAME to that process's
 * Output Item Name, and records PROCESS_NAME + COLOR on the row —
 * without touching Vendor/Item Master auto-extraction (VENDOR here is a
 * Contractor's name, not a real vendor).
 *
 * Run: node .pw-test/test_bill_labor_job.js
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
  getDataRange() {
    const lastRow = Math.max(this.getLastRow(), 1);
    const lastCol = Math.max(this.getLastColumn(), 1);
    return new FakeRange(this, 1, 1, lastRow, lastCol);
  }
  appendRow(arr) {
    const r = this.getLastRow() + 1;
    arr.forEach((v, i) => this._set(r, i + 1, v));
  }
  deleteRow(r) { this.rows.splice(r - 1, 1); }
  deleteRows(r, n) { this.rows.splice(r - 1, n); }
  insertRows(r, n) { for (let i = 0; i < n; i++) this.rows.splice(r - 1, 0, []); }
  insertColumnsAfter(afterCol, count) {
    this.rows.forEach(row => { for (let i = 0; i < count; i++) row.splice(afterCol, 0, ''); });
  }
  insertColumnsBefore(beforeCol, count) {
    this.rows.forEach(row => { for (let i = 0; i < count; i++) row.splice(beforeCol - 1, 0, ''); });
  }
  insertColumnAfter(afterCol) { this.insertColumnsAfter(afterCol, 1); }
  insertColumnBefore(beforeCol) { this.insertColumnsBefore(beforeCol, 1); }
}

class FakeSpreadsheet {
  constructor() { this.sheets = {}; }
  getSheetByName(name) { return this.sheets[name] || null; }
  insertSheet(name) { return this.addSheet(name); }
  addSheet(name) { const s = new FakeSheet(name); this.sheets[name] = s; return s; }
  getSpreadsheetTimeZone() { return 'Asia/Kolkata'; }
}

const ss = new FakeSpreadsheet();

const fakeCache = {
  _store: {},
  get(k) { return this._store[k]; },
  put(k, v) { this._store[k] = v; },
  remove(k) { delete this._store[k]; },
  removeAll() { this._store = {}; }
};

function pad(n, w) { return String(n).padStart(w, '0'); }
function fakeFormatDate(date, tz, pattern) {
  const d = date instanceof Date ? date : new Date(date);
  const map = {
    yyyy: d.getFullYear(), MM: pad(d.getMonth() + 1, 2), dd: pad(d.getDate(), 2),
    HH: pad(d.getHours(), 2), mm: pad(d.getMinutes(), 2), ss: pad(d.getSeconds(), 2)
  };
  return pattern.replace(/yyyy|MM|dd|HH|mm|ss/g, m => map[m]);
}

const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush: () => {} },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  CacheService: { getScriptCache: () => fakeCache },
  console,
  Logger: { log: () => {} },
  Utilities: {
    getUuid: () => 'uuid-' + Math.random().toString(36).slice(2),
    formatDate: fakeFormatDate
  },
  Session: {
    getActiveUser: () => ({ getEmail: () => 'test@example.com' }),
    getScriptTimeZone: () => 'Asia/Kolkata'
  },
  ScriptApp: { newTrigger: () => ({ timeBased: () => ({ everyHours: () => ({ create: () => {} }) }) }), getProjectTriggers: () => [] }
};
sandbox.global = sandbox;

const ctx = vm.createContext(sandbox);

const files = [
  'config.js', 'utils.js',
  'module_units.js', 'module_vendors.js', 'module_items.js', 'module_stock.js',
  'module_clients.js', 'module_contractors.js',
  'module_po.js', 'module_bill.js', 'module_return.js', 'module_wastage.js', 'module_issue.js',
  'module_bom.js', 'module_process.js', 'module_warehouse.js', 'module_production.js',
  'module_dispatch.js'
];

files.forEach(f => {
  const code = fs.readFileSync(path.join(ROOT, f), 'utf8');
  vm.runInContext(code, ctx, { filename: f });
});

vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.BILL_COL = BILL_COL; global.ITEMS_COL = ITEMS_COL; global.PROCESS_COL = PROCESS_COL;
`, ctx, { filename: 'expose.js' });

const {
  APP_CONFIG, BILL_COL, saveBill, getBillData, deleteBill, saveProcess
} = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); }
  else console.log('PASS:', msg);
}

ss.addSheet(APP_CONFIG.SHEETS.VENDORS);
ss.addSheet(APP_CONFIG.SHEETS.ITEMS);
ss.addSheet(APP_CONFIG.SHEETS.UNITS);
ss.addSheet(APP_CONFIG.SHEETS.PO);
const billSheet0 = ss.addSheet(APP_CONFIG.SHEETS.BILL);
// Occupies row 1 so getLastRow() starts at 1 and the first saved bill lands
// on DATA_START_ROW (2), not row 1 — mirrors every other Bill test harness.
billSheet0._set(1, 1, 'PO Number');

console.log('\n=== Setup: create an active Process with a Color-carrying Output Item ===');
// initProcessMasterSheet() auto-seeds demo rows (incl. an Output Item Name
// "Painted Frame") the first time the sheet is touched — use distinct
// "Sweep"-prefixed names, same convention as test_regression_sweep, to
// avoid colliding with them.
const procRes = saveProcess({
  processName: 'Sweep Painting', lotPrefix: 'SWP', outputItemName: 'Sweep Painted Frame', sequence: 1
});
assert(procRes.success, 'saveProcess succeeds: ' + procRes.message);

console.log('\n=== Test: saveBill rejects an unrecognized Process on a Labor Job line ===');
{
  const res = saveBill({
    billNumber: 'LB-BAD', billDate: '2026-01-10', vendor: 'Rajesh Contractor', contact: '', remarks: '',
    billType: 'LABOR',
    items: JSON.stringify([{ processName: 'Not A Real Process', color: 'Red', qty: 5, unit: 'Pcs', price: 20, gst: 18 }])
  });
  assert(!res.success, 'save is rejected for an unrecognized process');
  assert(/not a recognized active Process/.test(res.message), `error message names the problem (got "${res.message}")`);
}

console.log('\n=== Test: saveBill records a Labor Job bill line correctly ===');
{
  const res = saveBill({
    billNumber: 'LB-1', billDate: '2026-01-10', vendor: 'Rajesh Contractor', contact: '', remarks: 'Job work',
    billType: 'LABOR',
    items: JSON.stringify([{ processName: 'Sweep Painting', color: 'Red', qty: 5, unit: 'Pcs', price: 20, gst: 18, narration: 'Batch 1' }])
  });
  assert(res.success, 'saveBill succeeds: ' + res.message);

  const billSheet = ss.getSheetByName(APP_CONFIG.SHEETS.BILL);
  const lastRow = billSheet.getLastRow();
  assert(billSheet.getRange(lastRow, BILL_COL.ITEM_NAME).getValue() === 'Sweep Painted Frame', 'ITEM_NAME forced to the process Output Item Name (got "' + billSheet.getRange(lastRow, BILL_COL.ITEM_NAME).getValue() + '")');
  assert(billSheet.getRange(lastRow, BILL_COL.BILL_TYPE).getValue() === 'LABOR', 'BILL_TYPE column stored LABOR');
  assert(billSheet.getRange(lastRow, BILL_COL.PROCESS_NAME).getValue() === 'Sweep Painting', 'PROCESS_NAME column stored Sweep Painting');
  assert(billSheet.getRange(lastRow, BILL_COL.COLOR).getValue() === 'Red', 'COLOR column stored Red');
  assert(billSheet.getRange(lastRow, BILL_COL.VENDOR).getValue() === 'Rajesh Contractor', 'VENDOR column reused for the contractor name');

  // Vendor/Item Master auto-extraction must be SKIPPED for Labor bills —
  // Rajesh Contractor is not a real vendor and shouldn't pollute Vendor Master.
  const vendorSheet = ss.getSheetByName(APP_CONFIG.SHEETS.VENDORS);
  assert(vendorSheet.getLastRow() === 0, 'Vendor Master was NOT auto-populated with the contractor name (got ' + vendorSheet.getLastRow() + ' rows)');

  const billData = getBillData();
  assert(billData.success, 'getBillData succeeds');
  const bill = billData.data.find(b => b.billNumber === 'LB-1');
  assert(!!bill, 'saved Labor Job bill is present in getBillData()');
  assert(bill && bill.billType === 'LABOR', 'bill header reports billType LABOR (got ' + (bill && bill.billType) + ')');
  const item = bill && bill.items[0];
  assert(item && item.name === 'Sweep Painted Frame', 'item.name is the process Output Item Name (got ' + (item && item.name) + ')');
  assert(item && item.processName === 'Sweep Painting', 'item.processName round-trips (got ' + (item && item.processName) + ')');
  assert(item && item.color === 'Red', 'item.color round-trips (got ' + (item && item.color) + ')');
  assert(item && item.qty === 5, 'item.qty round-trips (got ' + (item && item.qty) + ')');
}

console.log('\n=== Test: a normal Goods bill still round-trips with billType GOODS and blank Process/Color ===');
{
  const res = saveBill({
    billNumber: 'GB-1', billDate: '2026-01-11', vendor: 'Acme Vendor', contact: '', remarks: '',
    items: JSON.stringify([{ name: 'Widget', size: '', narration: '', qty: 3, unit: 'Pcs', price: 10, gst: 18 }])
  });
  assert(res.success, 'saveBill (Goods) succeeds: ' + res.message);

  const billData = getBillData();
  const bill = billData.data.find(b => b.billNumber === 'GB-1');
  assert(!!bill, 'Goods bill present');
  assert(bill && bill.billType === 'GOODS', 'Goods bill billType defaults to GOODS (got ' + (bill && bill.billType) + ')');
  const item = bill && bill.items[0];
  assert(item && item.processName === '', 'Goods bill item has blank processName');
  assert(item && item.color === '', 'Goods bill item has blank color');

  // Vendor Master SHOULD gain a row for a real Goods bill.
  const vendorSheet = ss.getSheetByName(APP_CONFIG.SHEETS.VENDORS);
  assert(vendorSheet.getLastRow() === 1, 'Vendor Master WAS auto-populated for a real Goods bill (got ' + vendorSheet.getLastRow() + ' rows)');
}

console.log('\n=== Test: editing a Labor Job bill updates its Process/Color in place ===');
{
  const res = saveBill({
    existingBillNumber: 'LB-1', existingVendor: 'Rajesh Contractor',
    billNumber: 'LB-1', billDate: '2026-01-10', vendor: 'Rajesh Contractor', contact: '', remarks: 'Job work (edited)',
    billType: 'LABOR',
    items: JSON.stringify([{ processName: 'Sweep Painting', color: 'Blue', qty: 7, unit: 'Pcs', price: 20, gst: 18, narration: 'Batch 1' }])
  });
  assert(res.success, 'edit saveBill succeeds: ' + res.message);

  const billData = getBillData();
  const bill = billData.data.find(b => b.billNumber === 'LB-1');
  const item = bill && bill.items[0];
  assert(item && item.color === 'Blue', 'edited color round-trips (got ' + (item && item.color) + ')');
  assert(item && item.qty === 7, 'edited qty round-trips (got ' + (item && item.qty) + ')');
}

console.log(failures === 0 ? '\nALL TESTS PASSED' : `\n${failures} TEST(S) FAILED`);
process.exitCode = failures === 0 ? 0 : 1;
