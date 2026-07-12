/**
 * Standalone Node harness (same pattern as test_unit_conversion.js) that
 * loads the REAL server-side files and exercises the new PO status
 * computation end-to-end: savePO() + saveBill() write real sheet rows,
 * then getPOData() -> _attachPoStatus() must classify each PO as
 * 'PO Issued' / 'Partially Received' / 'Completed' purely from what's
 * actually been billed, with correct per-line receivedQty/pendingQty
 * (converted back to the PO line's own entered unit).
 *
 * Run: node .pw-test/test_po_status.js
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
  APP_CONFIG, ITEMS_COL, PO_COL,
  saveBill, savePO, getPOData
} = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); }
  else console.log('PASS:', msg);
}
function approx(a, b, eps = 0.001) { return Math.abs(a - b) < eps; }

// ── Seed Unit Master: Pcs (base) + Gross = 144 Pcs ──────────────────────
const unitsSheet = ss.addSheet(APP_CONFIG.SHEETS.UNITS);
unitsSheet._set(1, 1, 'Unit Name'); unitsSheet._set(1, 2, 'Family'); unitsSheet._set(1, 3, 'Factor'); unitsSheet._set(1, 4, 'Remarks');
[['Pcs', 'Count', 1, ''], ['Gross', 'Count', 144, '']]
  .forEach((row, i) => row.forEach((v, c) => unitsSheet._set(i + 2, c + 1, v)));

// ── Seed Items Master: Widget (Pcs base) + Spoke (Gross-purchased, Pcs base) ─
const itemsSheet = ss.addSheet(APP_CONFIG.SHEETS.ITEMS);
itemsSheet._set(1, 1, 'Item Name');
itemsSheet._set(2, ITEMS_COL.ITEM_NAME, 'Widget');
itemsSheet._set(2, ITEMS_COL.SIZE, '');
itemsSheet._set(2, ITEMS_COL.BASE_UNIT, 'Pcs');
itemsSheet._set(2, ITEMS_COL.PURCHASE_UNIT, 'Pcs');
itemsSheet._set(3, ITEMS_COL.ITEM_NAME, 'Spoke');
itemsSheet._set(3, ITEMS_COL.SIZE, '110mm');
itemsSheet._set(3, ITEMS_COL.BASE_UNIT, 'Pcs');
itemsSheet._set(3, ITEMS_COL.PURCHASE_UNIT, 'Gross');

ss.addSheet(APP_CONFIG.SHEETS.VENDORS);
const poSheet = ss.addSheet(APP_CONFIG.SHEETS.PO);
poSheet._set(2, PO_COL.PO_NUMBER, 'PO Number'); // header row so DATA_START_ROW=3 matches
const billSheet = ss.addSheet(APP_CONFIG.SHEETS.BILL);
billSheet._set(1, 1, 'PO Number');

function makePO(itemLines) {
  const res = savePO({
    poNumber: '', poDate: '2026-01-01', vendor: 'Acme Vendor', contact: '',
    poDescription: '', poRemarks: '', supplierRemarks: '',
    items: JSON.stringify(itemLines)
  });
  assert(res.success, 'savePO succeeds: ' + res.message);
  return res.data.poNumber;
}

function makeBill(poNumber, itemLines) {
  const res = saveBill({
    billNumber: 'BILL-' + poNumber + '-' + Math.random().toString(36).slice(2, 6),
    billDate: '2026-01-05', vendor: 'Acme Vendor', contact: '', remarks: '',
    poNumbers: JSON.stringify([poNumber]),
    items: JSON.stringify(itemLines.map(l => ({ ...l, po: poNumber })))
  });
  assert(res.success, 'saveBill succeeds: ' + res.message);
  return res;
}

// PO-1: issued, never billed -> 'PO Issued'
const po1 = makePO([{ name: 'Widget', size: '', narration: '', qty: 10, unit: 'Pcs', price: 50 }]);

// PO-2: billed 4 of 10 -> 'Partially Received', pendingQty 6
const po2 = makePO([{ name: 'Widget', size: '', narration: '', qty: 10, unit: 'Pcs', price: 50 }]);
makeBill(po2, [{ name: 'Widget', size: '', narration: '', qty: 4, unit: 'Pcs', price: 50 }]);

// PO-3: billed 10 of 10 -> 'Completed', pendingQty 0
const po3 = makePO([{ name: 'Widget', size: '', narration: '', qty: 10, unit: 'Pcs', price: 50 }]);
makeBill(po3, [{ name: 'Widget', size: '', narration: '', qty: 10, unit: 'Pcs', price: 50 }]);

// PO-4: two lines, one fully billed + one untouched -> must stay
// 'Partially Received' overall, NOT 'Completed' (per-line check, not summed).
const po4 = makePO([
  { name: 'Widget', size: '', narration: '', qty: 10, unit: 'Pcs', price: 50 },
  { name: 'Spoke', size: '110mm', narration: '', qty: 1, unit: 'Gross', price: 99 }
]);
makeBill(po4, [{ name: 'Widget', size: '', narration: '', qty: 10, unit: 'Pcs', price: 50 }]);

// PO-5: ordered 2 Gross of Spoke (=288 Pcs), billed 144 Pcs directly in Pcs
// -> receivedQty must convert back to 1 Gross (not 144), pendingQty 1 Gross.
const po5 = makePO([{ name: 'Spoke', size: '110mm', narration: '', qty: 2, unit: 'Gross', price: 99 }]);
makeBill(po5, [{ name: 'Spoke', size: '110mm', narration: '', qty: 144, unit: 'Pcs', price: 99 / 144 }]);

const data = getPOData().data;
const byNum = Object.fromEntries(data.map(po => [po.poNumber, po]));

assert(byNum[po1].status === 'PO Issued', `PO-1 (never billed) status = 'PO Issued' (got ${byNum[po1].status})`);
assert(approx(byNum[po1].items[0].receivedQty, 0), 'PO-1 receivedQty = 0');
assert(approx(byNum[po1].items[0].pendingQty, 10), 'PO-1 pendingQty = 10');

assert(byNum[po2].status === 'Partially Received', `PO-2 (4 of 10 billed) status = 'Partially Received' (got ${byNum[po2].status})`);
assert(approx(byNum[po2].items[0].receivedQty, 4), 'PO-2 receivedQty = 4');
assert(approx(byNum[po2].items[0].pendingQty, 6), 'PO-2 pendingQty = 6');

assert(byNum[po3].status === 'Completed', `PO-3 (10 of 10 billed) status = 'Completed' (got ${byNum[po3].status})`);
assert(approx(byNum[po3].items[0].pendingQty, 0), 'PO-3 pendingQty = 0');

assert(byNum[po4].status === 'Partially Received',
  `PO-4 (1 of 2 lines fully billed) status = 'Partially Received', not 'Completed' (got ${byNum[po4].status})`);
const po4Widget = byNum[po4].items.find(i => i.name === 'Widget');
const po4Spoke = byNum[po4].items.find(i => i.name === 'Spoke');
assert(approx(po4Widget.pendingQty, 0), 'PO-4 Widget line fully received (pendingQty 0)');
assert(approx(po4Spoke.pendingQty, 1), 'PO-4 Spoke line untouched (pendingQty 1 Gross)');

assert(byNum[po5].status === 'Partially Received', `PO-5 (cross-unit partial) status = 'Partially Received' (got ${byNum[po5].status})`);
assert(approx(byNum[po5].items[0].receivedQty, 1), `PO-5 receivedQty converts 144 Pcs back to 1 Gross (got ${byNum[po5].items[0].receivedQty})`);
assert(approx(byNum[po5].items[0].pendingQty, 1), `PO-5 pendingQty = 1 Gross remaining (got ${byNum[po5].items[0].pendingQty})`);

console.log(failures === 0 ? '\nALL TESTS PASSED' : `\n${failures} TEST(S) FAILED`);
process.exitCode = failures === 0 ? 0 : 1;
