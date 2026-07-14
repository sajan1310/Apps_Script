/**
 * Standalone Node harness (same pattern as test_po_bill_read_count.js) for
 * the narration-preference fix in suggestPoAllocations() (module_po.js).
 *
 * Before the fix, candidate PO lines were grouped/matched by name+size only
 * — narration (passed in from the bill row) was accepted but never used.
 * When a vendor's open POs carry two lines for the same name+size at the
 * SAME price but different narration (a real case: narration is how two
 * specs/variants of one item+size get told apart, e.g. two colors), the
 * algorithm couldn't tell them apart and fell back to oldest-PO-first,
 * silently drawing down whichever line sorted first regardless of which one
 * the bill actually matched. This test proves a bill row with narration text
 * now prefers the PO line with matching narration over an older PO whose
 * line has different narration.
 *
 * Run: node .pw-test/test_po_narration_disambiguation.js
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

const { APP_CONFIG, ITEMS_COL, PO_COL, savePO, suggestPoAllocations } = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); }
  else console.log('PASS:', msg);
}

// ── Seed Unit Master + Items Master ──
const unitsSheet = ss.addSheet(APP_CONFIG.SHEETS.UNITS);
unitsSheet._set(1, 1, 'Unit Name'); unitsSheet._set(1, 2, 'Family'); unitsSheet._set(1, 3, 'Factor'); unitsSheet._set(1, 4, 'Remarks');
unitsSheet._set(2, 1, 'Pcs'); unitsSheet._set(2, 2, 'Count'); unitsSheet._set(2, 3, 1); unitsSheet._set(2, 4, '');

const itemsSheet = ss.addSheet(APP_CONFIG.SHEETS.ITEMS);
itemsSheet._set(1, 1, 'Item Name');
itemsSheet._set(2, ITEMS_COL.ITEM_NAME, 'Frame');
itemsSheet._set(2, ITEMS_COL.SIZE, '');
itemsSheet._set(2, ITEMS_COL.BASE_UNIT, 'Pcs');
itemsSheet._set(2, ITEMS_COL.PURCHASE_UNIT, 'Pcs');

ss.addSheet(APP_CONFIG.SHEETS.VENDORS);
const poSheet = ss.addSheet(APP_CONFIG.SHEETS.PO);
poSheet._set(2, PO_COL.PO_NUMBER, 'PO Number');
ss.addSheet(APP_CONFIG.SHEETS.BILL);

// Two POs, same vendor, same item name+size+price ('Frame', '26 inch', ₹100),
// but different narration — an older PO for "Red" and a newer PO for "Blue".
// Oldest-PO-first (the old fallback order) would incorrectly pick the Red/
// older PO for a bill line that says "Blue".
const poOldRes = savePO({
  poNumber: '', poDate: '2026-01-01', vendor: 'Acme Vendor', contact: '',
  poDescription: '', poRemarks: '', supplierRemarks: '',
  items: JSON.stringify([{ name: 'Frame', size: '26 inch', narration: 'Red', qty: 10, unit: 'Pcs', price: 100 }])
});
assert(poOldRes.success, 'older PO (Red) saved: ' + poOldRes.message);

const poNewRes = savePO({
  poNumber: '', poDate: '2026-02-01', vendor: 'Acme Vendor', contact: '',
  poDescription: '', poRemarks: '', supplierRemarks: '',
  items: JSON.stringify([{ name: 'Frame', size: '26 inch', narration: 'Blue', qty: 10, unit: 'Pcs', price: 100 }])
});
assert(poNewRes.success, 'newer PO (Blue) saved: ' + poNewRes.message);

console.log('\n=== Test: bill row narration disambiguates between same name+size+price PO lines ===');
const res = suggestPoAllocations('Acme Vendor', [
  { rowIndex: 0, name: 'Frame', size: '26 inch', narration: 'Blue', qty: 4, unit: 'Pcs', price: 100 }
], '2026-03-01');
assert(res.success, 'suggestPoAllocations succeeds: ' + res.message);
const allocs = res.data[0].allocations;
assert(allocs.length === 1, `allocates against exactly one PO line (got ${allocs.length})`);
assert(allocs[0].poNumber === poNewRes.data.poNumber,
  `allocates to the narration-matching (Blue/newer) PO ${poNewRes.data.poNumber}, not the older Red one (got ${allocs[0].poNumber})`);

console.log('\n=== Test: no narration on the bill row still falls back to oldest-PO-first (unchanged behavior) ===');
const resNoNarration = suggestPoAllocations('Acme Vendor', [
  { rowIndex: 0, name: 'Frame', size: '26 inch', narration: '', qty: 4, unit: 'Pcs', price: 100 }
], '2026-03-01');
const allocsNoNarration = resNoNarration.data[0].allocations;
assert(allocsNoNarration.length === 1, `still allocates against exactly one PO line (got ${allocsNoNarration.length})`);
assert(allocsNoNarration[0].poNumber === poOldRes.data.poNumber,
  `falls back to oldest PO ${poOldRes.data.poNumber} when the bill row gives no narration (got ${allocsNoNarration[0].poNumber})`);

console.log('\n=== Test: narration that matches nothing open still matches on name+size (never forced to DIRECT) ===');
const resStrayNarration = suggestPoAllocations('Acme Vendor', [
  { rowIndex: 0, name: 'Frame', size: '26 inch', narration: 'Green (not on any PO)', qty: 4, unit: 'Pcs', price: 100 }
], '2026-03-01');
const allocsStray = resStrayNarration.data[0].allocations;
assert(allocsStray.length === 1, `still finds a match despite the unmatched narration (got ${allocsStray.length} allocations)`);

console.log(failures === 0 ? '\nALL TESTS PASSED' : `\n${failures} TEST(S) FAILED`);
process.exitCode = failures === 0 ? 0 : 1;
