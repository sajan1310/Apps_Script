/**
 * Regression harness for the PO auto-match date-boundary bug in
 * suggestPoAllocations (module_po.js). The eligibility filter that drops POs
 * dated AFTER the bill compared `new Date(po.poDateRaw)` against a
 * `toSafeDateObject(billDate)`. Those two parse the SAME calendar date into
 * DIFFERENT instants: a bare `new Date("2026-05-23")` reads the ISO date-only
 * string as UTC midnight, while toSafeDateObject builds LOCAL midnight. Under
 * this app's Asia/Kolkata timezone (UTC+5:30 — appsscript.json) the UTC value
 * lands 5.5h later, so a PO dated the SAME day as the bill computed
 * poTime > billTime and was wrongly excluded from auto-match — the operator
 * saw "no matching PO" for a bill raised the same day as its PO.
 *
 * The fix parses poDateRaw with toSafeDateObject too, so both sides are local
 * midnight and the filter is genuinely date-granular. This test pins the
 * process timezone to Asia/Kolkata so the hazard is actually exercised on any
 * host, and asserts (A) a same-day PO is now matched, (B) the normal
 * bill-after-PO case still matches, and (C) a PO dated AFTER the bill is still
 * correctly excluded — i.e. the fix widened the boundary by exactly one day
 * (same-day), without disabling the filter's real purpose.
 *
 * Run: node .pw-test/test_po_billdate_same_day_match.js
 */
// MUST precede any Date construction — V8 reads process.env.TZ lazily on
// first use, so setting it here makes new Date(...) behave as IST regardless
// of the host machine's own timezone.
process.env.TZ = 'Asia/Kolkata';

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
  global.PO_COL = PO_COL;
`, ctx, { filename: 'expose.js' });

const { APP_CONFIG, ITEMS_COL, PO_COL, savePO, suggestPoAllocations } = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); }
  else console.log('PASS:', msg);
}

// ── Sanity: confirm the timezone hazard this test targets is actually live in
// this run (if the host somehow ignored process.env.TZ and ran UTC, the two
// parses would coincide and the test couldn't catch the regression). ──
const utcMidnight = new Date('2026-05-23').getTime();
const localMidnight = new Date(2026, 4, 23).getTime();
assert(utcMidnight !== localMidnight,
  `UTC-parsed and local-parsed midnight differ under IST (hazard is live; delta ${(utcMidnight - localMidnight) / 3600000}h)`);

// ── Seed masters + a PO/Bill sheet (same shape as test_po_bill_read_count.js) ──
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

// PO dated exactly 2026-05-23 for Acme Vendor.
const poRes = savePO({
  poNumber: '', poDate: '2026-05-23', vendor: 'Acme Vendor', contact: '',
  poDescription: '', poRemarks: '', supplierRemarks: '',
  items: JSON.stringify([{ name: 'Widget', size: '', narration: '', qty: 10, unit: 'Pcs', price: 50 }])
});
assert(poRes.success, 'savePO (Acme, 2026-05-23) succeeds: ' + poRes.message);

// A separate PO dated 2026-05-25 (in the FUTURE relative to the 05-23 bill in
// test C) for a different vendor, to prove the after-the-bill exclusion holds.
const futureRes = savePO({
  poNumber: '', poDate: '2026-05-25', vendor: 'Future Vendor', contact: '',
  poDescription: '', poRemarks: '', supplierRemarks: '',
  items: JSON.stringify([{ name: 'Widget', size: '', narration: '', qty: 10, unit: 'Pcs', price: 50 }])
});
assert(futureRes.success, 'savePO (Future, 2026-05-25) succeeds: ' + futureRes.message);

const item = [{ rowIndex: 0, name: 'Widget', size: '', narration: '', qty: 5, unit: 'Pcs', price: 50 }];
const matched = (res) => Array.isArray(res.data) && res.data[0] && Array.isArray(res.data[0].allocations) && res.data[0].allocations.length > 0;

console.log('\n=== A) Bill dated the SAME day as the PO (the regression) ===');
const sameDay = suggestPoAllocations('Acme Vendor', item, '2026-05-23');
assert(sameDay.success, 'suggestPoAllocations succeeds: ' + sameDay.message);
assert(matched(sameDay), 'a PO dated the same day as the bill is offered for auto-match (was wrongly excluded before the fix)');

console.log('\n=== B) Bill dated the day AFTER the PO (control — always worked) ===');
const dayAfter = suggestPoAllocations('Acme Vendor', item, '2026-05-24');
assert(matched(dayAfter), 'a PO dated the day before the bill is still offered');

console.log('\n=== C) PO dated AFTER the bill (filter purpose preserved) ===');
const futureExcluded = suggestPoAllocations('Future Vendor', item, '2026-05-23');
assert(futureExcluded.success, 'suggestPoAllocations succeeds: ' + futureExcluded.message);
assert(!matched(futureExcluded), 'a PO dated after the bill is still excluded (fix did not disable the boundary, only made it inclusive of the same day)');

console.log(failures === 0 ? '\nALL TESTS PASSED' : `\n${failures} TEST(S) FAILED`);
process.exitCode = failures === 0 ? 0 : 1;
