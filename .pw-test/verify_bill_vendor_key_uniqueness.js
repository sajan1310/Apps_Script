/**
 * Standalone Node harness (mirrors .pw-test/test_merge_and_backfill.js) that
 * mocks the Apps Script runtime well enough to load and run the REAL
 * module_bill.js against in-memory fake sheets, to prove that bill
 * uniqueness is now keyed on (vendor, bill number) instead of bill number
 * alone — i.e. two different vendors CAN share a bill number, but the same
 * vendor CANNOT reuse one, and edit/delete/aggregation all correctly scope
 * to the (vendor, bill number) pair instead of colliding across vendors.
 *
 * Run: node .pw-test/verify_bill_vendor_key_uniqueness.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

// ─────────────────────────────────────────────────────────────────────────
// Minimal in-memory Sheet/Range/Spreadsheet mocks (same shape as
// test_merge_and_backfill.js's harness)
// ─────────────────────────────────────────────────────────────────────────

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
  }
  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowArr = [];
      for (let c = 0; c < this.numCols; c++) {
        rowArr.push(this.sheet._get(this.row + r, this.col + c));
      }
      out.push(rowArr);
    }
    return out;
  }
  getValue() {
    return this.sheet._get(this.row, this.col);
  }
  setValues(values) {
    values.forEach((rowArr, r) => {
      rowArr.forEach((val, c) => {
        this.sheet._set(this.row + r, this.col + c, val);
      });
    });
    return this;
  }
  setValue(v) {
    this.sheet._set(this.row, this.col, v);
    return this;
  }
  clearContent() {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) {
        this.sheet._set(this.row + r, this.col + c, '');
      }
    }
    return this;
  }
}

class FakeSheet {
  constructor(name) {
    this.name = name;
    this.rows = []; // rows[0] is sheet row 1
  }
  _ensureRow(r) {
    while (this.rows.length < r) this.rows.push([]);
  }
  _get(r, c) {
    this._ensureRow(r);
    const row = this.rows[r - 1];
    return row[c - 1] === undefined ? '' : row[c - 1];
  }
  _set(r, c, v) {
    this._ensureRow(r);
    const row = this.rows[r - 1];
    while (row.length < c) row.push('');
    row[c - 1] = v;
  }
  getName() { return this.name; }
  getLastRow() {
    for (let r = this.rows.length; r >= 1; r--) {
      const row = this.rows[r - 1];
      if (row.some(v => v !== '' && v !== undefined && v !== null)) return r;
    }
    return 0;
  }
  getLastColumn() {
    let max = 0;
    this.rows.forEach(row => {
      for (let c = row.length; c >= 1; c--) {
        if (row[c - 1] !== '' && row[c - 1] !== undefined && row[c - 1] !== null) {
          max = Math.max(max, c);
          break;
        }
      }
    });
    return max;
  }
  getRange(row, col, numRows = 1, numCols = 1) {
    return new FakeRange(this, row, col, numRows, numCols);
  }
  appendRow(arr) {
    const r = this.getLastRow() + 1;
    arr.forEach((v, i) => this._set(r, i + 1, v));
  }
  deleteRows(r, n) {
    this.rows.splice(r - 1, n);
  }
  insertRows(r, n) {
    for (let i = 0; i < n; i++) this.rows.splice(r - 1, 0, []);
  }
}

class FakeSpreadsheet {
  constructor() {
    this.sheets = {};
  }
  getSheetByName(name) {
    return this.sheets[name] || null;
  }
  addSheet(name) {
    const s = new FakeSheet(name);
    this.sheets[name] = s;
    return s;
  }
}

// ─────────────────────────────────────────────────────────────────────────
// Build sandbox, load real files into it
// ─────────────────────────────────────────────────────────────────────────

const ss = new FakeSpreadsheet();

const sandbox = {
  SpreadsheetApp: {
    getActiveSpreadsheet: () => ss,
    flush: () => {}
  },
  LockService: {
    getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} })
  },
  CacheService: {
    // Short-circuits _ensureBillSheetColumns' one-time migration path —
    // this test only cares about the uniqueness/edit/delete keying, not
    // the header self-heal migration.
    getScriptCache: () => ({ get: () => '1', put: () => {} })
  },
  console,
  Logger: { log: () => {} },
  Utilities: { getUuid: () => 'uuid-' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) },
  recalculateStock: undefined
};
sandbox.global = sandbox;

const ctx = vm.createContext(sandbox);

// module_stock.js / module_po.js are deliberately NOT loaded — this test
// isolates module_bill.js's own uniqueness/edit/delete logic. Their absence
// is safe: saveBill() only calls recalculateStock() if it's a function, and
// autoExtractFromPoOrBill()'s call to syncStockForItem() is inside its own
// try/catch, so an undefined reference there is swallowed, not fatal.
const files = [
  'config.js',
  'utils.js',
  'module_units.js',
  'module_items.js',
  'module_vendors.js',
  'module_bill.js'
];

files.forEach(f => {
  const code = fs.readFileSync(path.join(ROOT, f), 'utf8');
  vm.runInContext(code, ctx, { filename: f });
});

vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.BILL_COL = BILL_COL;
`, ctx, { filename: 'expose.js' });

const { APP_CONFIG, BILL_COL, saveBill, getBillData, deleteBill, deleteBillsBulk } = ctx;

// Real Google Sheets always have a header row already in place, so
// getLastRow() is never 0 before the first save (DATA_START_ROW=2 assumes
// row 1 is a pre-existing header). Seed one so the fake sheet matches that
// invariant instead of letting the very first save land in row 1.
const billSheet = ss.addSheet(APP_CONFIG.SHEETS.BILL);
billSheet._set(1, BILL_COL.BILL_NUMBER, 'Bill Number');

// Items Master / Unit Master / Vendors are all optional side-effect lookups
// during saveBill() (item-unit-conversion + vendor/rate auto-extraction) —
// module_items.js / module_units.js / module_vendors.js already fall back
// gracefully when these are missing, but they log a getSheet() error each
// time. Seed empty sheets purely to keep this test's output readable.
ss.addSheet(APP_CONFIG.SHEETS.ITEMS);
ss.addSheet(APP_CONFIG.SHEETS.UNITS);
ss.addSheet(APP_CONFIG.SHEETS.VENDORS);

// ─────────────────────────────────────────────────────────────────────────

let failures = 0;
function assert(cond, msg) {
  if (!cond) {
    failures++;
    console.error('FAIL:', msg);
  } else {
    console.log('PASS:', msg);
  }
}

function makeFormData(overrides) {
  return Object.assign({
    billNumber: '100',
    vendor: 'Vendor A',
    contact: 'Ludhiana',
    remarks: '',
    billDate: '2026-01-01',
    poNumbers: '[]',
    items: JSON.stringify([
      { name: 'Widget', size: 'Standard', narration: '', unit: 'Pcs', qty: 10, price: 50, gst: 18, po: 'DIRECT' }
    ])
  }, overrides);
}

// 1. Save Bill #100 for Vendor A — should succeed.
const r1 = saveBill(makeFormData({ vendor: 'Vendor A' }));
assert(r1.success, `Save Bill #100 for Vendor A succeeds (got: ${r1.message})`);

// 2. Save Bill #100 for Vendor B — different vendor, same number — should
//    ALSO succeed now (this is the actual behavior being fixed).
const r2 = saveBill(makeFormData({ vendor: 'Vendor B', remarks: 'From Vendor B' }));
assert(r2.success, `Save Bill #100 for Vendor B (different vendor) succeeds (got: ${r2.message})`);

// 3. Save Bill #100 for Vendor A AGAIN (new record, not an edit) — same
//    vendor + same bill number — must be rejected as a duplicate.
const r3 = saveBill(makeFormData({ vendor: 'Vendor A' }));
assert(!r3.success, 'Save Bill #100 for Vendor A again (true duplicate) is rejected');
assert(/already exists/i.test(r3.message || ''), `Duplicate rejection message mentions "already exists" (got: "${r3.message}")`);

// 4. getBillData() must return TWO distinct bills, not merge them into one.
const listResp = getBillData();
assert(listResp.success, 'getBillData() succeeds');
const bills = listResp.data || [];
assert(bills.length === 2, `getBillData() returns 2 distinct bills, not merged (got ${bills.length})`);
const vendorANames = bills.filter(b => b.vendor === 'Vendor A');
const vendorBNames = bills.filter(b => b.vendor === 'Vendor B');
assert(vendorANames.length === 1 && vendorANames[0].billNumber === '100', 'Vendor A bill #100 present with correct vendor');
assert(vendorBNames.length === 1 && vendorBNames[0].billNumber === '100', 'Vendor B bill #100 present with correct vendor');
assert(vendorBNames[0].remarks === 'From Vendor B', "Vendor B's bill kept its own remarks (not merged with Vendor A's)");

// 5. Edit Vendor A's bill #100 (existingVendor + existingBillNumber set) —
//    must only touch Vendor A's rows, leaving Vendor B's #100 untouched.
const editResp = saveBill(makeFormData({
  vendor: 'Vendor A',
  remarks: 'Edited remarks for A',
  existingBillNumber: '100',
  existingVendor: 'Vendor A'
}));
assert(editResp.success, `Edit Vendor A's Bill #100 succeeds (got: ${editResp.message})`);

const afterEdit = getBillData().data || [];
const aAfter = afterEdit.find(b => b.vendor === 'Vendor A');
const bAfter = afterEdit.find(b => b.vendor === 'Vendor B');
assert(afterEdit.length === 2, `Still exactly 2 bills after editing Vendor A's (got ${afterEdit.length})`);
assert(aAfter && aAfter.remarks === 'Edited remarks for A', "Vendor A's bill reflects the edit");
assert(bAfter && bAfter.remarks === 'From Vendor B', "Vendor B's bill #100 was NOT touched by editing Vendor A's #100");

// 6. deleteBill(vendor, billNumber) must only remove the targeted vendor's
//    bill, not the other vendor's same-numbered one.
const delResp = deleteBill('Vendor B', '100');
assert(delResp.success, `deleteBill('Vendor B', '100') succeeds (got: ${delResp.message})`);

const afterDelete = getBillData().data || [];
assert(afterDelete.length === 1, `Exactly 1 bill remains after deleting Vendor B's #100 (got ${afterDelete.length})`);
assert(afterDelete[0].vendor === 'Vendor A', "The remaining bill is Vendor A's (Vendor B's was deleted, not Vendor A's)");

// 7. deleteBillsBulk([{vendor, billNumber}]) removes the rest.
const bulkResp = deleteBillsBulk([{ vendor: 'Vendor A', billNumber: '100' }]);
assert(bulkResp.success, `deleteBillsBulk removes Vendor A's #100 (got: ${bulkResp.message})`);
const afterBulk = getBillData().data || [];
assert(afterBulk.length === 0, `No bills remain after bulk-deleting the last one (got ${afterBulk.length})`);

// ─────────────────────────────────────────────────────────────────────────

if (failures > 0) {
  console.error(`\n${failures} assertion(s) FAILED.`);
  process.exit(1);
} else {
  console.log('\nAll assertions PASSED.');
}
