/**
 * Standalone Node harness (same pattern as test_po_status.js) that loads the
 * REAL config.js + utils.js and exercises repairCorruptedContacts() against
 * fake Vendors/Clients/Contractors sheets seeded with cells corrupted by the
 * pre-fix sanitizeString() (which used to escape ANY leading '+', turning a
 * phone number like "+91 98765 43210" into the literal string
 * "'+91 98765 43210" via setValue()).
 *
 * Run: node .pw-test/test_repair_corrupted_contacts.js
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
  getRange(row, col, numRows = 1, numCols = 1) { return new FakeRange(this, row, col, numRows, numCols); }
}

class FakeSpreadsheet {
  constructor() { this.sheets = {}; }
  getSheetByName(name) { return this.sheets[name] || null; }
  addSheet(name) { const s = new FakeSheet(name); this.sheets[name] = s; return s; }
}

const ss = new FakeSpreadsheet();

const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  Logger: { log: () => {} },
  console
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);

['config.js', 'utils.js'].forEach(f =>
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f }));

// `const`-declared top-level bindings (unlike `function` declarations) don't
// become properties of the vm context object, so they must be re-exposed
// explicitly (same workaround as test_po_status.js).
vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.VENDORS_COL = VENDORS_COL;
  global.CLIENTS_COL = CLIENTS_COL;
  global.CONTRACTORS_COL = CONTRACTORS_COL;
`, ctx, { filename: 'expose.js' });

const { APP_CONFIG, VENDORS_COL, CLIENTS_COL, CONTRACTORS_COL, repairCorruptedContacts } = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); }
  else console.log('PASS:', msg);
}

// ── Seed Vendors: header row 1, data from row 2 ─────────────────────────
const vendorsSheet = ss.addSheet(APP_CONFIG.SHEETS.VENDORS);
vendorsSheet._set(1, 1, 'Vendor Name'); vendorsSheet._set(1, 2, 'Contact'); vendorsSheet._set(1, 5, 'Remarks');
vendorsSheet._set(2, VENDORS_COL.VENDOR_NAME, 'Acme Vendor');
vendorsSheet._set(2, VENDORS_COL.CONTACT, "'+91 98765 43210"); // corrupted phone
vendorsSheet._set(2, VENDORS_COL.REMARKS, "'-500 pending");     // intentionally-escaped negative, must NOT be touched
vendorsSheet._set(3, VENDORS_COL.VENDOR_NAME, 'Beta Vendor');
vendorsSheet._set(3, VENDORS_COL.CONTACT, '022-12345678');       // never corrupted, untouched
vendorsSheet._set(3, VENDORS_COL.REMARKS, 'Reliable supplier');  // plain text, untouched

// ── Seed Clients: same shape ─────────────────────────────────────────────
const clientsSheet = ss.addSheet(APP_CONFIG.SHEETS.CLIENTS);
clientsSheet._set(1, 1, 'Client Name');
clientsSheet._set(2, CLIENTS_COL.CLIENT_NAME, 'Client A');
clientsSheet._set(2, CLIENTS_COL.CONTACT, "'+1 (555) 123-4567"); // corrupted phone

// ── Seed Contractors: same shape, plus a non-string cell to ignore ───────
const contractorsSheet = ss.addSheet(APP_CONFIG.SHEETS.CONTRACTORS);
contractorsSheet._set(1, 1, 'Contractor Name');
contractorsSheet._set(2, CONTRACTORS_COL.CONTRACTOR_NAME, 'Contractor X');
contractorsSheet._set(2, CONTRACTORS_COL.CONTACT, "'+91 99999 00000"); // corrupted phone
contractorsSheet._set(2, CONTRACTORS_COL.REMARKS, 42); // numeric cell, must be skipped, not stringified

const res = repairCorruptedContacts();

assert(res.success, 'repairCorruptedContacts() returns success');
assert(res.data.repairedCount === 3, `repaired exactly 3 cells (got ${res.data.repairedCount})`);

assert(vendorsSheet._get(2, VENDORS_COL.CONTACT) === '+91 98765 43210',
  `Vendors contact repaired (got ${JSON.stringify(vendorsSheet._get(2, VENDORS_COL.CONTACT))})`);
assert(vendorsSheet._get(2, VENDORS_COL.REMARKS) === "'-500 pending",
  `Vendors remarks with intentional '-' escape left untouched (got ${JSON.stringify(vendorsSheet._get(2, VENDORS_COL.REMARKS))})`);
assert(vendorsSheet._get(3, VENDORS_COL.CONTACT) === '022-12345678',
  'Vendors never-corrupted contact left untouched');
assert(vendorsSheet._get(3, VENDORS_COL.REMARKS) === 'Reliable supplier',
  'Vendors plain-text remarks left untouched');

assert(clientsSheet._get(2, CLIENTS_COL.CONTACT) === '+1 (555) 123-4567',
  `Clients contact repaired (got ${JSON.stringify(clientsSheet._get(2, CLIENTS_COL.CONTACT))})`);

assert(contractorsSheet._get(2, CONTRACTORS_COL.CONTACT) === '+91 99999 00000',
  `Contractors contact repaired (got ${JSON.stringify(contractorsSheet._get(2, CONTRACTORS_COL.CONTACT))})`);
assert(contractorsSheet._get(2, CONTRACTORS_COL.REMARKS) === 42,
  'Contractors numeric remarks cell left untouched (not coerced to string)');

// Re-running against already-repaired sheets should be a clean no-op.
const res2 = repairCorruptedContacts();
assert(res2.success && res2.data.repairedCount === 0, 'second run is a no-op (idempotent)');

console.log(failures === 0 ? '\nALL TESTS PASSED' : `\n${failures} TEST(S) FAILED`);
process.exitCode = failures === 0 ? 0 : 1;
