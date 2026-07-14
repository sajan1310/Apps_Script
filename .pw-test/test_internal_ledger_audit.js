/**
 * Standalone Node harness (same pattern as test_po_bill_read_count.js /
 * test_bom_recipe_drift.js) covering module_audit.js — the backend-only
 * internal ledger audit. Unlike the front-end-facing diagnostics elsewhere
 * in this codebase (e.g. getBomProcessComponentsDrift), this one is NOT
 * called from the web app: it runs on a scheduled trigger
 * (installInternalLedgerAuditTrigger) and writes findings straight to the
 * Logs sheet via logAction(), so this test asserts against the fake Logs
 * sheet's rows rather than an API response rendered in a modal.
 *
 * Covers all 4 checks:
 *  1. over_billed_po_line      — billed qty on a PO line exceeds ordered qty
 *  2. orphaned_return_bill_ref — Return names a Bill Number that doesn't exist
 *  3. return_item_not_on_bill  — Return item isn't on the Bill it references
 *  4. over_consumed_item       — total Returned + Wasted + Issued qty for an
 *     item exceeds total Billed qty for that item
 *
 * Run: node .pw-test/test_internal_ledger_audit.js
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
  insertSheet(name) { return this.addSheet(name); }
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
      return fmt === 'dd/MM/yyyy' ? `${d}/${m}/${y}` : `${y}${m}${d}`;
    }
  },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }), getScriptTimeZone: () => 'Asia/Kolkata' }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);

const files = [
  'config.js', 'utils.js', 'module_units.js', 'module_items.js',
  'module_stock.js', 'module_bill.js', 'module_po.js', 'module_return.js',
  'module_wastage.js', 'module_issue.js', 'module_vendors.js', 'module_audit.js',
  'module_dashboard.js'
];
files.forEach(f => vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f }));

vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.ITEMS_COL = ITEMS_COL;
  global.STOCK_COL = STOCK_COL;
  global.BILL_COL = BILL_COL;
  global.PO_COL = PO_COL;
  global.RETURN_COL = RETURN_COL;
  global.UNITS_COL = UNITS_COL;
  global.LOGS_COL = LOGS_COL;
`, ctx, { filename: 'expose.js' });

const {
  APP_CONFIG, ITEMS_COL, PO_COL, LOGS_COL,
  savePO, saveBill, saveReturn, saveWastage, saveIssueStock,
  _computeInternalLedgerAuditFindings, runInternalLedgerAudit
} = ctx;

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
itemsSheet._set(2, ITEMS_COL.ITEM_NAME, 'Widget');
itemsSheet._set(2, ITEMS_COL.SIZE, '');
itemsSheet._set(2, ITEMS_COL.BASE_UNIT, 'Pcs');
itemsSheet._set(2, ITEMS_COL.PURCHASE_UNIT, 'Pcs');
itemsSheet._set(3, ITEMS_COL.ITEM_NAME, 'Bolt');
itemsSheet._set(3, ITEMS_COL.SIZE, '');
itemsSheet._set(3, ITEMS_COL.BASE_UNIT, 'Pcs');
itemsSheet._set(3, ITEMS_COL.PURCHASE_UNIT, 'Pcs');
itemsSheet._set(4, ITEMS_COL.ITEM_NAME, 'Screw');
itemsSheet._set(4, ITEMS_COL.SIZE, '');
itemsSheet._set(4, ITEMS_COL.BASE_UNIT, 'Pcs');
itemsSheet._set(4, ITEMS_COL.PURCHASE_UNIT, 'Pcs');

const vendorsSheet = ss.addSheet(APP_CONFIG.SHEETS.VENDORS);
vendorsSheet._set(1, 1, 'Vendor Name');

const poSheet = ss.addSheet(APP_CONFIG.SHEETS.PO);
poSheet._set(2, PO_COL.PO_NUMBER, 'PO Number');
const billSheet = ss.addSheet(APP_CONFIG.SHEETS.BILL);
billSheet._set(1, 1, 'PO Number');
const returnSheet = ss.addSheet(APP_CONFIG.SHEETS.RETURN);
returnSheet._set(1, 1, 'Return Number');
const wastageSheet = ss.addSheet(APP_CONFIG.SHEETS.WASTAGE);
wastageSheet._set(1, 1, 'Wastage ID');
const issueSheet = ss.addSheet(APP_CONFIG.SHEETS.ISSUE);
issueSheet._set(1, 1, 'Issue ID');
const logsSheet = ss.addSheet(APP_CONFIG.SHEETS.LOGS);
logsSheet._set(1, 1, 'Timestamp'); logsSheet._set(1, 2, 'User'); logsSheet._set(1, 3, 'Action');
logsSheet._set(1, 4, 'Sheet'); logsSheet._set(1, 5, 'RecordId'); logsSheet._set(1, 6, 'Details'); logsSheet._set(1, 7, 'Status');

// ── Fixture: PO for 10 Widgets @ ₹50, billed 15 total (over-billed by 5) ──
const poRes = savePO({
  poNumber: '', poDate: '2026-01-01', vendor: 'Acme Vendor', contact: '',
  poDescription: '', poRemarks: '', supplierRemarks: '',
  items: JSON.stringify([{ name: 'Widget', size: '', narration: '', qty: 10, unit: 'Pcs', price: 50 }])
});
assert(poRes.success, 'PO for 10 Widgets saved: ' + poRes.message);
const poNumber = poRes.data.poNumber;

const bill1Res = saveBill({
  billNumber: 'BILL-1', billDate: '2026-01-05', vendor: 'Acme Vendor', contact: '', remarks: '',
  poNumbers: JSON.stringify([poNumber]),
  items: JSON.stringify([{ name: 'Widget', size: '', narration: '', qty: 10, unit: 'Pcs', price: 50, po: poNumber }])
});
assert(bill1Res.success, 'Bill #1 (10 Widgets) saved: ' + bill1Res.message);

const bill2Res = saveBill({
  billNumber: 'BILL-2', billDate: '2026-01-10', vendor: 'Acme Vendor', contact: '', remarks: '',
  poNumbers: JSON.stringify([poNumber]),
  items: JSON.stringify([{ name: 'Widget', size: '', narration: '', qty: 5, unit: 'Pcs', price: 50, po: poNumber }])
});
assert(bill2Res.success, 'Bill #2 (5 more Widgets, over-billing the PO) saved: ' + bill2Res.message);

// ── Fixture: a clean Bill for 20 Bolts and 10 Screws, used by the return/
// wastage/issue checks below ──
const bill3Res = saveBill({
  billNumber: 'BILL-3', billDate: '2026-01-06', vendor: 'Acme Vendor', contact: '', remarks: '',
  items: JSON.stringify([
    { name: 'Bolt', size: '', narration: '', qty: 20, unit: 'Pcs', price: 5, po: 'DIRECT' },
    { name: 'Screw', size: '', narration: '', qty: 10, unit: 'Pcs', price: 2, po: 'DIRECT' }
  ])
});
assert(bill3Res.success, 'Bill #3 (20 Bolts + 10 Screws) saved: ' + bill3Res.message);

// ── Fixture: Return referencing a Bill Number that does not exist ──
const returnOrphanRes = saveReturn({
  returnNumber: 'RET-1', returnDate: '2026-01-12', vendor: 'Acme Vendor', contact: '',
  billNumber: 'BILL-999', remarks: '',
  items: JSON.stringify([{ name: 'Bolt', size: '', narration: '', qty: 2, unit: 'Pcs', price: 5, reason: 'Defective' }])
});
assert(returnOrphanRes.success, 'Return referencing a nonexistent Bill saved: ' + returnOrphanRes.message);

// ── Fixture: Return referencing a real Bill (#3) but for an item that
// bill never billed ──
const returnMismatchRes = saveReturn({
  returnNumber: 'RET-2', returnDate: '2026-01-13', vendor: 'Acme Vendor', contact: '',
  billNumber: 'BILL-3', remarks: '',
  items: JSON.stringify([{ name: 'Widget', size: '', narration: '', qty: 1, unit: 'Pcs', price: 50, reason: 'Wrong item' }])
});
assert(returnMismatchRes.success, 'Return referencing real Bill #3 but wrong item saved: ' + returnMismatchRes.message);

// ── Fixture: Bolts over-consumed once Wastage + Issue are combined with a
// smaller Return: 20 billed, but 5 returned + 10 wasted + 10 issued = 25 ──
const returnBoltRes = saveReturn({
  returnNumber: 'RET-3', returnDate: '2026-01-14', vendor: 'Acme Vendor', contact: '',
  billNumber: '', remarks: '',
  items: JSON.stringify([{ name: 'Bolt', size: '', narration: '', qty: 5, unit: 'Pcs', price: 5, reason: 'Excess' }])
});
assert(returnBoltRes.success, 'Return of 5 Bolts saved: ' + returnBoltRes.message);

const wastageRes = saveWastage({
  date: '2026-01-15', vendor: '', remarks: '',
  items: JSON.stringify([{ name: 'Bolt', size: '', qty: 10, unit: 'Pcs', reason: 'Damaged in handling' }])
});
assert(wastageRes.success, 'Wastage of 10 Bolts saved: ' + wastageRes.message);

const issueRes = saveIssueStock({
  date: '2026-01-16', issuedTo: 'Production Floor', reference: '', remarks: '',
  items: JSON.stringify([{ name: 'Bolt', size: '', qty: 10, unit: 'Pcs' }])
});
assert(issueRes.success, 'Issue of 10 Bolts saved: ' + issueRes.message);

// Screws stay untouched by Return/Wastage/Issue -- must NOT show up as
// over-consumed just because it appears in a "consumedByItem" map at all.

console.log('\n=== _computeInternalLedgerAuditFindings: runs and returns all 4 checks ===');
const result = _computeInternalLedgerAuditFindings();
const findings = result.findings;

console.log('\n--- Check 1: over_billed_po_line ---');
{
  const f = findings.filter(x => x.type === 'over_billed_po_line');
  assert(f.length === 1, `exactly 1 over-billed PO line found (got ${f.length})`);
  if (f[0]) {
    assert(f[0].poNumber === poNumber, `flags the right PO (got ${f[0].poNumber})`);
    assert(f[0].orderedQty === 10 && f[0].billedQty === 15 && f[0].overBy === 5,
      `ordered/billed/overBy = 10/15/5 (got ${f[0].orderedQty}/${f[0].billedQty}/${f[0].overBy})`);
  }
}

console.log('\n--- Check 2: orphaned_return_bill_ref ---');
{
  const f = findings.filter(x => x.type === 'orphaned_return_bill_ref');
  assert(f.length === 1, `exactly 1 orphaned return->bill reference found (got ${f.length})`);
  if (f[0]) assert(f[0].returnNumber === 'RET-1' && f[0].billNumber === 'BILL-999', `flags RET-1 -> BILL-999 (got ${f[0].returnNumber} -> ${f[0].billNumber})`);
  assert(!findings.some(x => x.type === 'orphaned_return_bill_ref' && x.returnNumber === 'RET-3'),
    'RET-3 (no bill reference at all) is NOT flagged as orphaned -- the field is optional');
}

console.log('\n--- Check 3: return_item_not_on_bill ---');
{
  const f = findings.filter(x => x.type === 'return_item_not_on_bill');
  assert(f.length === 1, `exactly 1 return-item-not-on-bill found (got ${f.length})`);
  if (f[0]) assert(f[0].returnNumber === 'RET-2' && f[0].itemName === 'Widget', `flags RET-2/Widget (got ${f[0].returnNumber}/${f[0].itemName})`);
}

console.log('\n--- Check 4: over_consumed_item ---');
{
  const f = findings.filter(x => x.type === 'over_consumed_item');
  assert(f.length === 1, `exactly 1 over-consumed item found (got ${f.length}: ${JSON.stringify(f)})`);
  if (f[0]) {
    assert(f[0].itemName === 'Bolt', `flags Bolt (got ${f[0].itemName})`);
    // Bolt returns come from two records: RET-1 (2, orphaned-bill-ref fixture)
    // + RET-3 (5, this fixture) = 7 total.
    assert(f[0].totalBilledBaseQty === 20, `total billed = 20 (got ${f[0].totalBilledBaseQty})`);
    assert(f[0].totalReturnedBaseQty === 7, `total returned = 7 (2 from RET-1 + 5 from RET-3) (got ${f[0].totalReturnedBaseQty})`);
    assert(f[0].totalWastedBaseQty === 10, `total wasted = 10 (got ${f[0].totalWastedBaseQty})`);
    assert(f[0].totalIssuedBaseQty === 10, `total issued = 10 (got ${f[0].totalIssuedBaseQty})`);
    assert(f[0].totalConsumedBaseQty === 27, `total consumed = 27 (got ${f[0].totalConsumedBaseQty})`);
    assert(f[0].overBy === 7, `over by 7 (got ${f[0].overBy})`);
  }
  assert(!findings.some(x => x.type === 'over_consumed_item' && x.itemName === 'Screw'),
    'Screw (never returned/wasted/issued) is NOT flagged');
}

console.log('\n=== runInternalLedgerAudit: writes findings to the Logs sheet, not any front-end-reachable output ===');
{
  const beforeRowCount = logsSheet.getLastRow();
  const runRes = runInternalLedgerAudit();
  assert(runRes.success, 'runInternalLedgerAudit succeeds: ' + runRes.message);
  assert(runRes.data.findingsCount === findings.length, `response reports the same finding count (got ${runRes.data.findingsCount}, expected ${findings.length})`);

  const newRows = logsSheet.rows.slice(beforeRowCount);
  const findingRows = newRows.filter(r => r[LOGS_COL.ACTION - 1] === 'LEDGER_AUDIT_FINDING');
  const summaryRows = newRows.filter(r => r[LOGS_COL.ACTION - 1] === 'LEDGER_AUDIT_SUMMARY');

  assert(findingRows.length === findings.length, `one LEDGER_AUDIT_FINDING log row per finding (got ${findingRows.length}, expected ${findings.length})`);
  assert(findingRows.every(r => r[LOGS_COL.STATUS - 1] === 'WARNING'), 'every finding row is logged with status WARNING');
  assert(findingRows.every(r => r[LOGS_COL.SHEET - 1] === 'Internal Audit'), 'every finding row is scoped to the "Internal Audit" sheet label (not a real ledger sheet name)');

  assert(summaryRows.length === 1, `exactly 1 LEDGER_AUDIT_SUMMARY row written (got ${summaryRows.length})`);
  assert(summaryRows[0][LOGS_COL.STATUS - 1] === 'SUCCESS', 'summary row is logged with status SUCCESS (the audit RAN fine, even though it found problems)');
}

console.log('\n=== Clean state: zero findings, still writes exactly one SUCCESS summary row ===');
{
  const ss2 = new FakeSpreadsheet();
  const ctx2 = vm.createContext(Object.assign({}, sandbox, {
    SpreadsheetApp: { getActiveSpreadsheet: () => ss2, flush: () => {} }
  }));
  ctx2.global = ctx2;
  files.forEach(f => vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx2, { filename: f }));
  vm.runInContext(`
    global.APP_CONFIG = APP_CONFIG; global.ITEMS_COL = ITEMS_COL; global.PO_COL = PO_COL; global.LOGS_COL = LOGS_COL;
  `, ctx2, { filename: 'expose2.js' });

  const u2 = ss2.addSheet(ctx2.APP_CONFIG.SHEETS.UNITS);
  u2._set(1, 1, 'Unit Name'); u2._set(1, 2, 'Family'); u2._set(1, 3, 'Factor'); u2._set(1, 4, 'Remarks');
  u2._set(2, 1, 'Pcs'); u2._set(2, 2, 'Count'); u2._set(2, 3, 1); u2._set(2, 4, '');
  const i2 = ss2.addSheet(ctx2.APP_CONFIG.SHEETS.ITEMS);
  i2._set(1, 1, 'Item Name');
  i2._set(2, ctx2.ITEMS_COL.ITEM_NAME, 'Widget'); i2._set(2, ctx2.ITEMS_COL.BASE_UNIT, 'Pcs'); i2._set(2, ctx2.ITEMS_COL.PURCHASE_UNIT, 'Pcs');
  const v2 = ss2.addSheet(ctx2.APP_CONFIG.SHEETS.VENDORS);
  v2._set(1, 1, 'Vendor Name');
  const p2 = ss2.addSheet(ctx2.APP_CONFIG.SHEETS.PO);
  p2._set(2, ctx2.PO_COL.PO_NUMBER, 'PO Number');
  const b2 = ss2.addSheet(ctx2.APP_CONFIG.SHEETS.BILL);
  b2._set(1, 1, 'PO Number');
  const r2 = ss2.addSheet(ctx2.APP_CONFIG.SHEETS.RETURN);
  r2._set(1, 1, 'Return Number');
  const w2 = ss2.addSheet(ctx2.APP_CONFIG.SHEETS.WASTAGE);
  w2._set(1, 1, 'Wastage ID');
  const iss2 = ss2.addSheet(ctx2.APP_CONFIG.SHEETS.ISSUE);
  iss2._set(1, 1, 'Issue ID');
  const logs2 = ss2.addSheet(ctx2.APP_CONFIG.SHEETS.LOGS);
  logs2._set(1, 1, 'Timestamp');

  const poClean = ctx2.savePO({
    poNumber: '', poDate: '2026-01-01', vendor: 'Clean Vendor', contact: '',
    poDescription: '', poRemarks: '', supplierRemarks: '',
    items: JSON.stringify([{ name: 'Widget', size: '', narration: '', qty: 10, unit: 'Pcs', price: 50 }])
  });
  ctx2.saveBill({
    billNumber: 'C-BILL-1', billDate: '2026-01-05', vendor: 'Clean Vendor', contact: '', remarks: '',
    poNumbers: JSON.stringify([poClean.data.poNumber]),
    items: JSON.stringify([{ name: 'Widget', size: '', narration: '', qty: 10, unit: 'Pcs', price: 50, po: poClean.data.poNumber }])
  });

  const cleanRunRes = ctx2.runInternalLedgerAudit();
  assert(cleanRunRes.success, 'clean-state run succeeds: ' + cleanRunRes.message);
  assert(cleanRunRes.data.findingsCount === 0, `no findings for a fully clean PO+Bill pair (got ${cleanRunRes.data.findingsCount})`);
  assert(/no discrepancies found/i.test(cleanRunRes.message || ''), `message confirms clean state (got "${cleanRunRes.message}")`);

  const logRows = logs2.rows.slice(1); // skip header
  const summaryRows2 = logRows.filter(r => r[ctx2.LOGS_COL.ACTION - 1] === 'LEDGER_AUDIT_SUMMARY');
  assert(summaryRows2.length === 1, `exactly 1 summary row written even with zero findings (got ${summaryRows2.length})`);
  assert(summaryRows2[0][ctx2.LOGS_COL.STATUS - 1] === 'SUCCESS', 'summary row for a clean run is still status SUCCESS');
  assert(logRows.filter(r => r[ctx2.LOGS_COL.ACTION - 1] === 'LEDGER_AUDIT_FINDING').length === 0, 'no finding rows written when there are none');
}

console.log(failures === 0 ? '\nALL TESTS PASSED' : `\n${failures} TEST(S) FAILED`);
process.exitCode = failures === 0 ? 0 : 1;
