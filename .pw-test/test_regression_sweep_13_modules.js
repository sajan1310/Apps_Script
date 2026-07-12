/**
 * P3 Verification — "Regression sweep": create + edit + delete one record in
 * each of PO, Bill, Item, Vendor, Client, Contractor, BOM, Process,
 * Production, Issue, Dispatch, Return, Wastage — asserting every save/delete
 * writes the correct sheet row(s) AND the correct Logs entry.
 *
 * Standalone Node harness (same mock-SpreadsheetApp pattern as
 * test_merge_and_backfill.js / test_unit_conversion.js) — loads the REAL
 * server .js files into a vm sandbox with Fake Sheet/Spreadsheet classes and
 * calls server functions directly.
 *
 * Issue and Wastage have no edit function at the API level (create + bulk-
 * delete only, confirmed by code inspection) — their "edit" step is skipped
 * and explicitly logged as such rather than faked.
 *
 * Run: node .pw-test/test_regression_sweep_13_modules.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

// ─────────────────────────────────────────────────────────────────────────
// Fake Sheet/Range/Spreadsheet/Cache — augmented with getDataRange() and
// insertColumnsAfter() (needed by Vendor/Client/Contractor saves and every
// module's self-healing ensure*Column helpers, per the research pass).
// ─────────────────────────────────────────────────────────────────────────

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

// Minimal but real date-pattern support — Issue/Wastage ID generation and
// PO's Date-object round-trip both call this unconditionally.
function pad(n, w) { return String(n).padStart(w, '0'); }
function fakeFormatDate(date, tz, pattern) {
  const d = date instanceof Date ? date : new Date(date);
  const map = {
    yyyy: d.getFullYear(), MM: pad(d.getMonth() + 1, 2), dd: pad(d.getDate(), 2),
    HH: pad(d.getHours(), 2), mm: pad(d.getMinutes(), 2), ss: pad(d.getSeconds(), 2)
  };
  return pattern.replace(/yyyy|MM|dd|HH|mm|ss/g, m => map[m]);
}

// ─────────────────────────────────────────────────────────────────────────
// Build sandbox, load real files into it
// ─────────────────────────────────────────────────────────────────────────

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

// const/let top-level bindings need re-exposing onto `global` (see other
// harnesses in this directory for why).
vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.PO_COL = PO_COL; global.BILL_COL = BILL_COL; global.ITEMS_COL = ITEMS_COL;
  global.VENDORS_COL = VENDORS_COL; global.CLIENTS_COL = CLIENTS_COL; global.CONTRACTORS_COL = CONTRACTORS_COL;
  global.BOM_COL = BOM_COL; global.PROCESS_COL = PROCESS_COL; global.PRODUCTION_COL = PRODUCTION_COL;
  global.ISSUE_COL = ISSUE_COL; global.DISPATCH_COL = DISPATCH_COL; global.RETURN_COL = RETURN_COL;
  global.WASTAGE_COL = WASTAGE_COL;
`, ctx, { filename: 'expose.js' });

const {
  APP_CONFIG, PO_COL, BILL_COL, ITEMS_COL, VENDORS_COL, CLIENTS_COL, CONTRACTORS_COL,
  BOM_COL, PROCESS_COL, PRODUCTION_COL, ISSUE_COL, DISPATCH_COL, RETURN_COL, WASTAGE_COL,
  savePO, deletePO,
  saveBill, deleteBill,
  saveItem, deleteItem,
  saveVendor, deleteVendor,
  saveClient, deleteClient,
  saveContractor, deleteContractor,
  saveBOM, deleteBOM,
  saveProcess, deleteProcess,
  saveProduction, deleteProduction,
  saveIssueStock, deleteIssueBulk,
  saveDispatch, deleteDispatch,
  saveReturn, deleteReturn,
  saveWastage, deleteWastageBulk,
  initVendorsSheet, initItemsSheet, initUnitsSheet, initClientsSheet, initContractorsSheet,
  initBOMSheet, initProcessMasterSheet, initProcessComponentsSheet, initProductionSheet,
  initIssueSheet, initDispatchSheet, initReturnSheet, initWastageSheet,
  initStockSheet, initWarehousePoolSheet, initWarehousePoolOpeningSheet
} = ctx;

// ─────────────────────────────────────────────────────────────────────────
let failures = 0;
function assert(cond, msg) {
  if (cond) { console.log('  PASS:', msg); }
  else { failures++; console.error('  FAIL:', msg); }
}

// ─────────────────────────────────────────────────────────────────────────
// One-time global setup: every sheet the sweep could touch, pre-created —
// mirrors a real spreadsheet where every tab's sheet already exists.
// ─────────────────────────────────────────────────────────────────────────
console.log('=== Setup: seeding every sheet the 13-module sweep touches ===');
ss.addSheet(APP_CONFIG.SHEETS.LOGS).appendRow(['Timestamp', 'User', 'Action', 'Sheet', 'RecordId', 'Details', 'Status']);
// Some init*Sheet() helpers self-create their sheet if missing (initVendorsSheet);
// others (initItemsSheet, etc.) only write headers to an ALREADY-existing sheet
// and throw via getSheet() otherwise. Pre-create every target sheet first so
// every init function's behavior is uniform regardless of which family it's in.
[
  APP_CONFIG.SHEETS.VENDORS, APP_CONFIG.SHEETS.ITEMS, APP_CONFIG.SHEETS.UNITS,
  APP_CONFIG.SHEETS.CLIENTS, APP_CONFIG.SHEETS.CONTRACTORS, APP_CONFIG.SHEETS.BOM,
  APP_CONFIG.SHEETS.PROCESS_MASTER, APP_CONFIG.SHEETS.PROCESS_COMPONENTS, APP_CONFIG.SHEETS.PRODUCTION,
  APP_CONFIG.SHEETS.ISSUE, APP_CONFIG.SHEETS.DISPATCH, APP_CONFIG.SHEETS.RETURN, APP_CONFIG.SHEETS.WASTAGE,
  APP_CONFIG.SHEETS.STOCK, APP_CONFIG.SHEETS.WAREHOUSE_POOL, APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING
].forEach(name => { if (!ss.getSheetByName(name)) ss.addSheet(name); });

initVendorsSheet();
initItemsSheet();
initUnitsSheet();
initClientsSheet();
initContractorsSheet();
initBOMSheet();
initProcessMasterSheet();
initProcessComponentsSheet();
initProductionSheet();
initIssueSheet();
initDispatchSheet();
initReturnSheet();
initWastageSheet();
initStockSheet();
initWarehousePoolSheet();
initWarehousePoolOpeningSheet();
// PO / Bill have no init*Sheet() helper in the codebase — seed headers by
// hand, matching test_unit_conversion.js's established pattern.
{
  const poSheet = ss.addSheet(APP_CONFIG.SHEETS.PO);
  poSheet._set(2, PO_COL.PO_NUMBER, 'PO Number'); // 2 header rows — DATA_START_ROW = 3
  const billSheet = ss.addSheet(APP_CONFIG.SHEETS.BILL);
  billSheet._set(1, BILL_COL.BILL_NUMBER, 'Bill Number');
}
console.log('  done.\n');

function logsTail(n = 1) {
  const logsSheet = ss.getSheetByName(APP_CONFIG.SHEETS.LOGS);
  const lastRow = logsSheet.getLastRow();
  const rows = [];
  for (let r = Math.max(2, lastRow - n + 1); r <= lastRow; r++) {
    rows.push(logsSheet.getRange(r, 1, 1, 7).getValues()[0]);
  }
  return rows; // [timestamp, user, action, sheet, recordId, details, status]
}
function assertLogged(actionSubstr, sheetName, msg) {
  const rows = logsTail(5);
  const found = rows.some(r => String(r[2]).includes(actionSubstr) && r[3] === sheetName && r[6] === 'SUCCESS');
  assert(found, msg + ` (recent Logs actions: ${rows.map(r => `${r[2]}/${r[3]}/${r[6]}`).join(', ')})`);
}

// ─────────────────────────────────────────────────────────────────────────
console.log('=== 1. PO ===');
{
  const createRes = savePO({
    vendor: 'Sweep Vendor', poDate: '2026-01-01',
    items: JSON.stringify([{ name: 'Sweep Item', qty: 2, unit: 'Pcs', price: 50 }])
  });
  assert(createRes.success, 'create succeeds: ' + createRes.message);
  const poNumber = createRes.data.poNumber;
  assertLogged('CREATE', APP_CONFIG.SHEETS.PO, 'Logs has a CREATE/PO entry');

  const poSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PO);
  const row = poSheet.getRange(3, PO_COL.PO_NUMBER, 1, 8).getValues()[0];
  assert(row[PO_COL.VENDOR - PO_COL.PO_NUMBER] === 'Sweep Vendor', 'sheet row has correct vendor');

  const editRes = savePO({
    existingPoNumber: poNumber, poNumber, vendor: 'Sweep Vendor', poDate: '2026-01-02',
    poRemarks: 'edited', items: JSON.stringify([{ name: 'Sweep Item', qty: 3, unit: 'Pcs', price: 50 }])
  });
  assert(editRes.success, 'edit succeeds: ' + editRes.message);
  const rowAfterEdit = poSheet.getRange(3, PO_COL.PO_NUMBER, 1, 8).getValues()[0];
  assert(rowAfterEdit[PO_COL.QTY - PO_COL.PO_NUMBER] === 3, 'sheet row reflects edited qty (got ' + rowAfterEdit[PO_COL.QTY - PO_COL.PO_NUMBER] + ')');
  // PO logs 'CREATE' even on edit (hardcoded, confirmed by research) — assert that, not 'UPDATE'.
  assertLogged('CREATE', APP_CONFIG.SHEETS.PO, 'Logs has a CREATE/PO entry for the edit too (PO always logs CREATE)');

  const deleteRes = deletePO(poNumber);
  assert(deleteRes.success, 'delete succeeds: ' + deleteRes.message);
  assert(poSheet.getLastRow() < 3, 'sheet row removed after delete');
  assertLogged('DELETE', APP_CONFIG.SHEETS.PO, 'Logs has a DELETE/PO entry');
}

console.log('\n=== 2. Bill ===');
{
  const createRes = saveBill({
    billNumber: 'SWEEP-B-1', billDate: '2026-01-01', vendor: 'Sweep Vendor',
    items: JSON.stringify([{ name: 'Sweep Item', qty: 2, unit: 'Pcs', price: 50, gst: 18 }])
  });
  assert(createRes.success, 'create succeeds: ' + createRes.message);
  assertLogged('CREATE', APP_CONFIG.SHEETS.BILL, 'Logs has a CREATE/Bill entry');

  const billSheet = ss.getSheetByName(APP_CONFIG.SHEETS.BILL);
  const row = billSheet.getRange(2, 1, 1, billSheet.getLastColumn()).getValues()[0];
  assert(row[BILL_COL.VENDOR - 1] === 'Sweep Vendor', 'sheet row has correct vendor');

  const editRes = saveBill({
    existingBillNumber: 'SWEEP-B-1', existingVendor: 'Sweep Vendor',
    billNumber: 'SWEEP-B-1', billDate: '2026-01-02', vendor: 'Sweep Vendor', remarks: 'edited',
    items: JSON.stringify([{ name: 'Sweep Item', qty: 4, unit: 'Pcs', price: 50, gst: 18 }])
  });
  assert(editRes.success, 'edit succeeds: ' + editRes.message);
  const rowAfterEdit = billSheet.getRange(2, 1, 1, billSheet.getLastColumn()).getValues()[0];
  assert(rowAfterEdit[BILL_COL.QTY - 1] === 4, 'sheet row reflects edited qty (got ' + rowAfterEdit[BILL_COL.QTY - 1] + ')');
  assertLogged('UPDATE', APP_CONFIG.SHEETS.BILL, 'Logs has an UPDATE/Bill entry for the edit');

  const deleteRes = deleteBill('Sweep Vendor', 'SWEEP-B-1');
  assert(deleteRes.success, 'delete succeeds: ' + deleteRes.message);
  assert(billSheet.getLastRow() < 2, 'sheet row removed after delete');
  assertLogged('DELETE', APP_CONFIG.SHEETS.BILL, 'Logs has a DELETE/Bill entry');
}

console.log('\n=== 3. Item ===');
{
  const createRes = saveItem({
    itemName: 'Sweep Test Item', itemSize: 'Std',
    vendors: JSON.stringify([{ vendor: 'Sweep Vendor', rate: 25 }])
  });
  assert(createRes.success, 'create succeeds: ' + createRes.message);
  assertLogged('CREATE', APP_CONFIG.SHEETS.ITEMS, 'Logs has a CREATE/Items entry');

  const itemsSheet = ss.getSheetByName(APP_CONFIG.SHEETS.ITEMS);
  const found1 = itemsSheet.getRange(2, ITEMS_COL.ITEM_NAME, itemsSheet.getLastRow() - 1, 2).getValues()
    .some(r => r[0] === 'Sweep Test Item' && r[1] === 'Std');
  assert(found1, 'sheet has the new item row');

  const editRes = saveItem({
    itemName: 'Sweep Test Item', itemSize: 'Std', itemRemarks: 'edited',
    vendors: JSON.stringify([{ vendor: 'Sweep Vendor', rate: 30 }]),
    originalName: 'Sweep Test Item', originalSize: 'Std'
  });
  assert(editRes.success, 'edit succeeds: ' + editRes.message);
  assertLogged('UPDATE', APP_CONFIG.SHEETS.ITEMS, 'Logs has an UPDATE/Items entry for the edit');

  const deleteRes = deleteItem('Sweep Test Item', 'Std');
  assert(deleteRes.success, 'delete succeeds: ' + deleteRes.message);
  const stillThere = itemsSheet.getRange(2, ITEMS_COL.ITEM_NAME, Math.max(itemsSheet.getLastRow() - 1, 0), 2).getValues()
    .some(r => r[0] === 'Sweep Test Item' && r[1] === 'Std');
  assert(!stillThere, 'sheet row removed after delete');
  assertLogged('DELETE', APP_CONFIG.SHEETS.ITEMS, 'Logs has a DELETE/Items entry');
}

console.log('\n=== 4. Vendor ===');
{
  const createRes = saveVendor({ vendorName: 'Sweep Test Vendor', contact: '9999999999' });
  assert(createRes.success, 'create succeeds: ' + createRes.message);
  assertLogged('CREATE_VENDOR', APP_CONFIG.SHEETS.VENDORS, 'Logs has a CREATE_VENDOR entry');

  const vendorsSheet = ss.getSheetByName(APP_CONFIG.SHEETS.VENDORS);
  const found1 = vendorsSheet.getRange(2, 1, vendorsSheet.getLastRow() - 1, 5).getValues()
    .some(r => r[0] === 'Sweep Test Vendor' && r[1] === '9999999999');
  assert(found1, 'sheet has the new vendor row');

  const editRes = saveVendor({ vendorName: 'Sweep Test Vendor', contact: '8888888888', originalVendorName: 'Sweep Test Vendor' });
  assert(editRes.success, 'edit succeeds: ' + editRes.message);
  assertLogged('UPDATE_VENDOR', APP_CONFIG.SHEETS.VENDORS, 'Logs has an UPDATE_VENDOR entry for the edit');

  const deleteRes = deleteVendor('Sweep Test Vendor');
  assert(deleteRes.success, 'delete succeeds: ' + deleteRes.message);
  assertLogged('DELETE_VENDOR', APP_CONFIG.SHEETS.VENDORS, 'Logs has a DELETE_VENDOR entry');
}

console.log('\n=== 5. Client ===');
{
  const createRes = saveClient({ clientName: 'Sweep Test Client', contact: '7777777777' });
  assert(createRes.success, 'create succeeds: ' + createRes.message);
  assertLogged('CREATE', APP_CONFIG.SHEETS.CLIENTS, 'Logs has a CREATE/Clients entry');

  const clientsSheet = ss.getSheetByName(APP_CONFIG.SHEETS.CLIENTS);
  const found1 = clientsSheet.getRange(2, 1, Math.max(clientsSheet.getLastRow() - 1, 0), 5).getValues()
    .some(r => r[0] === 'Sweep Test Client');
  assert(found1, 'sheet has the new client row');

  const editRes = saveClient({ clientName: 'Sweep Test Client', contact: '6666666666', originalClientName: 'Sweep Test Client' });
  assert(editRes.success, 'edit succeeds: ' + editRes.message);
  assertLogged('UPDATE', APP_CONFIG.SHEETS.CLIENTS, 'Logs has an UPDATE/Clients entry for the edit');

  const deleteRes = deleteClient('Sweep Test Client');
  assert(deleteRes.success, 'delete succeeds: ' + deleteRes.message);
  assertLogged('DELETE', APP_CONFIG.SHEETS.CLIENTS, 'Logs has a DELETE/Clients entry');
}

console.log('\n=== 6. Contractor ===');
{
  const createRes = saveContractor({ contractorName: 'Sweep Test Contractor' });
  assert(createRes.success, 'create succeeds: ' + createRes.message);
  assertLogged('CREATE_CONTRACTOR', APP_CONFIG.SHEETS.CONTRACTORS, 'Logs has a CREATE_CONTRACTOR entry');

  const contractorsSheet = ss.getSheetByName(APP_CONFIG.SHEETS.CONTRACTORS);
  const found1 = contractorsSheet.getRange(2, 1, Math.max(contractorsSheet.getLastRow() - 1, 0), 5).getValues()
    .some(r => r[0] === 'Sweep Test Contractor');
  assert(found1, 'sheet has the new contractor row');

  const editRes = saveContractor({ contractorName: 'Sweep Test Contractor', remarks: 'edited', originalContractorName: 'Sweep Test Contractor' });
  assert(editRes.success, 'edit succeeds: ' + editRes.message);
  assertLogged('UPDATE_CONTRACTOR', APP_CONFIG.SHEETS.CONTRACTORS, 'Logs has an UPDATE_CONTRACTOR entry for the edit');

  const deleteRes = deleteContractor('Sweep Test Contractor');
  assert(deleteRes.success, 'delete succeeds: ' + deleteRes.message);
  assertLogged('DELETE_CONTRACTOR', APP_CONFIG.SHEETS.CONTRACTORS, 'Logs has a DELETE_CONTRACTOR entry');
}

console.log('\n=== 7. BOM ===');
const BOM_TOKEN = 'sweep-test-token';
fakeCache.put('BOM_AUTH_TOKEN_' + BOM_TOKEN, '1');
{
  const createRes = saveBOM({
    productName: 'Sweep Test Bike',
    components: JSON.stringify([{ itemName: 'Sweep Frame', qtyPerProduct: 1 }])
  }, BOM_TOKEN);
  assert(createRes.success, 'create succeeds: ' + createRes.message);
  const productId = createRes.data.productId;
  assertLogged('CREATE', APP_CONFIG.SHEETS.BOM, 'Logs has a CREATE/BOM entry');

  const bomSheet = ss.getSheetByName(APP_CONFIG.SHEETS.BOM);
  const found1 = bomSheet.getRange(2, BOM_COL.PRODUCT_ID, Math.max(bomSheet.getLastRow() - 1, 0), 3).getValues()
    .some(r => r[0] === productId && r[2] === 'Sweep Frame');
  assert(found1, 'sheet has the new BOM component row');

  const editRes = saveBOM({
    productId, productName: 'Sweep Test Bike',
    components: JSON.stringify([{ itemName: 'Sweep Frame', qtyPerProduct: 2 }])
  }, BOM_TOKEN);
  assert(editRes.success, 'edit succeeds: ' + editRes.message);
  assertLogged('UPDATE', APP_CONFIG.SHEETS.BOM, 'Logs has an UPDATE/BOM entry for the edit');

  const deleteRes = deleteBOM(productId, BOM_TOKEN);
  assert(deleteRes.success, 'delete succeeds: ' + deleteRes.message);
  assertLogged('DELETE', APP_CONFIG.SHEETS.BOM, 'Logs has a DELETE/BOM entry');
}

console.log('\n=== 8. Process ===');
{
  const createRes = saveProcess({
    processName: 'Sweep Test Process', lotPrefix: 'SWP', outputItemName: 'Sweep Output', sequence: 1
  });
  assert(createRes.success, 'create succeeds: ' + createRes.message);
  const processId = createRes.data.processId;
  assertLogged('CREATE', APP_CONFIG.SHEETS.PROCESS_MASTER, 'Logs has a CREATE/Process Master entry');

  const processSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PROCESS_MASTER);
  const found1 = processSheet.getRange(2, PROCESS_COL.PROCESS_ID, Math.max(processSheet.getLastRow() - 1, 0), 2).getValues()
    .some(r => r[0] === processId && r[1] === 'Sweep Test Process');
  assert(found1, 'sheet has the new process row');

  const editRes = saveProcess({
    processId, processName: 'Sweep Test Process (edited)', lotPrefix: 'SWP', outputItemName: 'Sweep Output', sequence: 2
  });
  assert(editRes.success, 'edit succeeds: ' + editRes.message);
  assertLogged('UPDATE', APP_CONFIG.SHEETS.PROCESS_MASTER, 'Logs has an UPDATE/Process Master entry for the edit');

  const deleteRes = deleteProcess(processId);
  assert(deleteRes.success, 'delete succeeds: ' + deleteRes.message);
  assertLogged('DELETE', APP_CONFIG.SHEETS.PROCESS_MASTER, 'Logs has a DELETE/Process Master entry');
}

console.log('\n=== 9. Production ===');
{
  // Dedicated helper process (kept alive independent of section 8's, which
  // was already deleted above) — Production hard-requires an existing
  // Process Master row for its processId.
  const procRes = saveProcess({
    processName: 'Sweep Production Helper Process', lotPrefix: 'SPH', outputItemName: 'Sweep Prod Output', sequence: 1
  });
  assert(procRes.success, 'helper process for Production created: ' + procRes.message);
  const processId = procRes.data.processId;

  const createRes = saveProduction({
    processId, assignedTo: 'Sweep Contractor', qty: 5, date: '2026-01-01', status: 'Pending',
    componentsConsumed: JSON.stringify([{ itemName: 'Sweep Comp', qty: 5, sourceType: 'ITEM', colorGroup: 'COMMON' }])
  });
  assert(createRes.success, 'create succeeds: ' + createRes.message);
  assertLogged('CREATE', APP_CONFIG.SHEETS.PRODUCTION, 'Logs has a CREATE/Production entry');

  const prodSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PRODUCTION);
  const rowIdx = prodSheet.getLastRow();
  assert(prodSheet.getRange(rowIdx, PRODUCTION_COL.PROCESS_ID).getValue() === processId, 'sheet row has correct processId');
  assert(prodSheet.getRange(rowIdx, PRODUCTION_COL.QTY).getValue() === 5, 'sheet row has correct qty (got ' + prodSheet.getRange(rowIdx, PRODUCTION_COL.QTY).getValue() + ')');

  const editRes = saveProduction({
    rowIdx, processId, assignedTo: 'Sweep Contractor', qty: 8, date: '2026-01-01', status: 'Pending', remarks: 'edited',
    componentsConsumed: JSON.stringify([{ itemName: 'Sweep Comp', qty: 8, sourceType: 'ITEM', colorGroup: 'COMMON' }])
  });
  assert(editRes.success, 'edit succeeds: ' + editRes.message);
  const rowAfterEdit = prodSheet.getRange(rowIdx, PRODUCTION_COL.QTY, 1, 1).getValues()[0];
  assert(rowAfterEdit[0] === 8, 'sheet row reflects edited qty (got ' + rowAfterEdit[0] + ')');
  assertLogged('UPDATE', APP_CONFIG.SHEETS.PRODUCTION, 'Logs has an UPDATE/Production entry for the edit');

  const deleteRes = deleteProduction(rowIdx, undefined, undefined);
  assert(deleteRes.success, 'delete succeeds: ' + deleteRes.message);
  assertLogged('DELETE', APP_CONFIG.SHEETS.PRODUCTION, 'Logs has a DELETE/Production entry');
}

console.log('\n=== 10. Issue ===');
{
  const createRes = saveIssueStock({
    issuedTo: 'Sweep Dept', date: '2026-01-01',
    items: JSON.stringify([{ name: 'Sweep Issue Item', qty: 2, unit: 'Pcs' }])
  });
  assert(createRes.success, 'create succeeds: ' + createRes.message);
  const issueId = createRes.data.issueId;
  assertLogged('CREATE', APP_CONFIG.SHEETS.ISSUE, 'Logs has a CREATE/Issue entry');

  const issueSheet = ss.getSheetByName(APP_CONFIG.SHEETS.ISSUE);
  const found1 = issueSheet.getRange(2, ISSUE_COL.ISSUE_ID, Math.max(issueSheet.getLastRow() - 1, 0), 1).getValues()
    .some(r => r[0] === issueId);
  assert(found1, 'sheet has the new issue row');

  console.log('  SKIP: no edit — saveIssueStock is create-only at the API level (confirmed by code inspection, no isEdit branch exists)');

  const deleteRes = deleteIssueBulk([issueId]);
  assert(deleteRes.success, 'delete (bulk, 1 item — no single-row delete exists) succeeds: ' + deleteRes.message);
  const stillThere = issueSheet.getRange(2, ISSUE_COL.ISSUE_ID, Math.max(issueSheet.getLastRow() - 1, 0), 1).getValues()
    .some(r => r[0] === issueId);
  assert(!stillThere, 'sheet row removed after delete');
  assertLogged('BULK_DELETE', APP_CONFIG.SHEETS.ISSUE, 'Logs has a BULK_DELETE/Issue entry');
}

console.log('\n=== 11. Dispatch ===');
{
  // Dedicated final-stage process + a Completed production lot so
  // _computeReadyToDispatchMap() has a real balance to dispatch against.
  const procRes = saveProcess({
    processName: 'Sweep Dispatch Helper Process', lotPrefix: 'SDH', outputItemName: 'Sweep Dispatch Output',
    sequence: 1, isFinalStage: true
  });
  assert(procRes.success, 'helper final-stage process for Dispatch created: ' + procRes.message);
  const processId = procRes.data.processId;

  const prodRes = saveProduction({
    processId, assignedTo: 'Sweep Contractor', qty: 10, date: '2026-01-01', status: 'Completed',
    productId: 'SWEEP-PRD-1', productName: 'Sweep Dispatch Product',
    componentsConsumed: JSON.stringify([{ itemName: 'Sweep Comp', qty: 10, sourceType: 'ITEM', colorGroup: 'COMMON' }])
  });
  assert(prodRes.success, 'helper completed production lot for Dispatch created: ' + prodRes.message);

  const createRes = saveDispatch({
    productId: 'SWEEP-PRD-1', productName: 'Sweep Dispatch Product', qty: 4, dispatchDate: '2026-01-02'
  });
  assert(createRes.success, 'create succeeds: ' + createRes.message);
  const dispatchNumber = createRes.data.dispatchNumber;
  assertLogged('CREATE', APP_CONFIG.SHEETS.DISPATCH, 'Logs has a CREATE/Dispatch entry');

  const dispatchSheet = ss.getSheetByName(APP_CONFIG.SHEETS.DISPATCH);
  const rowIdx = dispatchSheet.getLastRow();
  const row = dispatchSheet.getRange(rowIdx, DISPATCH_COL.PRODUCT_ID, 1, 3).getValues()[0];
  assert(row[0] === 'SWEEP-PRD-1', 'sheet row has correct productId');
  assert(row[2] === 4, 'sheet row has correct qty (got ' + row[2] + ')');

  const editRes = saveDispatch({
    rowIdx, productId: 'SWEEP-PRD-1', productName: 'Sweep Dispatch Product', qty: 6,
    dispatchDate: '2026-01-02', remarks: 'edited'
  });
  assert(editRes.success, 'edit succeeds: ' + editRes.message);
  const rowAfterEdit = dispatchSheet.getRange(rowIdx, DISPATCH_COL.QTY, 1, 1).getValues()[0];
  assert(rowAfterEdit[0] === 6, 'sheet row reflects edited qty (got ' + rowAfterEdit[0] + ')');
  assertLogged('UPDATE', APP_CONFIG.SHEETS.DISPATCH, 'Logs has an UPDATE/Dispatch entry for the edit');

  const deleteRes = deleteDispatch(rowIdx, undefined, undefined);
  assert(deleteRes.success, 'delete succeeds: ' + deleteRes.message);
  assertLogged('DELETE', APP_CONFIG.SHEETS.DISPATCH, 'Logs has a DELETE/Dispatch entry');
}

console.log('\n=== 12. Return ===');
{
  const createRes = saveReturn({
    returnDate: '2026-01-01', vendor: 'Sweep Vendor',
    items: JSON.stringify([{ name: 'Sweep Return Item', qty: 1, unit: 'Pcs', price: 20, reason: 'Defective' }])
  });
  assert(createRes.success, 'create succeeds: ' + createRes.message);
  const returnNumber = createRes.data.returnNumber;
  assertLogged('CREATE', APP_CONFIG.SHEETS.RETURN, 'Logs has a CREATE/Return entry');

  const returnSheet = ss.getSheetByName(APP_CONFIG.SHEETS.RETURN);
  const found1 = returnSheet.getRange(2, RETURN_COL.RETURN_NUMBER, Math.max(returnSheet.getLastRow() - 1, 0), 1).getValues()
    .some(r => r[0] === returnNumber);
  assert(found1, 'sheet has the new return row');

  const editRes = saveReturn({
    existingReturnNumber: returnNumber, returnNumber, returnDate: '2026-01-02', vendor: 'Sweep Vendor',
    items: JSON.stringify([{ name: 'Sweep Return Item', qty: 3, unit: 'Pcs', price: 20, reason: 'Defective - edited' }])
  });
  assert(editRes.success, 'edit succeeds: ' + editRes.message);
  assertLogged('UPDATE', APP_CONFIG.SHEETS.RETURN, 'Logs has an UPDATE/Return entry for the edit');

  const deleteRes = deleteReturn(returnNumber);
  assert(deleteRes.success, 'delete succeeds: ' + deleteRes.message);
  assertLogged('DELETE', APP_CONFIG.SHEETS.RETURN, 'Logs has a DELETE/Return entry');
}

console.log('\n=== 13. Wastage ===');
{
  const createRes = saveWastage({
    date: '2026-01-01',
    items: JSON.stringify([{ name: 'Sweep Wastage Item', qty: 1, unit: 'Pcs', reason: 'Damaged' }])
  });
  assert(createRes.success, 'create succeeds: ' + createRes.message);
  const wastageId = createRes.data.wastageId;
  assertLogged('CREATE', APP_CONFIG.SHEETS.WASTAGE, 'Logs has a CREATE/Wastage entry');

  const wastageSheet = ss.getSheetByName(APP_CONFIG.SHEETS.WASTAGE);
  const found1 = wastageSheet.getRange(2, WASTAGE_COL.WASTAGE_ID, Math.max(wastageSheet.getLastRow() - 1, 0), 1).getValues()
    .some(r => r[0] === wastageId);
  assert(found1, 'sheet has the new wastage row');

  console.log('  SKIP: no edit — saveWastage is create-only at the API level (confirmed by code inspection, no isEdit branch exists)');

  const deleteRes = deleteWastageBulk([wastageId]);
  assert(deleteRes.success, 'delete (bulk, 1 item — no single-row delete exists) succeeds: ' + deleteRes.message);
  const stillThere = wastageSheet.getRange(2, WASTAGE_COL.WASTAGE_ID, Math.max(wastageSheet.getLastRow() - 1, 0), 1).getValues()
    .some(r => r[0] === wastageId);
  assert(!stillThere, 'sheet row removed after delete');
  assertLogged('BULK_DELETE', APP_CONFIG.SHEETS.WASTAGE, 'Logs has a BULK_DELETE/Wastage entry');
}

// ─────────────────────────────────────────────────────────────────────────
console.log('\n' + (failures === 0 ? 'ALL TESTS PASSED' : failures + ' TEST(S) FAILED'));
process.exit(failures === 0 ? 0 : 1);
