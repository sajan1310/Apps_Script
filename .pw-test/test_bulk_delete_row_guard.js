/**
 * Standalone Node harness (same mock-SpreadsheetApp pattern as
 * test_production_negative_qty.js) exercising the Phase 4 delete-safety
 * audit fix: deleteProductionBulk / deleteDispatchBulk / deleteWarehousePoolOpening
 * previously trusted a raw client-supplied row index with NO content check —
 * unlike their single-record siblings (deleteProduction, deleteDispatch),
 * which already verify an expected productId/dispatchNumber + qty before
 * deleting. A stale client selection (row shifted by a concurrent edit)
 * could silently delete the wrong row.
 *
 * Covers: a mismatched guard causes the row to be SKIPPED (not deleted, and
 * for the bulk functions, other correctly-matching rows in the same batch
 * still go through); a matching guard deletes normally; omitting the guard
 * entirely still deletes (backward compatible with any caller that doesn't
 * supply expected values).
 *
 * Run: node .pw-test/test_bulk_delete_row_guard.js
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
  appendRow(arr) { const r = this.getLastRow() + 1; arr.forEach((v, i) => this._set(r, i + 1, v)); }
  deleteRow(r) { this.rows.splice(r - 1, 1); }
  deleteRows(r, n) { this.rows.splice(r - 1, n); }
  insertRows(r, n) { for (let i = 0; i < n; i++) this.rows.splice(r - 1, 0, []); }
  insertColumnsAfter(afterPosition, howMany) {
    this.rows.forEach(row => { const blanks = new Array(howMany).fill(''); row.splice(afterPosition, 0, ...blanks); });
  }
}

class FakeSpreadsheet {
  constructor() { this.sheets = {}; }
  getSheetByName(name) { return this.sheets[name] || null; }
  addSheet(name) { const s = new FakeSheet(name); this.sheets[name] = s; return s; }
  insertSheet(name) { return this.addSheet(name); }
}

const ss = new FakeSpreadsheet();

const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush: () => {} },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {} }) },
  console,
  Logger: { log: () => {} },
  Utilities: { getUuid: () => 'uuid-' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);

// module_po.js is loaded for _rewriteWithoutMatchingRowsBulk, which lives
// there but is called by deleteDispatchBulk — Apps Script shares one global
// scope across files, so that resolves in production; this sandbox has to be
// told. Without it deleteDispatchBulk threw ReferenceError and every
// assertion below it failed for a reason that had nothing to do with the guard
// being tested.
['config.js', 'utils.js', 'module_po.js', 'module_process.js', 'module_production.js', 'module_dispatch.js', 'module_warehouse.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});

vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.PRODUCTION_COL = PRODUCTION_COL;
  global.DISPATCH_COL = DISPATCH_COL;
  global.WAREHOUSE_POOL_OPENING_COL = WAREHOUSE_POOL_OPENING_COL;
`, ctx, { filename: 'expose.js' });

const {
  APP_CONFIG, PRODUCTION_COL, DISPATCH_COL, WAREHOUSE_POOL_OPENING_COL,
  deleteProductionBulk, deleteDispatchBulk, deleteWarehousePoolOpening
} = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); }
}

// Seed two Production rows directly (bypassing saveProduction — this test
// targets the delete guard only, not the save pipeline).
const prodSheet = ss.addSheet(APP_CONFIG.SHEETS.PRODUCTION);
prodSheet._set(1, 1, 'Date'); // header row placeholder
const prodRow = (r, productId, qty, status) => {
  prodSheet._set(r, PRODUCTION_COL.DATE, '01/01/2026');
  prodSheet._set(r, PRODUCTION_COL.PRODUCT_ID, productId);
  prodSheet._set(r, PRODUCTION_COL.QTY, qty);
  prodSheet._set(r, PRODUCTION_COL.ASSIGNED_TO, 'Tester');
  prodSheet._set(r, PRODUCTION_COL.STATUS, status);
  prodSheet._set(r, PRODUCTION_COL.PROCESS_ID, 'PRC-1');
  prodSheet._set(r, PRODUCTION_COL.LOT_NUMBER, 'LOT-' + r);
  prodSheet._set(r, PRODUCTION_COL.OUTPUT_ITEM_NAME, 'Widget');
  prodSheet._set(r, PRODUCTION_COL.COMPONENTS_CONSUMED, '[]');
};
prodRow(2, 'PRD-100', 10, 'Completed');
prodRow(3, 'PRD-200', 20, 'Completed');

console.log('\n=== deleteProductionBulk: mismatched guard skips that row, matching row still deletes ===');
{
  const res = deleteProductionBulk([2, 3], [
    { rowIdx: 2, expectedProductId: 'PRD-100', expectedQty: 999 }, // qty mismatch -> skip
    { rowIdx: 3, expectedProductId: 'PRD-200', expectedQty: 20 }   // matches -> delete
  ]);
  assert(res.success, 'call succeeds even with one mismatch: ' + res.message);
  assert(/Deleted 1/.test(res.message), `message reports exactly 1 deleted (got "${res.message}")`);
  assert(/Skipped 1/.test(res.message), `message reports exactly 1 skipped (got "${res.message}")`);
  assert(prodSheet.getLastRow() === 2, `sheet still has the mismatched row (last row now ${prodSheet.getLastRow()})`);
  assert(String(prodSheet._get(2, PRODUCTION_COL.PRODUCT_ID)) === 'PRD-100', 'the mismatched row (PRD-100) is still present');
}

console.log('\n=== deleteProductionBulk: all rows mismatched -> nothing deleted, clear error ===');
{
  const res = deleteProductionBulk([2], [{ rowIdx: 2, expectedProductId: 'WRONG-ID', expectedQty: 1 }]);
  assert(res.success === false, 'call reports failure when every row mismatches: ' + res.message);
  assert(/mismatch/i.test(res.message || ''), `message mentions the mismatch (got "${res.message}")`);
  assert(prodSheet.getLastRow() === 2, 'no row was deleted');
}

console.log('\n=== deleteProductionBulk: no expectedRows supplied -> old behavior preserved (deletes unconditionally) ===');
{
  const res = deleteProductionBulk([2]);
  assert(res.success, 'call succeeds with no guard supplied: ' + res.message);
  assert(/Deleted 1/.test(res.message), `deletes the row with no guard (got "${res.message}")`);
  assert(prodSheet.getLastRow() === 1, 'sheet is back to just the header row (last data row gone)');
}

// Seed two Dispatch rows.
const dispSheet = ss.addSheet(APP_CONFIG.SHEETS.DISPATCH);
const dispRow = (r, dispatchNumber, qty) => {
  dispSheet._set(r, DISPATCH_COL.DISPATCH_NUMBER, dispatchNumber);
  dispSheet._set(r, DISPATCH_COL.DISPATCH_DATE, '01/01/2026');
  dispSheet._set(r, DISPATCH_COL.CLIENT_NAME, 'Test Client');
  dispSheet._set(r, DISPATCH_COL.PRODUCT_ID, 'PRD-1');
  dispSheet._set(r, DISPATCH_COL.QTY, qty);
};
dispRow(2, 'DSP-100', 5);
dispRow(3, 'DSP-200', 8);

// deleteDispatchBulk selects by Dispatch NUMBER, not row index (a bill spans
// however many line rows it has), so its guard is per-bill item count + total
// qty — the same pair deleteDispatch checks for a single bill — rather than the
// productId+qty pair deleteProductionBulk uses for its one-row-per-lot sheet.
console.log('\n=== deleteDispatchBulk: mismatched guard skips that bill, matching bill still deletes ===');
{
  const res = deleteDispatchBulk(['DSP-100', 'DSP-200'], [
    { dispatchNumber: 'DSP-100', expectedItemCount: 1, expectedTotalQty: 999 }, // mismatch -> skip
    { dispatchNumber: 'DSP-200', expectedItemCount: 1, expectedTotalQty: 8 }    // matches -> delete
  ]);
  assert(res.success, 'call succeeds even with one mismatch: ' + res.message);
  assert(/Deleted 1/.test(res.message), `message reports exactly 1 deleted (got "${res.message}")`);
  assert(/Skipped 1/.test(res.message), `message reports exactly 1 skipped (got "${res.message}")`);
  assert(dispSheet.getLastRow() === 2, `sheet still has the mismatched row (last row now ${dispSheet.getLastRow()})`);
  assert(String(dispSheet._get(2, DISPATCH_COL.DISPATCH_NUMBER)) === 'DSP-100', 'the mismatched row (DSP-100) is still present');
}

console.log('\n=== deleteDispatchBulk: no guard supplied -> old unconditional behavior ===');
{
  const res = deleteDispatchBulk(['DSP-100']);
  assert(res.success, 'call succeeds with no guard supplied: ' + res.message);
  assert(/Deleted 1/.test(res.message), `deletes the bill with no guard (got "${res.message}")`);
  // No header row was seeded on this sheet (unlike prodSheet above), so an
  // emptied Dispatch sheet reports last row 0, not 1.
  assert(dispSheet.getLastRow() === 0, `sheet has no data rows left (got ${dispSheet.getLastRow()})`);
}

console.log('\n=== deleteDispatchBulk: bill numbers match case-insensitively ===');
{
  dispRow(2, 'DSP-300', 4);
  const res = deleteDispatchBulk(['dsp-300']);
  assert(res.success, 'lowercased bill number still resolves: ' + res.message);
  assert(/Deleted 1 dispatch bill\(s\) \(1 item/.test(res.message),
    `reports 1 bill and 1 item actually removed (got "${res.message}")`);
  assert(dispSheet.getLastRow() === 0, `the bill is gone (got last row ${dispSheet.getLastRow()})`);
}

console.log('\n=== deleteDispatchBulk: a selection naming bills that no longer exist reports honestly ===');
{
  const res = deleteDispatchBulk(['DSP-GONE-1', 'DSP-GONE-2']);
  assert(!res.success, `reports failure rather than "deleted 2" (got success=${res.success})`);
  assert(/were found|not found/i.test(res.message), `explains nothing matched (got "${res.message}")`);
}

// Seed two Warehouse Pool Opening rows.
const openingSheet = ss.addSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
const openingRow = (r, outputItemName, qty) => {
  openingSheet._set(r, WAREHOUSE_POOL_OPENING_COL.OUTPUT_ITEM_NAME, outputItemName);
  openingSheet._set(r, WAREHOUSE_POOL_OPENING_COL.PROCESS_ID, 'PRC-1');
  openingSheet._set(r, WAREHOUSE_POOL_OPENING_COL.QTY, qty);
  openingSheet._set(r, WAREHOUSE_POOL_OPENING_COL.DATE, '01/01/2026');
};
openingRow(2, 'Painted Frame', 15);

console.log('\n=== deleteWarehousePoolOpening: mismatched guard rejects the whole call (single-row op) ===');
{
  const res = deleteWarehousePoolOpening(2, 'Painted Frame', 999);
  assert(res.success === false, 'call reports failure on a qty mismatch: ' + res.message);
  assert(/mismatch/i.test(res.message || ''), `message mentions the mismatch (got "${res.message}")`);
  assert(openingSheet.getLastRow() === 2, 'the entry is still present');
}

console.log('\n=== deleteWarehousePoolOpening: matching guard deletes normally ===');
{
  const res = deleteWarehousePoolOpening(2, 'Painted Frame', 15);
  assert(res.success, 'call succeeds when the guard matches: ' + res.message);
  assert(openingSheet.getLastRow() === 0, 'the entry is now gone');
}

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
