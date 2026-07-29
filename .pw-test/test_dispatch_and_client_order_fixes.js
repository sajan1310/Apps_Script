/**
 * Standalone Node harness (same mock-SpreadsheetApp pattern as
 * test_bulk_delete_row_guard.js) covering two more of the 9 architectural
 * gaps verified+fixed on 2026-07-13 (see
 * verification_2026_07_13_architectural_gaps in project memory):
 *
 * Fix #5 — saveDispatch now also validates against the SPECIFIC PI/Estimate
 * line's own remaining qty (not just aggregate product stock), so one PI
 * can no longer be over-fulfilled using stock meant for a different order
 * of the same product.
 *
 * Fix #6 — deleteClientOrder/deleteClientOrdersBulk now also check
 * Production (via the new structured PRODUCTION_COL.ORDER_NUMBER column,
 * see _pushOrderLinesToProduction), not just Dispatch, before allowing a
 * PI/Estimate to be deleted.
 *
 * Run: node .pw-test/test_dispatch_and_client_order_fixes.js
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

['config.js', 'utils.js', 'module_process.js', 'module_production.js', 'module_warehouse.js', 'module_bom.js', 'module_po.js', 'module_dispatch.js', 'module_clients.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});

vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.CLIENT_ORDERS_COL = CLIENT_ORDERS_COL;
  global.WAREHOUSE_POOL_OPENING_COL = WAREHOUSE_POOL_OPENING_COL;
  global.DISPATCH_COL = DISPATCH_COL;
  global.PRODUCTION_COL = PRODUCTION_COL;
`, ctx, { filename: 'expose.js' });

const {
  APP_CONFIG, CLIENT_ORDERS_COL, WAREHOUSE_POOL_OPENING_COL, DISPATCH_COL, PRODUCTION_COL,
  saveDispatch, deleteClientOrder, deleteClientOrdersBulk, recalculateWarehousePool
} = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); }
}

console.log('\n=== Fix #5: Dispatch qty is validated against the specific PI line, not just aggregate stock ===');
{
  const ordersSheet = ss.addSheet(APP_CONFIG.SHEETS.CLIENT_ORDERS);
  ordersSheet._set(2, CLIENT_ORDERS_COL.ORDER_NUMBER, 'ORD-1');
  ordersSheet._set(2, CLIENT_ORDERS_COL.CLIENT_NAME, 'Test Client');
  ordersSheet._set(2, CLIENT_ORDERS_COL.PRODUCT_ID, 'PRD-1');
  ordersSheet._set(2, CLIENT_ORDERS_COL.QTY_ORDERED, 10);
  ordersSheet._set(2, CLIENT_ORDERS_COL.STATUS, 'Order Confirmed');

  // Plenty of AGGREGATE ready-to-dispatch stock (20) for PRD-1 -- the old
  // check alone would happily allow way more than this one order needs.
  // Seeded via Warehouse Pool Opening (the persistent source recalculateWarehousePool
  // folds in on every rebuild — saveDispatch triggers exactly that rebuild
  // internally, so seeding the derived WAREHOUSE_POOL sheet directly would
  // just get wiped on the first call) rather than the derived sheet itself.
  const openingSheet = ss.addSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
  openingSheet._set(2, WAREHOUSE_POOL_OPENING_COL.OUTPUT_ITEM_NAME, 'Test Bike');
  openingSheet._set(2, WAREHOUSE_POOL_OPENING_COL.PRODUCT_TAG, 'PRD-1');
  openingSheet._set(2, WAREHOUSE_POOL_OPENING_COL.QTY, 20);
  openingSheet._set(2, WAREHOUSE_POOL_OPENING_COL.DATE, '01/01/2026');

  // _computeReadyToDispatchMap only surfaces a tagged bucket whose
  // Process ID belongs to a known final-stage process (or is blank/legacy
  // -- see that function's own comment) -- seed one so the Opening row above
  // is actually visible to Dispatch.
  const procSheet = ss.addSheet(APP_CONFIG.SHEETS.PROCESS_MASTER);
  procSheet._set(2, 1, 'PRC-FINAL');   // PROCESS_ID
  procSheet._set(2, 2, 'Final Packing'); // PROCESS_NAME
  procSheet._set(2, 3, 1);             // SEQUENCE
  procSheet._set(2, 4, 'PK');          // LOT_PREFIX
  procSheet._set(2, 5, true);          // IS_FINAL_STAGE
  procSheet._set(2, 6, true);          // ACTIVE
  procSheet._set(2, 8, 'Test Bike');   // OUTPUT_ITEM_NAME
  openingSheet._set(2, WAREHOUSE_POOL_OPENING_COL.PROCESS_ID, 'PRC-FINAL');

  // Materialize the derived WAREHOUSE_POOL sheet from the Opening seed above
  // -- saveDispatch's own availability check reads that derived sheet
  // (getWarehousePoolData), not Opening directly, and only saveDispatch's
  // own POST-write recalc would otherwise build it for the first time.
  recalculateWarehousePool();

  // saveDispatch now takes a bill's item lines as a JSON array (see
  // feature_dispatch_record_new_bill_multiitem in project memory) instead
  // of a single top-level productId/qty — each call below is a one-item bill.
  const oneLine = (productId, qty) => JSON.stringify([{ productId, productName: 'Test Bike', qty }]);

  const res1 = saveDispatch({ lines: oneLine('PRD-1', 6), orderNumber: 'ORD-1' });
  assert(res1.success, 'first dispatch of 6 against ORD-1 (10 ordered) succeeds: ' + res1.message);

  const res2 = saveDispatch({ lines: oneLine('PRD-1', 6), orderNumber: 'ORD-1' });
  assert(res2.success === false, `a second dispatch of 6 (12 total > 10 ordered on ORD-1) is rejected even though 20 units of aggregate stock remain: ${res2.message}`);
  assert(/only 4/i.test(res2.message || ''), `message reports exactly 4 remaining on the order (10 - 6) (got "${res2.message}")`);

  const res3 = saveDispatch({ lines: oneLine('PRD-1', 6), orderNumber: '' });
  assert(res3.success, `a Direct dispatch (no order reference) for the same product is unaffected by ORD-1's own limit: ${res3.message}`);

  const res4 = saveDispatch({ lines: oneLine('PRD-1', 4), orderNumber: 'ORD-1' });
  assert(res4.success, `dispatching exactly the remaining 4 against ORD-1 succeeds: ${res4.message}`);
}

console.log('\n=== Fix #6: deleteClientOrder / deleteClientOrdersBulk also check Production, not just Dispatch ===');
{
  const ordersSheet = ss.getSheetByName(APP_CONFIG.SHEETS.CLIENT_ORDERS);
  const dispatchSheet = ss.getSheetByName(APP_CONFIG.SHEETS.DISPATCH); // created by saveDispatch above

  // ORD-2: no Dispatch, no Production reference -- deletable.
  const row2 = ordersSheet.getLastRow() + 1;
  ordersSheet._set(row2, CLIENT_ORDERS_COL.ORDER_NUMBER, 'ORD-2');
  ordersSheet._set(row2, CLIENT_ORDERS_COL.PRODUCT_ID, 'PRD-2');
  ordersSheet._set(row2, CLIENT_ORDERS_COL.QTY_ORDERED, 5);

  const del2 = deleteClientOrder('ORD-2');
  assert(del2.success, 'ORD-2 (no references anywhere) deletes cleanly: ' + del2.message);

  // ORD-3: referenced by a Production lot's structured ORDER_NUMBER column.
  const row3 = ordersSheet.getLastRow() + 1;
  ordersSheet._set(row3, CLIENT_ORDERS_COL.ORDER_NUMBER, 'ORD-3');
  ordersSheet._set(row3, CLIENT_ORDERS_COL.PRODUCT_ID, 'PRD-3');
  ordersSheet._set(row3, CLIENT_ORDERS_COL.QTY_ORDERED, 5);

  const prodSheet = ss.addSheet(APP_CONFIG.SHEETS.PRODUCTION);
  prodSheet._set(2, PRODUCTION_COL.DATE, '01/01/2026');
  prodSheet._set(2, PRODUCTION_COL.PRODUCT_ID, 'PRD-3');
  prodSheet._set(2, PRODUCTION_COL.QTY, 5);
  prodSheet._set(2, PRODUCTION_COL.STATUS, 'Pending');
  prodSheet._set(2, PRODUCTION_COL.PROCESS_ID, 'PRC-1');
  prodSheet._set(2, PRODUCTION_COL.LOT_NUMBER, 'LOT-1');
  prodSheet._set(2, PRODUCTION_COL.ORDER_NUMBER, 'ORD-3');

  const del3 = deleteClientOrder('ORD-3');
  assert(del3.success === false, 'ORD-3 (queued into Production) is blocked from deletion: ' + del3.message);
  assert(/production lot/i.test(del3.message || ''), `rejection message explains it's the Production reference (got "${del3.message}")`);

  // Bulk: ORD-4 clean, ORD-5 referenced by Dispatch (the pre-existing
  // guard), ORD-6 referenced by Production (the new guard) -- one call,
  // three different outcomes.
  const row4 = ordersSheet.getLastRow() + 1;
  ordersSheet._set(row4, CLIENT_ORDERS_COL.ORDER_NUMBER, 'ORD-4');
  ordersSheet._set(row4, CLIENT_ORDERS_COL.PRODUCT_ID, 'PRD-4');
  ordersSheet._set(row4, CLIENT_ORDERS_COL.QTY_ORDERED, 1);

  const row5 = ordersSheet.getLastRow() + 1;
  ordersSheet._set(row5, CLIENT_ORDERS_COL.ORDER_NUMBER, 'ORD-5');
  ordersSheet._set(row5, CLIENT_ORDERS_COL.PRODUCT_ID, 'PRD-5');
  ordersSheet._set(row5, CLIENT_ORDERS_COL.QTY_ORDERED, 1);
  const dRow = dispatchSheet.getLastRow() + 1;
  dispatchSheet._set(dRow, DISPATCH_COL.DISPATCH_NUMBER, 'DSP-99');
  dispatchSheet._set(dRow, DISPATCH_COL.ORDER_NUMBER, 'ORD-5');
  dispatchSheet._set(dRow, DISPATCH_COL.PRODUCT_ID, 'PRD-5');
  dispatchSheet._set(dRow, DISPATCH_COL.QTY, 1);

  const row6 = ordersSheet.getLastRow() + 1;
  ordersSheet._set(row6, CLIENT_ORDERS_COL.ORDER_NUMBER, 'ORD-6');
  ordersSheet._set(row6, CLIENT_ORDERS_COL.PRODUCT_ID, 'PRD-6');
  ordersSheet._set(row6, CLIENT_ORDERS_COL.QTY_ORDERED, 1);
  const pRow = prodSheet.getLastRow() + 1;
  prodSheet._set(pRow, PRODUCTION_COL.DATE, '01/01/2026');
  prodSheet._set(pRow, PRODUCTION_COL.PRODUCT_ID, 'PRD-6');
  prodSheet._set(pRow, PRODUCTION_COL.QTY, 1);
  prodSheet._set(pRow, PRODUCTION_COL.STATUS, 'Pending');
  prodSheet._set(pRow, PRODUCTION_COL.LOT_NUMBER, 'LOT-2');
  prodSheet._set(pRow, PRODUCTION_COL.ORDER_NUMBER, 'ORD-6');

  const bulkRes = deleteClientOrdersBulk(['ORD-4', 'ORD-5', 'ORD-6']);
  assert(bulkRes.data.deletedIds.map(String).includes('ORD-4'), `ORD-4 (clean) was deleted (got deletedIds=${JSON.stringify(bulkRes.data.deletedIds)})`);
  assert(bulkRes.data.skipped.includes('ORD-5'), `ORD-5 (Dispatch reference) was skipped (got skipped=${JSON.stringify(bulkRes.data.skipped)})`);
  assert(bulkRes.data.skipped.includes('ORD-6'), `ORD-6 (Production reference, the new guard) was skipped (got skipped=${JSON.stringify(bulkRes.data.skipped)})`);
}

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
