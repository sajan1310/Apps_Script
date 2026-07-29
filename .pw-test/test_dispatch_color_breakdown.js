/**
 * Verifies getReadyToDispatchData() now includes a per-color colorBreakdown
 * alongside the existing color-blind aggregate: a Product Tag credited by
 * 2 different Completed lots under 2 different color combos should show
 * ONE aggregate row (unchanged row count/behavior) plus a colorBreakdown
 * array with each color's own produced/dispatched/ready numbers.
 *
 * Run: node .pw-test/test_dispatch_color_breakdown.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = 'c:\\Users\\erkar\\my-app-script-project';

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) { this.sheet = sheet; this.row = row; this.col = col; this.numRows = numRows; this.numCols = numCols; }
  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) { const rowArr = []; for (let c = 0; c < this.numCols; c++) rowArr.push(this.sheet._get(this.row + r, this.col + c)); out.push(rowArr); }
    return out;
  }
  getValue() { return this.sheet._get(this.row, this.col); }
  setValues(values) { values.forEach((rowArr, r) => rowArr.forEach((val, c) => this.sheet._set(this.row + r, this.col + c, val))); return this; }
  setValue(v) { this.sheet._set(this.row, this.col, v); return this; }
  clearContent() { for (let r = 0; r < this.numRows; r++) for (let c = 0; c < this.numCols; c++) this.sheet._set(this.row + r, this.col + c, ''); return this; }
  setFontWeight() { return this; } setBackground() { return this; } setNumberFormat() { return this; }
}
class FakeSheet {
  constructor(name) { this.name = name; this.rows = []; }
  getName() { return this.name; }
  _ensureRow(r) { while (this.rows.length < r) this.rows.push([]); }
  _get(r, c) { this._ensureRow(r); const row = this.rows[r - 1]; return row[c - 1] === undefined ? '' : row[c - 1]; }
  _set(r, c, v) { this._ensureRow(r); const row = this.rows[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; }
  getLastRow() { for (let r = this.rows.length; r >= 1; r--) { if (this.rows[r - 1].some(v => v !== '' && v !== undefined && v !== null)) return r; } return 0; }
  getLastColumn() { let max = 0; this.rows.forEach(row => { for (let c = row.length; c >= 1; c--) { if (row[c - 1] !== '' && row[c - 1] !== undefined && row[c - 1] !== null) { max = Math.max(max, c); break; } } }); return max; }
  getRange(row, col, numRows = 1, numCols = 1) { return new FakeRange(this, row, col, numRows, numCols); }
  appendRow(arr) { const r = this.getLastRow() + 1; arr.forEach((v, i) => this._set(r, i + 1, v)); }
  deleteRow(r) { this.rows.splice(r - 1, 1); } deleteRows(r, n) { this.rows.splice(r - 1, n); }
  insertRows(r, n) { for (let i = 0; i < n; i++) this.rows.splice(r - 1, 0, []); }
  insertColumnsAfter(afterPosition, howMany) { this.rows.forEach(row => { const blanks = new Array(howMany).fill(''); row.splice(afterPosition, 0, ...blanks); }); }
}
class FakeSpreadsheet { constructor() { this.sheets = {}; } getSheetByName(name) { return this.sheets[name] || null; } addSheet(name) { const s = new FakeSheet(name); this.sheets[name] = s; return s; } insertSheet(name) { return this.addSheet(name); } }
const ss = new FakeSpreadsheet();
const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush: () => {} },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
  console, Logger: { log: () => {} },
  Utilities: { getUuid: () => 'uuid-' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);
['config.js', 'utils.js', 'module_process.js', 'module_production.js', 'module_warehouse.js', 'module_bom.js', 'module_po.js', 'module_dispatch.js', 'module_clients.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});
vm.runInContext('global.APP_CONFIG=APP_CONFIG; global.PRODUCTION_COL=PRODUCTION_COL; global.PROCESS_COL=PROCESS_COL;', ctx, { filename: 'expose.js' });
const { APP_CONFIG, PRODUCTION_COL, PROCESS_COL, recalculateWarehousePool, getReadyToDispatchData } = ctx;

let failures = 0;
function assert(cond, msg) { if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); } }

const processSheet = ss.addSheet(APP_CONFIG.SHEETS.PROCESS_MASTER);
processSheet._set(2, PROCESS_COL.PROCESS_ID, 'PRC-PACK');
processSheet._set(2, PROCESS_COL.PROCESS_NAME, 'Packing');
processSheet._set(2, PROCESS_COL.SEQUENCE, 6);
processSheet._set(2, PROCESS_COL.IS_FINAL_STAGE, true);
processSheet._set(2, PROCESS_COL.ACTIVE, true);
processSheet._set(2, PROCESS_COL.OUTPUT_ITEM_NAME, 'Packed Bicycle 16 inch');

const prodSheet = ss.addSheet(APP_CONFIG.SHEETS.PRODUCTION);
function setLot(row, { qty, lotNumber, productTag, color, colorBreakdown }) {
  prodSheet._set(row, PRODUCTION_COL.DATE, '01/01/2026');
  prodSheet._set(row, PRODUCTION_COL.QTY, qty);
  prodSheet._set(row, PRODUCTION_COL.STATUS, 'Completed');
  prodSheet._set(row, PRODUCTION_COL.PROCESS_ID, 'PRC-PACK');
  prodSheet._set(row, PRODUCTION_COL.LOT_NUMBER, lotNumber);
  prodSheet._set(row, PRODUCTION_COL.OUTPUT_ITEM_NAME, 'Packed Bicycle 16 inch');
  prodSheet._set(row, PRODUCTION_COL.PRODUCT_ID, productTag);
  prodSheet._set(row, PRODUCTION_COL.COLOR, color);
  prodSheet._set(row, PRODUCTION_COL.COLOR_BREAKDOWN, JSON.stringify(colorBreakdown));
  prodSheet._set(row, PRODUCTION_COL.COMPONENTS_CONSUMED, '[]');
}

// Same Product Tag "PROD-001", 2 different batches with different color combos.
setLot(2, {
  qty: 15, lotNumber: 'PKG-0001', productTag: 'PROD-001',
  color: 'Blue-White / BCP', colorBreakdown: [
    { color: 'Blue-White', qty: 15, countsTowardTotal: true, axisKey: 'pool:frame' },
    { color: 'BCP', qty: 15, countsTowardTotal: false, axisKey: 'pool:rim' }
  ]
});
setLot(3, {
  qty: 10, lotNumber: 'PKG-0002', productTag: 'PROD-001',
  color: 'Red-White / Black', colorBreakdown: [
    { color: 'Red-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:frame' },
    { color: 'Black', qty: 10, countsTowardTotal: false, axisKey: 'pool:rim' }
  ]
});

recalculateWarehousePool();

const resp = getReadyToDispatchData();
assert(resp.success, 'getReadyToDispatchData succeeds: ' + resp.message);
const records = resp.data;
console.log(JSON.stringify(records, null, 2));

assert(records.length === 1, `exactly 1 aggregate row for PROD-001 (unchanged row count) (got ${records.length})`);
const rec = records[0];
assert(rec.productId === 'PROD-001', `productId is PROD-001 (got ${rec.productId})`);
assert(rec.producedQty === 25, `aggregate producedQty is 25 (15+10) (got ${rec.producedQty})`);
assert(rec.readyQty === 25, `aggregate readyQty is 25 (got ${rec.readyQty})`);

assert(Array.isArray(rec.colorBreakdown), 'colorBreakdown array exists');
assert(rec.colorBreakdown.length === 2, `colorBreakdown has 2 entries, one per color combo (got ${rec.colorBreakdown.length})`);

const blueWhiteBcp = rec.colorBreakdown.find(c => c.color === 'Blue-White / BCP');
const redWhiteBlack = rec.colorBreakdown.find(c => c.color === 'Red-White / Black');
assert(!!blueWhiteBcp && blueWhiteBcp.producedQty === 15 && blueWhiteBcp.readyQty === 15,
  `"Blue-White / BCP" breakdown entry shows produced=15, ready=15 (got ${JSON.stringify(blueWhiteBcp)})`);
assert(!!redWhiteBlack && redWhiteBlack.producedQty === 10 && redWhiteBlack.readyQty === 10,
  `"Red-White / Black" breakdown entry shows produced=10, ready=10 (got ${JSON.stringify(redWhiteBlack)})`);

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
