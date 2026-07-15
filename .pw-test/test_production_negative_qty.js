/**
 * Standalone Node harness (same mock-SpreadsheetApp pattern as
 * test_color_axes.js) exercising saveProduction()'s relaxed quantity
 * validation: zero AND negative quantities are both allowed, for the
 * produced qty itself and for each component consumed (ITEM or POOL
 * sourced) — a lot can be logged as a correction/reversal, or an explicit
 * zero-output record, without editing history, mirroring
 * adjustStockManually/adjustWarehousePoolManually's existing "negative is
 * allowed, only manual review needed later" behavior. Covers both the flat
 * (no color-specific components) and per-color-breakdown save paths.
 *
 * Run: node .pw-test/test_production_negative_qty.js
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
  console,
  Logger: { log: () => {} },
  Utilities: { getUuid: () => 'uuid-' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);

['config.js', 'utils.js', 'module_process.js', 'module_production.js', 'module_warehouse.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});

vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
`, ctx, { filename: 'expose.js' });

const { APP_CONFIG, saveProcess, saveProduction, getProductionData } = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); }
}

console.log('\n=== Setup: plain process, no color-specific components (flat qty path) ===');
let flatProcessId;
{
  const res = saveProcess({
    processName: 'Plain Assembly',
    sequence: 1,
    lotPrefix: 'PA',
    outputItemName: 'Plain Widget',
    isFinalStage: false,
    active: true,
    remarks: '',
    components: JSON.stringify([
      { itemName: 'Assembly Screws', sourceType: 'ITEM', qtyPerUnit: 4, colorGroup: 'COMMON' }
    ])
  });
  assert(res.success, 'saveProcess succeeds: ' + res.message);
  flatProcessId = res.data && res.data.processId;
  assert(!!flatProcessId, 'processId returned');
}

console.log('\n=== Test 1: flat path accepts an exact-zero quantity (explicit zero-output record) ===');
{
  const res = saveProduction({
    processId: flatProcessId,
    assignedTo: 'Test Contractor',
    status: 'Pending',
    qty: 0,
    componentsConsumed: JSON.stringify([
      { itemName: 'Assembly Screws', sourceType: 'ITEM', qty: 0, colorGroup: 'COMMON' }
    ])
  });
  assert(res.success, 'zero-quantity flat lot saves successfully: ' + res.message);
  const lotNumber = res.data && res.data.lotNumber;
  const lot = (getProductionData().data || []).find(l => l.lotNumber === lotNumber);
  assert(!!lot, 'saved zero-qty lot found by lot number');
  assert(lot && lot.qty === 0, `saved lot qty is 0 (got ${lot && lot.qty})`);
}

console.log('\n=== Test 2: flat path accepts a negative quantity (correction/reversal lot) ===');
{
  const res = saveProduction({
    processId: flatProcessId,
    assignedTo: 'Test Contractor',
    status: 'Pending',
    qty: -5,
    componentsConsumed: JSON.stringify([
      { itemName: 'Assembly Screws', sourceType: 'ITEM', qty: 20, colorGroup: 'COMMON' }
    ])
  });
  assert(res.success, 'negative-quantity flat lot saves successfully: ' + res.message);
  const lotNumber = res.data && res.data.lotNumber;
  const lot = (getProductionData().data || []).find(l => l.lotNumber === lotNumber);
  assert(!!lot, 'saved negative lot found by lot number');
  assert(lot && lot.qty === -5, `saved lot qty is -5 (got ${lot && lot.qty})`);
}

console.log('\n=== Setup: process with a simple (non-axis) color sub-group ===');
let colorProcessId;
{
  const res = saveProcess({
    processName: 'Painted Widget Assembly',
    sequence: 1,
    lotPrefix: 'PWA',
    outputItemName: 'Painted Widget',
    isFinalStage: false,
    active: true,
    remarks: '',
    components: JSON.stringify([
      { itemName: 'Assembly Screws', sourceType: 'ITEM', qtyPerUnit: 4, colorGroup: 'COMMON' },
      { itemName: 'Red Paint', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Red' },
      { itemName: 'Blue Paint', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Blue' }
    ])
  });
  assert(res.success, 'saveProcess succeeds: ' + res.message);
  colorProcessId = res.data && res.data.processId;
  assert(!!colorProcessId, 'processId returned');
}

console.log('\n=== Test 3: color-breakdown path accepts a single checked color at qty 0 ===');
{
  const res = saveProduction({
    processId: colorProcessId,
    assignedTo: 'Test Contractor',
    status: 'Pending',
    colorBreakdown: JSON.stringify([{ color: 'Red', qty: 0 }]),
    componentsConsumed: JSON.stringify([
      { itemName: 'Red Paint', sourceType: 'ITEM', qty: 0, colorGroup: 'Red' }
    ])
  });
  assert(res.success, 'zero-qty color-breakdown lot saves successfully: ' + res.message);
  const lotNumber = res.data && res.data.lotNumber;
  const lot = (getProductionData().data || []).find(l => l.lotNumber === lotNumber);
  assert(!!lot, 'saved zero-qty color lot found by lot number');
  assert(lot && lot.qty === 0, `saved lot qty is 0 (got ${lot && lot.qty})`);
}

console.log('\n=== Test 5: componentsConsumed accepts a negative ITEM-sourced qty (credits Stock back) ===');
{
  const res = saveProduction({
    processId: flatProcessId,
    assignedTo: 'Test Contractor',
    status: 'Pending',
    qty: 5,
    componentsConsumed: JSON.stringify([
      { itemName: 'Assembly Screws', sourceType: 'ITEM', qty: -20, colorGroup: 'COMMON' }
    ])
  });
  assert(res.success, 'negative component-qty lot saves successfully: ' + res.message);
  const lotNumber = res.data && res.data.lotNumber;
  const lot = (getProductionData().data || []).find(l => l.lotNumber === lotNumber);
  const comp = lot && lot.componentsConsumed && lot.componentsConsumed[0];
  assert(!!comp, 'saved lot retained its negative-qty component (not filtered out)');
  assert(comp && comp.qty === -20, `saved component qty is -20 (got ${comp && comp.qty})`);
}

console.log('\n=== Test 4: color-breakdown path accepts a negative total (correction/reversal lot) ===');
{
  const res = saveProduction({
    processId: colorProcessId,
    assignedTo: 'Test Contractor',
    status: 'Pending',
    colorBreakdown: JSON.stringify([{ color: 'Red', qty: -8 }]),
    componentsConsumed: JSON.stringify([
      { itemName: 'Red Paint', sourceType: 'ITEM', qty: 8, colorGroup: 'Red' }
    ])
  });
  assert(res.success, 'negative color-breakdown lot saves successfully: ' + res.message);
  const lotNumber = res.data && res.data.lotNumber;
  const lot = (getProductionData().data || []).find(l => l.lotNumber === lotNumber);
  assert(!!lot, 'saved negative color lot found by lot number');
  assert(lot && lot.qty === -8, `saved lot qty is -8 (got ${lot && lot.qty})`);
}

console.log(failures === 0 ? '\nALL TESTS PASSED' : `\n${failures} TEST(S) FAILED`);
process.exitCode = failures === 0 ? 0 : 1;
