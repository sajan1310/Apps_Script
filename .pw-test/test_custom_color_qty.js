/**
 * Regression test: a custom-added color (Script.html's addCustomColorRow,
 * "+ Add Custom Sub-Group") on a process with a resolved Primary Color Axis
 * was being silently dropped from the saved lot's total quantity. Root
 * cause: saveProduction's qty summation, once a primaryAxisColorsLower set
 * is resolved, only counted colorBreakdown entries whose color name matched
 * that axis's own recipe-derived colors — a custom color is by definition
 * not part of any configured axis, so it never matched, even though the
 * client's own running total (_currentLotTotalQty) already included it.
 *
 * Run: node .pw-test/test_custom_color_qty.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) { this.sheet = sheet; this.row = row; this.col = col; this.numRows = numRows; this.numCols = numCols; }
  getValues() { const out = []; for (let r = 0; r < this.numRows; r++) { const rowArr = []; for (let c = 0; c < this.numCols; c++) rowArr.push(this.sheet._get(this.row + r, this.col + c)); out.push(rowArr); } return out; }
  getValue() { return this.sheet._get(this.row, this.col); }
  setValues(values) { values.forEach((rowArr, r) => rowArr.forEach((val, c) => this.sheet._set(this.row + r, this.col + c, val))); return this; }
  setValue(v) { this.sheet._set(this.row, this.col, v); return this; }
  clearContent() { for (let r = 0; r < this.numRows; r++) for (let c = 0; c < this.numCols; c++) this.sheet._set(this.row + r, this.col + c, ''); return this; }
  setFontWeight() { return this; }
  setBackground() { return this; }
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
  deleteRow(r) { this.rows.splice(r - 1, 1); }
  deleteRows(r, n) { this.rows.splice(r - 1, n); }
  insertRows(r, n) { for (let i = 0; i < n; i++) this.rows.splice(r - 1, 0, []); }
  insertColumnsAfter(afterPosition, howMany) { this.rows.forEach(row => { const blanks = new Array(howMany).fill(''); row.splice(afterPosition, 0, ...blanks); }); }
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
  console, Logger: { log: () => {} },
  Utilities: { getUuid: () => 'uuid-' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);
['config.js', 'utils.js', 'module_process.js', 'module_production.js', 'module_warehouse.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});
vm.runInContext('global.APP_CONFIG = APP_CONFIG;', ctx, { filename: 'expose.js' });
const { saveProcess, saveProduction, getProductionData } = ctx;

let failures = 0;
function assert(cond, msg) { if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); } }

console.log('\n=== Setup: process with 2 tag-based Color Axes, Rim Color primary (mirrors test_color_axes.js) ===');
const res = saveProcess({
  processName: 'Fitted Frame Assembly', sequence: 1, lotPrefix: 'FFA2', outputItemName: 'Fitted Frame 2',
  isFinalStage: false, active: true, remarks: '', primaryColorAxis: 'Rim Color',
  components: JSON.stringify([
    { itemName: 'Assembly Screws', sourceType: 'ITEM', qtyPerUnit: 4, colorGroup: 'COMMON' },
    { itemName: 'Painted Rim - Blue-White', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Blue-White', colorAxis: 'Rim Color' },
    { itemName: 'Painted Rim - Red-White', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Red-White', colorAxis: 'Rim Color' },
    { itemName: 'Mudguard - Blue', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Blue', colorAxis: 'Mudguard Color' },
    { itemName: 'Mudguard - Red', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Red', colorAxis: 'Mudguard Color' }
  ])
});
assert(res.success, 'saveProcess succeeds: ' + res.message);
const processId = res.data && res.data.processId;

console.log('\n=== Test: custom color counted alongside a real primary-axis color ===');
{
  // Mirrors what Script.html's getCheckedColorQtys() actually sends for a
  // custom row: isCustom true, countsTowardTotal true (no data-primary
  // attribute at all on a custom row, so it's never flagged non-primary).
  const r = saveProduction({
    processId, assignedTo: 'Test Contractor', status: 'Pending',
    colorBreakdown: JSON.stringify([
      { color: 'Red-White', qty: 10, isCustom: false, countsTowardTotal: true },
      { color: 'Neon Green', qty: 5, isCustom: true, countsTowardTotal: true }
    ]),
    componentsConsumed: JSON.stringify([
      { itemName: 'Painted Rim - Red-White', sourceType: 'ITEM', qty: 10, colorGroup: 'Red-White' }
    ])
  });
  assert(r.success, 'saveProduction succeeds: ' + r.message);
  const lotNumber = r.data && r.data.lotNumber;
  const lot = (getProductionData().data || []).find(l => l.lotNumber === lotNumber);
  assert(!!lot, 'saved lot found');
  assert(lot && lot.qty === 15, 'lot quantity is 15 (10 Red-White + 5 custom Neon Green), NOT 10 (got ' + (lot && lot.qty) + ')');
}

console.log('\n=== Test: custom color alone (no real axis color checked) still counts ===');
{
  const r = saveProduction({
    processId, assignedTo: 'Test Contractor', status: 'Pending',
    colorBreakdown: JSON.stringify([
      { color: 'One-Off Special', qty: 7, isCustom: true, countsTowardTotal: true }
    ]),
    componentsConsumed: JSON.stringify([
      { itemName: 'Assembly Screws', sourceType: 'ITEM', qty: 4, colorGroup: 'COMMON' }
    ])
  });
  assert(r.success, 'saveProduction succeeds with only a custom color: ' + r.message);
  const lotNumber = r.data && r.data.lotNumber;
  const lot = (getProductionData().data || []).find(l => l.lotNumber === lotNumber);
  assert(lot && lot.qty === 7, 'lot quantity is 7 (the custom color alone), NOT 0 (got ' + (lot && lot.qty) + ')');
}

console.log('\n=== Test: custom color placed in a NON-primary group must NOT double-count ===');
{
  // Mirrors what addCustomColorRow now sends when the operator picks a
  // non-primary group (e.g. "Mudguard Color") from the new group selector:
  // data-primary="false" on that row means getCheckedColorQtys() computes
  // countsTowardTotal: false for it, same as any real Mudguard Color row.
  const r = saveProduction({
    processId, assignedTo: 'Test Contractor', status: 'Pending',
    colorBreakdown: JSON.stringify([
      { color: 'Red-White', qty: 10, isCustom: false, countsTowardTotal: true },
      { color: 'Neon Green Mudguard', qty: 10, isCustom: true, countsTowardTotal: false }
    ]),
    componentsConsumed: JSON.stringify([
      { itemName: 'Painted Rim - Red-White', sourceType: 'ITEM', qty: 10, colorGroup: 'Red-White' }
    ])
  });
  assert(r.success, 'saveProduction succeeds: ' + r.message);
  const lotNumber = r.data && r.data.lotNumber;
  const lot = (getProductionData().data || []).find(l => l.lotNumber === lotNumber);
  assert(lot && lot.qty === 10, 'lot quantity is 10 (Red-White only) — the non-primary custom Mudguard color describes the SAME batch, NOT an extra 10 (got ' + (lot && lot.qty) + ')');
  assert(Array.isArray(lot.colorBreakdown) && lot.colorBreakdown.length === 2, 'both breakdown rows still recorded for consumption/history (got ' + JSON.stringify(lot && lot.colorBreakdown) + ')');
}

console.log('\n' + (failures === 0 ? 'ALL TESTS PASSED' : failures + ' TEST(S) FAILED'));
process.exit(failures === 0 ? 0 : 1);
