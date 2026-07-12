/**
 * Variant of repro_fitting_frame_blue.js: same 3-pool-axis "Fitting Frame"
 * setup, but WITHOUT ever configuring/saving a Primary Color Axis on the
 * process (i.e. process.primaryColorAxis stays blank, only picked live on
 * the form via formData.primaryColorAxis, as the very first lot for this
 * process would look before anyone visits the Process editor's Primary
 * Axis picker). Checks whether getProcessColorGroups (validation source)
 * still recognizes "Blue" as valid in this state.
 *
 * Run: node .pw-test/repro_fitting_frame_no_primary.js
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
vm.runInContext('global.APP_CONFIG = APP_CONFIG; global.PROCESS_COL = PROCESS_COL;', ctx, { filename: 'expose.js' });
const { APP_CONFIG, saveProcess, getProcessColorAxes, getProcessColorGroups, saveProduction } = ctx;

let failures = 0;
function assert(cond, msg) { if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); } }

console.log('\n=== Seed pool stock ===');
const poolSheet = ss.addSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
poolSheet.appendRow(['Output Item Name', 'Process ID', 'Product Tag', 'Produced Qty', 'Consumed Qty', 'Available Qty', 'Color']);
[
  ['Painted Frame Crysta 20 inch D/Gaddi', 'PRC-P', '', 'Blue-White'],
  ['Painted Frame Crysta 20 inch D/Gaddi', 'PRC-P', '', 'Orange-White'],
  ['Fitted Rim 20 inch', 'PRC-R', '', 'BCP'],
  ['Fitted Rim 20 inch', 'PRC-R', '', 'Black'],
  ['20 inch Mudguard', 'PRC-M', '', 'Blue'],
  ['20 inch Mudguard', 'PRC-M', '', 'Orange']
].forEach(function (row) {
  poolSheet.appendRow([row[0], row[1], row[2], 100, 0, 100, row[3]]);
});

console.log('\n=== Create Fitting Frame process WITHOUT ever setting primaryColorAxis ===');
const res = saveProcess({
  processName: 'Fitting Frame', sequence: 5, lotPrefix: 'FTF2', outputItemName: 'Fitted Frame Assembled 2',
  isFinalStage: false, active: true, remarks: '',
  components: JSON.stringify([
    { itemName: 'Painted Frame Crysta 20 inch D/Gaddi', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' },
    { itemName: 'Fitted Rim 20 inch', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' },
    { itemName: '20 inch Mudguard', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' }
  ])
});
assert(res.success, 'saveProcess succeeds: ' + res.message);
const processId = res.data && res.data.processId;

console.log('\n=== Client would render (getProcessColorAxes) ===');
{
  const r = getProcessColorAxes(processId);
  const axes = (r.data && r.data.axes) || [];
  console.log('  axes:', axes.map(function (a) { return a.label + '[' + a.colors.join(',') + ']'; }).join(' | '));
  console.log('  primaryAxisKey:', JSON.stringify(r.data.primaryAxisKey));
  assert(axes.length >= 2, 'client sees 2+ axes -> renders SPLIT checklist with individual colors like "Blue"');
}

console.log('\n=== Server validates against (getProcessColorGroups) ===');
{
  const r = getProcessColorGroups(processId);
  console.log('  availableColorGroups:', JSON.stringify(r.data));
  const hasBlue = (r.data || []).some(function (c) { return c.toLowerCase() === 'blue'; });
  assert(hasBlue, '"Blue" present as standalone entry in availableColorGroups (got ' + JSON.stringify(r.data) + ')');
}

console.log('\n=== saveProduction, operator picks Primary live on the form (formData.primaryColorAxis) ===');
{
  const r = saveProduction({
    processId: processId, assignedTo: 'Sanjay', status: 'Completed',
    primaryColorAxis: 'Painted Frame Crysta 20 inch D/Gaddi',
    colorBreakdown: JSON.stringify([
      { color: 'Blue-White', qty: 10 },
      { color: 'BCP', qty: 10 },
      { color: 'Blue', qty: 10 }
    ]),
    componentsConsumed: JSON.stringify([
      { itemName: 'Painted Frame Crysta 20 inch D/Gaddi', sourceType: 'POOL', qty: 10, colorGroup: 'Blue-White' },
      { itemName: 'Fitted Rim 20 inch', sourceType: 'POOL', qty: 10, colorGroup: 'BCP' },
      { itemName: '20 inch Mudguard', sourceType: 'POOL', qty: 10, colorGroup: 'Blue' }
    ])
  });
  console.log('  saveProduction result:', JSON.stringify({ success: r.success, message: r.message }));
  assert(r.success, 'saveProduction succeeds even though process.primaryColorAxis was never saved (got: ' + r.message + ')');
}

console.log('\n' + (failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'));
process.exit(failures === 0 ? 0 : 1);
