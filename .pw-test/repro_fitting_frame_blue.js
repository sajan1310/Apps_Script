/**
 * Repro: "Fitting Frame" process with 3 POOL-sourced, auto-detected color
 * axes (Painted Frame / Fitted Rim / Mudguard), mirroring the live screenshot
 * that showed a save-time rejection: 'Color "Blue" is not a configured color
 * sub-group for this process.' Checks whether this is reproducible with NO
 * data change between "form open" (getProcessColorAxes) and "form save"
 * (saveProduction -> getProcessColorGroups), i.e. whether it's a real bug or
 * requires a genuine timing/data-change race.
 *
 * Run: node .pw-test/repro_fitting_frame_blue.js
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
  global.PROCESS_COL = PROCESS_COL;
`, ctx, { filename: 'expose.js' });

const {
  APP_CONFIG,
  saveProcess, getProcessColorAxes, getProcessColorGroups,
  saveProduction
} = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); }
}

// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Setup: seed Warehouse Pool stock for 3 pool items (mirrors screenshot) ===');
{
  const poolSheet = ss.addSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
  poolSheet.appendRow(['Output Item Name', 'Process ID', 'Product Tag', 'Produced Qty', 'Consumed Qty', 'Available Qty', 'Color']);

  const rows = [
    ['Painted Frame Crysta 20 inch D/Gaddi', 'PRC-P', '', 'Blue-White'],
    ['Painted Frame Crysta 20 inch D/Gaddi', 'PRC-P', '', 'Orange-White'],
    ['Painted Frame Crysta 20 inch D/Gaddi', 'PRC-P', '', 'Pink-White'],
    ['Painted Frame Crysta 20 inch D/Gaddi', 'PRC-P', '', 'Purple-White'],
    ['Painted Frame Crysta 20 inch D/Gaddi', 'PRC-P', '', 'Red-White'],
    ['Painted Frame Crysta 20 inch D/Gaddi', 'PRC-P', '', 'Sea Green-White'],
    ['Fitted Rim 20 inch', 'PRC-R', '', 'BCP'],
    ['Fitted Rim 20 inch', 'PRC-R', '', 'Black'],
    ['20 inch Mudguard', 'PRC-M', '', 'B/T Green'],
    ['20 inch Mudguard', 'PRC-M', '', 'Black'],
    ['20 inch Mudguard', 'PRC-M', '', 'Blue'],
    ['20 inch Mudguard', 'PRC-M', '', 'Grey'],
    ['20 inch Mudguard', 'PRC-M', '', 'Green'],
    ['20 inch Mudguard', 'PRC-M', '', 'Metallic Green'],
    ['20 inch Mudguard', 'PRC-M', '', 'Orange'],
    ['20 inch Mudguard', 'PRC-M', '', 'Pink'],
    ['20 inch Mudguard', 'PRC-M', '', 'Purple'],
    ['20 inch Mudguard', 'PRC-M', '', 'Red'],
    ['20 inch Mudguard', 'PRC-M', '', 'Sea Green'],
    ['20 inch Mudguard', 'PRC-M', '', 'Silky Blue']
  ];
  rows.forEach(([item, pid, tag, color]) => {
    poolSheet.appendRow([item, pid, tag, 100, 0, 100, color]);
  });
  console.log('  Seeded', rows.length, 'pool rows across 3 items.');
}

console.log('\n=== Setup: Fitting Frame process, 3 POOL-sourced COMMON components, no primaryColorAxis yet ===');
let processId;
{
  const res = saveProcess({
    processName: 'Fitting Frame',
    sequence: 5,
    lotPrefix: 'FTF',
    outputItemName: 'Fitted Frame Assembled',
    isFinalStage: false,
    active: true,
    remarks: '',
    components: JSON.stringify([
      { itemName: 'Painted Frame Crysta 20 inch D/Gaddi', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' },
      { itemName: 'Fitted Rim 20 inch', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' },
      { itemName: '20 inch Mudguard', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' }
    ])
  });
  assert(res.success, 'saveProcess succeeds: ' + res.message);
  processId = res.data && res.data.processId;
}

console.log('\n=== Step 1 ("form open"): getProcessColorAxes ===');
let axes = [];
{
  const res = getProcessColorAxes(processId);
  assert(res.success, 'getProcessColorAxes succeeds: ' + res.message);
  axes = (res.data && res.data.axes) || [];
  console.log('  Axes found:', axes.map(a => `${a.label} [${a.colors.join(', ')}]`).join(' | '));
  assert(axes.length === 3, 'exactly 3 axes reported (got ' + axes.length + ')');
}

console.log('\n=== Step 1b: set Primary Axis to the Painted Frame axis (mirrors operator/default pick) ===');
{
  const paintedAxis = axes.find(a => a.label.includes('Painted Frame'));
  assert(!!paintedAxis, 'Painted Frame axis present');
  const res2 = saveProcess({
    processId,
    processName: 'Fitting Frame',
    sequence: 5,
    lotPrefix: 'FTF',
    outputItemName: 'Fitted Frame Assembled',
    isFinalStage: false,
    active: true,
    remarks: '',
    primaryColorAxis: paintedAxis.label,
    components: JSON.stringify([
      { itemName: 'Painted Frame Crysta 20 inch D/Gaddi', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' },
      { itemName: 'Fitted Rim 20 inch', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' },
      { itemName: '20 inch Mudguard', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' }
    ])
  });
  assert(res2.success, 'saveProcess (set primary axis) succeeds: ' + res2.message);
}

console.log('\n=== Step 2 ("form open" again, as operator would see it): getProcessColorGroups (validation-source list) ===');
let availableColors = [];
{
  const res = getProcessColorGroups(processId);
  assert(res.success, 'getProcessColorGroups succeeds: ' + res.message);
  availableColors = res.data || [];
  console.log('  availableColorGroups:', JSON.stringify(availableColors));
  assert(availableColors.some(c => c.toLowerCase() === 'blue'), '"Blue" IS present in availableColorGroups (got ' + JSON.stringify(availableColors) + ')');
}

console.log('\n=== Step 3 ("form save", NO data change since Step 2): saveProduction with Blue-White(primary) + matching Mudguard colors incl. Blue ===');
{
  const res = saveProduction({
    processId: processId,
    assignedTo: 'Sanjay',
    status: 'Completed',
    colorBreakdown: JSON.stringify([
      { color: 'Blue-White', qty: 10 },
      { color: 'Orange-White', qty: 7 },
      { color: 'Pink-White', qty: 17 },
      { color: 'Purple-White', qty: 7 },
      { color: 'Red-White', qty: 17 },
      { color: 'Sea Green-White', qty: 10 },
      { color: 'BCP', qty: 68 },
      { color: 'Blue', qty: 10 },
      { color: 'Orange', qty: 7 },
      { color: 'Pink', qty: 17 },
      { color: 'Purple', qty: 7 },
      { color: 'Red', qty: 17 },
      { color: 'Sea Green', qty: 10 }
    ]),
    componentsConsumed: JSON.stringify([
      { itemName: 'Painted Frame Crysta 20 inch D/Gaddi', sourceType: 'POOL', qty: 10, colorGroup: 'Blue-White' },
      { itemName: 'Fitted Rim 20 inch', sourceType: 'POOL', qty: 68, colorGroup: 'BCP' },
      { itemName: '20 inch Mudguard', sourceType: 'POOL', qty: 10, colorGroup: 'Blue' }
    ])
  });
  console.log('  saveProduction result:', JSON.stringify({ success: res.success, message: res.message }));
  assert(res.success, 'saveProduction succeeds with no data change since form-open (got: ' + res.message + ')');
}

console.log('\n' + (failures === 0 ? `ALL PASS` : `${failures} FAILURE(S)`));
process.exit(failures === 0 ? 0 : 1);
