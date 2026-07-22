/**
 * Standalone Node harness (same pattern as test_color_master_widening.js)
 * covering the follow-up fix to getAllProcessColorGroups (module_process.js):
 * the Warehouse Pool breakdown dialog's known-colors list now also includes
 * colors this process's own Production history has actually logged, not
 * just recipe-tagged/pool-consumed-item colors. This catches a custom/
 * off-recipe color an operator picked at production time (via the
 * checklist's own separate full Color Master widening) that neither
 * signal would otherwise surface here - without needing to know which
 * specific Color Master name was used ahead of time. Covers both storage
 * shapes a saved lot's color can take: the JSON colorBreakdown array
 * (multi-color/axis lots) and the legacy comma-joined COLOR display
 * string (flat single-qty lots, including a "(Size)" qualifier suffix).
 *
 * Run: node .pw-test/test_warehouse_pool_logged_color_history.js
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
  global.PRODUCTION_COL = PRODUCTION_COL;
`, ctx, { filename: 'expose.js' });

const { APP_CONFIG, PRODUCTION_COL, saveProcess, saveProduction, getAllProcessColorGroups } = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); }
}

console.log('\n=== Setup: a plain process with NO recipe color tagging and NO pool-sourced components ===');
const procRes = saveProcess({
  processName: 'Freeform Packing',
  sequence: 1,
  lotPrefix: 'FFP',
  outputItemName: 'Freeform Packed Output',
  isFinalStage: false,
  active: true,
  remarks: '',
  components: JSON.stringify([
    { itemName: 'Box', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'COMMON' }
  ])
});
assert(procRes.success, 'saveProcess succeeds: ' + procRes.message);
const processId = procRes.data && procRes.data.processId;
assert(!!processId, 'processId returned');

console.log('\n=== Test 1: no logged lots yet -> zero known colors (no recipe/pool signal at all) ===');
{
  const res = getAllProcessColorGroups();
  assert(res.success, 'getAllProcessColorGroups succeeds');
  assert((res.data[processId].colors || []).length === 0, `zero colors before anything is logged (got ${JSON.stringify(res.data[processId])})`);
}

console.log('\n=== Test 2: a lot logged with a JSON colorBreakdown surfaces its color, even with no recipe/pool signal ===');
{
  const res = saveProduction({
    processId,
    assignedTo: 'Test Contractor',
    status: 'Completed',
    colorBreakdown: JSON.stringify([{ color: 'Sunrise Coral', qty: 5, isCustom: true }]),
    componentsConsumed: JSON.stringify([{ itemName: 'Box', sourceType: 'ITEM', qty: 1, colorGroup: 'COMMON' }])
  });
  assert(res.success, 'saveProduction (custom colorBreakdown color) succeeds: ' + res.message);

  const bulk = getAllProcessColorGroups();
  assert((bulk.data[processId].colors || []).includes('Sunrise Coral'), `"Sunrise Coral" now appears in known colors purely from logged history (got ${JSON.stringify(bulk.data[processId])})`);
  assert((bulk.data[processId].removable || []).includes('Sunrise Coral'), '"Sunrise Coral" is removable-looking (not recipe/pool-configured)');
}

console.log('\n=== Test 3: a legacy flat lot (comma-joined COLOR string, no colorBreakdown) also surfaces its color ===');
{
  // Directly seed a row shaped like a pre-colorBreakdown-era legacy lot:
  // COLOR_BREAKDOWN blank, COLOR holds the plain display string.
  const prodSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PRODUCTION);
  const row = new Array(PRODUCTION_COL.ORDER_NUMBER).fill('');
  row[PRODUCTION_COL.PROCESS_ID - 1] = processId;
  row[PRODUCTION_COL.STATUS - 1] = 'Completed';
  row[PRODUCTION_COL.COLOR - 1] = 'Twilight Amber (Large)'; // "(Size)" qualifier suffix must be stripped
  row[PRODUCTION_COL.COLOR_BREAKDOWN - 1] = '';
  row[PRODUCTION_COL.OUTPUT_ITEM_NAME - 1] = 'Freeform Packed Output';
  prodSheet.appendRow(row);

  const bulk = getAllProcessColorGroups();
  assert((bulk.data[processId].colors || []).includes('Twilight Amber'), `"Twilight Amber" (size suffix stripped) surfaces from the legacy comma-joined COLOR column (got ${JSON.stringify(bulk.data[processId])})`);
}

console.log('\n=== Test 4: a comma-joined multi-color legacy COLOR string splits into separate entries ===');
{
  const prodSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PRODUCTION);
  const row = new Array(PRODUCTION_COL.ORDER_NUMBER).fill('');
  row[PRODUCTION_COL.PROCESS_ID - 1] = processId;
  row[PRODUCTION_COL.STATUS - 1] = 'Completed';
  row[PRODUCTION_COL.COLOR - 1] = 'Marigold, Steel Grey';
  row[PRODUCTION_COL.COLOR_BREAKDOWN - 1] = '';
  row[PRODUCTION_COL.OUTPUT_ITEM_NAME - 1] = 'Freeform Packed Output';
  prodSheet.appendRow(row);

  const bulk = getAllProcessColorGroups();
  const colors = bulk.data[processId].colors || [];
  assert(colors.includes('Marigold') && colors.includes('Steel Grey'), `comma-joined "Marigold, Steel Grey" split into 2 separate known colors (got ${JSON.stringify(colors)})`);
}

console.log('\n=== Test 5: a DIFFERENT process\'s logged colors never leak into this one ===');
{
  const otherProc = saveProcess({
    processName: 'Unrelated Process',
    sequence: 2,
    lotPrefix: 'UNR',
    outputItemName: 'Unrelated Output',
    isFinalStage: false,
    active: true,
    remarks: '',
    components: JSON.stringify([{ itemName: 'Widget', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'COMMON' }])
  });
  const otherId = otherProc.data && otherProc.data.processId;
  saveProduction({
    processId: otherId,
    assignedTo: 'Test Contractor',
    status: 'Completed',
    colorBreakdown: JSON.stringify([{ color: 'Should Not Leak', qty: 1, isCustom: true }]),
    componentsConsumed: JSON.stringify([{ itemName: 'Widget', sourceType: 'ITEM', qty: 1, colorGroup: 'COMMON' }])
  });

  const bulk = getAllProcessColorGroups();
  assert(!(bulk.data[processId].colors || []).includes('Should Not Leak'), 'the other process\'s own logged color does not appear on Freeform Packing\'s list');
  assert((bulk.data[otherId].colors || []).includes('Should Not Leak'), 'but it does correctly appear on ITS OWN process\'s list');
}

console.log('\n' + (failures === 0 ? 'ALL TESTS PASSED' : failures + ' TEST(S) FAILED'));
process.exit(failures === 0 ? 0 : 1);
