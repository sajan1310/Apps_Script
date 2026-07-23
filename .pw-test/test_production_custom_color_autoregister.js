/**
 * Standalone Node harness (same mock-SpreadsheetApp+CacheService pattern as
 * test_color_master_widening.js) covering the new auto-register feature: a
 * genuinely new color typed via the Production form's "+ Add Custom
 * Sub-Group" control (colorBreakdown entry with isCustom:true) is now
 * automatically added to Color Master on save, instead of staying
 * invisible everywhere else Color Master feeds a picker (the Warehouse
 * Pool "+ Add Combination" datalist, another process's own custom-color
 * autocomplete) until someone separately re-typed it into the Color Master
 * screen by hand.
 *
 * Run: node .pw-test/test_production_custom_color_autoregister.js
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

// Real (not no-op) cache - needed for getColors()'s cache-hit path, and
// crucially includes .remove() (see the FakeScriptCache-mock gap found and
// fixed in test_color_master_widening.js) so invalidateListCache actually
// clears stale entries after _ensureColorMasterEntries writes new rows.
class FakeScriptCache {
  constructor() { this.store = {}; }
  get(key) { return Object.prototype.hasOwnProperty.call(this.store, key) ? this.store[key] : null; }
  put(key, value) { this.store[key] = value; }
  remove(key) { delete this.store[key]; }
  removeAll() { this.store = {}; }
}

const ss = new FakeSpreadsheet();
const scriptCache = new FakeScriptCache();

const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush: () => {} },
  CacheService: { getScriptCache: () => scriptCache },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  console,
  Logger: { log: () => {} },
  Utilities: { getUuid: () => 'uuid-' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);

['config.js', 'utils.js', 'module_tags.js', 'module_process.js', 'module_production.js', 'module_warehouse.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});

const { saveProcess, saveProduction, getColors, saveColor } = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); }
}

console.log('\n=== Setup: Color Master seeded with one existing color; a color-enabled process ===');
assert(saveColor({ name: 'Red' }).success, 'Color Master "Red" seeded');

const procRes = saveProcess({
  processName: 'Custom Color Test Process',
  sequence: 1,
  lotPrefix: 'CCT',
  outputItemName: 'Custom Color Test Output',
  isFinalStage: false,
  active: true,
  remarks: '',
  components: JSON.stringify([
    { itemName: 'Red Paint', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Red' }
  ])
});
assert(procRes.success, 'saveProcess succeeds: ' + procRes.message);
const processId = procRes.data && procRes.data.processId;

console.log('\n=== Test 1: a genuinely new custom color (not in Color Master) gets auto-registered on save ===');
{
  const before = getColors();
  assert(!before.data.some(c => c.name === 'Sunrise Coral'), 'sanity: "Sunrise Coral" is NOT in Color Master yet');

  const res = saveProduction({
    processId,
    assignedTo: 'Test Contractor',
    status: 'Pending',
    colorBreakdown: JSON.stringify([
      { color: 'Sunrise Coral', qty: 5, isCustom: true, countsTowardTotal: true }
    ]),
    componentsConsumed: JSON.stringify([{ itemName: 'Red Paint', sourceType: 'ITEM', qty: 5, colorGroup: 'COMMON' }])
  });
  assert(res.success, 'saveProduction (custom color "Sunrise Coral") succeeds: ' + res.message);

  const after = getColors();
  assert(after.data.some(c => c.name === 'Sunrise Coral'), `"Sunrise Coral" is now a real Color Master entry (got ${JSON.stringify(after.data.map(c => c.name))})`);
}

console.log('\n=== Test 2: a custom color that happens to match an EXISTING Color Master name (case-insensitively) does not create a duplicate ===');
{
  const beforeCount = getColors().data.length;
  const res = saveProduction({
    processId,
    assignedTo: 'Test Contractor',
    status: 'Pending',
    colorBreakdown: JSON.stringify([
      { color: 'red', qty: 3, isCustom: true, countsTowardTotal: true } // lowercase, but "Red" already exists
    ]),
    componentsConsumed: JSON.stringify([{ itemName: 'Red Paint', sourceType: 'ITEM', qty: 3, colorGroup: 'COMMON' }])
  });
  assert(res.success, 'saveProduction (custom color matching existing "Red") succeeds: ' + res.message);

  const afterColors = getColors();
  assert(afterColors.data.length === beforeCount, `no new Color Master row added - count unchanged (was ${beforeCount}, now ${afterColors.data.length})`);
  assert(afterColors.data.filter(c => c.name.toLowerCase() === 'red').length === 1, 'still exactly one "Red" entry, original casing preserved');
}

console.log('\n=== Test 3: a NON-custom colorBreakdown entry (a real recipe-tagged color) is never auto-registered - only isCustom:true entries are ===');
{
  const beforeCount = getColors().data.length;
  const res = saveProduction({
    processId,
    assignedTo: 'Test Contractor',
    status: 'Pending',
    colorBreakdown: JSON.stringify([
      { color: 'Red', qty: 2, isCustom: false, countsTowardTotal: true }
    ]),
    componentsConsumed: JSON.stringify([{ itemName: 'Red Paint', sourceType: 'ITEM', qty: 2, colorGroup: 'COMMON' }])
  });
  assert(res.success, 'saveProduction (non-custom "Red", already recipe-tagged) succeeds: ' + res.message);
  assert(getColors().data.length === beforeCount, 'Color Master count unchanged - nothing to register for a non-custom entry');
}

console.log('\n=== Test 4: two DIFFERENT new custom colors in the SAME lot both get registered ===');
{
  const res = saveProduction({
    processId,
    assignedTo: 'Test Contractor',
    status: 'Pending',
    colorBreakdown: JSON.stringify([
      { color: 'Midnight Teal', qty: 4, isCustom: true, countsTowardTotal: true },
      { color: 'Desert Sand', qty: 4, isCustom: true, countsTowardTotal: false }
    ]),
    componentsConsumed: JSON.stringify([{ itemName: 'Red Paint', sourceType: 'ITEM', qty: 4, colorGroup: 'COMMON' }])
  });
  assert(res.success, 'saveProduction (2 new custom colors in one lot) succeeds: ' + res.message);
  const names = getColors().data.map(c => c.name);
  assert(names.includes('Midnight Teal') && names.includes('Desert Sand'), `both new colors registered (got ${JSON.stringify(names)})`);
}

console.log('\n' + (failures === 0 ? 'ALL TESTS PASSED' : failures + ' TEST(S) FAILED'));
process.exit(failures === 0 ? 0 : 1);
