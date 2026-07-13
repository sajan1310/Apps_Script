/**
 * Standalone Node harness that mocks the Apps Script runtime (SpreadsheetApp,
 * LockService, etc.) well enough to load and execute the REAL server-side
 * files (config.js, utils.js, module_process.js, module_production.js,
 * module_warehouse.js) and exercise the Process recipe color sub-group
 * feature end-to-end: saving a process with Common + color-scoped
 * components, reading it back grouped, legacy-sheet column backfill, and
 * the Production Lot Color requirement/validation this feature adds.
 *
 * Run: node .pw-test/test_process_color_groups.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

// ─────────────────────────────────────────────────────────────────────────
// Minimal in-memory Sheet/Range/Spreadsheet mocks (mirrors test_merge_and_backfill.js)
// ─────────────────────────────────────────────────────────────────────────

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
  }
  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowArr = [];
      for (let c = 0; c < this.numCols; c++) {
        rowArr.push(this.sheet._get(this.row + r, this.col + c));
      }
      out.push(rowArr);
    }
    return out;
  }
  getValue() {
    return this.sheet._get(this.row, this.col);
  }
  setValues(values) {
    values.forEach((rowArr, r) => {
      rowArr.forEach((val, c) => {
        this.sheet._set(this.row + r, this.col + c, val);
      });
    });
    return this;
  }
  setValue(v) {
    this.sheet._set(this.row, this.col, v);
    return this;
  }
  clearContent() {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) {
        this.sheet._set(this.row + r, this.col + c, '');
      }
    }
    return this;
  }
  setFontWeight() { return this; }
  setBackground() { return this; }
}

class FakeSheet {
  constructor(name) {
    this.name = name;
    this.rows = []; // rows[0] is sheet row 1
  }
  getName() { return this.name; }
  _ensureRow(r) {
    while (this.rows.length < r) this.rows.push([]);
  }
  _get(r, c) {
    this._ensureRow(r);
    const row = this.rows[r - 1];
    return row[c - 1] === undefined ? '' : row[c - 1];
  }
  _set(r, c, v) {
    this._ensureRow(r);
    const row = this.rows[r - 1];
    while (row.length < c) row.push('');
    row[c - 1] = v;
  }
  getLastRow() {
    for (let r = this.rows.length; r >= 1; r--) {
      const row = this.rows[r - 1];
      if (row.some(v => v !== '' && v !== undefined && v !== null)) return r;
    }
    return 0;
  }
  getLastColumn() {
    let max = 0;
    this.rows.forEach(row => {
      for (let c = row.length; c >= 1; c--) {
        if (row[c - 1] !== '' && row[c - 1] !== undefined && row[c - 1] !== null) {
          max = Math.max(max, c);
          break;
        }
      }
    });
    return max;
  }
  getRange(row, col, numRows = 1, numCols = 1) {
    return new FakeRange(this, row, col, numRows, numCols);
  }
  appendRow(arr) {
    const r = this.getLastRow() + 1;
    arr.forEach((v, i) => this._set(r, i + 1, v));
  }
  deleteRow(r) {
    this.rows.splice(r - 1, 1);
  }
  deleteRows(r, n) {
    this.rows.splice(r - 1, n);
  }
  insertRows(r, n) {
    for (let i = 0; i < n; i++) this.rows.splice(r - 1, 0, []);
  }
  // Inserts `howMany` blank columns immediately after 1-based column
  // `afterPosition`, shifting any existing data right — mirrors the real
  // Sheets API call used by the ensure*Column() backfill helpers.
  insertColumnsAfter(afterPosition, howMany) {
    this.rows.forEach(row => {
      const blanks = new Array(howMany).fill('');
      row.splice(afterPosition, 0, ...blanks);
    });
  }
}

class FakeSpreadsheet {
  constructor() {
    this.sheets = {};
  }
  getSheetByName(name) {
    return this.sheets[name] || null;
  }
  addSheet(name) {
    const s = new FakeSheet(name);
    this.sheets[name] = s;
    return s;
  }
  insertSheet(name) {
    return this.addSheet(name);
  }
}

// ─────────────────────────────────────────────────────────────────────────
// Build sandbox, load real files into it
// ─────────────────────────────────────────────────────────────────────────

const ss = new FakeSpreadsheet();

const sandbox = {
  SpreadsheetApp: {
    getActiveSpreadsheet: () => ss,
    flush: () => {}
  },
  LockService: {
    getDocumentLock: () => ({
      tryLock: () => true,
      releaseLock: () => {}
    })
  },
  console,
  Logger: { log: () => {} },
  Utilities: { getUuid: () => 'uuid-' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) }
};
sandbox.global = sandbox;

const ctx = vm.createContext(sandbox);

const files = [
  'config.js',
  'utils.js',
  'module_process.js',
  'module_production.js',
  'module_warehouse.js'
];

files.forEach(f => {
  const code = fs.readFileSync(path.join(ROOT, f), 'utf8');
  vm.runInContext(code, ctx, { filename: f });
});

// `const`/`let` top-level declarations don't auto-expose onto the sandbox
// object the way `var`/function declarations do — re-expose what we need.
vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.PROCESS_COMPONENTS_COL = PROCESS_COMPONENTS_COL;
  global.PRODUCTION_COL = PRODUCTION_COL;
  global.COMPONENT_COLOR_GROUP_COMMON = COMPONENT_COLOR_GROUP_COMMON;
`, ctx, { filename: 'expose.js' });

const {
  APP_CONFIG, PROCESS_COMPONENTS_COL, PRODUCTION_COL, COMPONENT_COLOR_GROUP_COMMON,
  saveProcess, getProcessComponentsData, getProcessColorGroups,
  saveProduction, getProductionData
} = ctx;

// ─────────────────────────────────────────────────────────────────────────
let failures = 0;
function assert(cond, msg) {
  if (!cond) {
    failures++;
    console.error('FAIL:', msg);
  } else {
    console.log('PASS:', msg);
  }
}

// ─────────────────────────────────────────────────────────────────────────
// Test 1: create a process with Common + 2 color sub-groups
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 1: saveProcess persists Common + color sub-group components ===');
let framePaintingId;
{
  const res = saveProcess({
    processName: 'Frame Painting',
    sequence: 1,
    lotPrefix: 'ZFP',
    outputItemName: 'Painted Frame Test Output',
    isFinalStage: false,
    active: true,
    remarks: '',
    components: JSON.stringify([
      { itemName: 'Brush', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'COMMON' },
      { itemName: 'Sandpaper', sourceType: 'ITEM', qtyPerUnit: 1 }, // no colorGroup at all -> must default to COMMON
      { itemName: 'Red Paint', sourceType: 'ITEM', qtyPerUnit: 2, colorGroup: 'Red' },
      { itemName: 'Blue Paint', sourceType: 'ITEM', qtyPerUnit: 2, colorGroup: 'Blue' }
    ])
  });
  assert(res.success, 'saveProcess succeeds: ' + res.message);
  framePaintingId = res.data && res.data.processId;
  assert(!!framePaintingId, 'processId returned (got "' + framePaintingId + '")');

  const compRes = getProcessComponentsData(framePaintingId);
  assert(compRes.success, 'getProcessComponentsData succeeds');
  const comps = compRes.data || [];
  assert(comps.length === 4, 'all 4 components persisted (got ' + comps.length + ')');

  const byName = Object.fromEntries(comps.map(c => [c.itemName, c]));
  assert(byName['Brush'] && byName['Brush'].colorGroup === 'COMMON', 'Brush tagged COMMON (got "' + (byName['Brush'] && byName['Brush'].colorGroup) + '")');
  assert(byName['Sandpaper'] && byName['Sandpaper'].colorGroup === 'COMMON', 'Sandpaper with no colorGroup defaults to COMMON (got "' + (byName['Sandpaper'] && byName['Sandpaper'].colorGroup) + '")');
  assert(byName['Red Paint'] && byName['Red Paint'].colorGroup === 'Red', 'Red Paint tagged Red (got "' + (byName['Red Paint'] && byName['Red Paint'].colorGroup) + '")');
  assert(byName['Blue Paint'] && byName['Blue Paint'].colorGroup === 'Blue', 'Blue Paint tagged Blue (got "' + (byName['Blue Paint'] && byName['Blue Paint'].colorGroup) + '")');

  const compSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);
  // 10, not 9 -- PROCESS_COMPONENTS_COL.UNIT (Fix #3, 2026-07-13) added a
  // 10th column after Color Axis.
  assert(compSheet.getLastColumn() === 10, 'Process Components sheet has 10 columns (got ' + compSheet.getLastColumn() + ')');
  assert(compSheet.getRange(1, 8, 1, 1).getValue() === 'Color Group', 'column 8 header is "Color Group"');
  assert(compSheet.getRange(1, 9, 1, 1).getValue() === 'Color Axis', 'column 9 header is "Color Axis"');
  assert(compSheet.getRange(1, 10, 1, 1).getValue() === 'Unit', 'column 10 header is "Unit"');

  const colorsRes = getProcessColorGroups(framePaintingId);
  assert(colorsRes.success, 'getProcessColorGroups succeeds');
  assert(JSON.stringify(colorsRes.data) === JSON.stringify(['Blue', 'Red']), 'distinct color groups sorted, COMMON excluded (got ' + JSON.stringify(colorsRes.data) + ')');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 2: legacy Process Components rows (no Color Group column) backfill to COMMON
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 2: legacy Process Components rows backfill to COMMON ===');
{
  // Simulate a pre-existing sheet from before this feature: only 7 columns,
  // written directly (bypassing saveProcess/_saveProcessComponentsForProcess).
  const compSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);
  const legacyRow = compSheet.getLastRow() + 1;
  compSheet._set(legacyRow, PROCESS_COMPONENTS_COL.PROCESS_ID, 'PRC-LEGACY');
  compSheet._set(legacyRow, PROCESS_COMPONENTS_COL.ITEM_NAME, 'Old Item');
  compSheet._set(legacyRow, PROCESS_COMPONENTS_COL.SIZE, '');
  compSheet._set(legacyRow, PROCESS_COMPONENTS_COL.NARRATION, '');
  compSheet._set(legacyRow, PROCESS_COMPONENTS_COL.QTY_PER_UNIT, 1);
  compSheet._set(legacyRow, PROCESS_COMPONENTS_COL.REMARKS, '');
  compSheet._set(legacyRow, PROCESS_COMPONENTS_COL.SOURCE_TYPE, 'ITEM');
  // Column 8 (Color Group) deliberately left unset, as if it never existed.

  const res = getProcessComponentsData('PRC-LEGACY');
  assert(res.success, 'getProcessComponentsData succeeds on legacy row');
  const legacy = (res.data || []).find(c => c.itemName === 'Old Item');
  assert(!!legacy, 'legacy row found');
  assert(legacy.colorGroup === COMPONENT_COLOR_GROUP_COMMON, 'legacy row defaults to COMMON (got "' + (legacy && legacy.colorGroup) + '")');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 3: a process with no color sub-groups reports an empty color list
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 3: process with only Common components has no color groups ===');
let packingId;
{
  const res = saveProcess({
    processName: 'Bicycle Packing',
    sequence: 4,
    lotPrefix: 'ZPK',
    outputItemName: 'Packed Bicycle Test Output',
    isFinalStage: true,
    active: true,
    remarks: '',
    components: JSON.stringify([
      { itemName: 'Packing Box', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'COMMON' }
    ])
  });
  assert(res.success, 'saveProcess (Bicycle Packing) succeeds: ' + res.message);
  packingId = res.data && res.data.processId;

  const colorsRes = getProcessColorGroups(packingId);
  assert(colorsRes.success && colorsRes.data.length === 0, 'no color groups reported (got ' + JSON.stringify(colorsRes.data) + ')');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 4: saveProduction requires Color when the process has color sub-groups
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 4: saveProduction Color requirement/validation ===');
{
  const baseLot = {
    processId: framePaintingId,
    qty: 5,
    assignedTo: 'Test Contractor',
    status: 'Pending',
    componentsConsumed: JSON.stringify([
      { itemName: 'Brush', sourceType: 'ITEM', qty: 5 },
      { itemName: 'Red Paint', sourceType: 'ITEM', qty: 10 }
    ])
  };

  const noColor = saveProduction(Object.assign({}, baseLot)); // color omitted entirely
  assert(noColor.success === false, 'missing Color is rejected when process has color sub-groups');
  assert(/Color with a non-zero quantity/i.test(noColor.message), 'rejection message explains Color is required (got "' + noColor.message + '")');

  const badColor = saveProduction(Object.assign({}, baseLot, { colorBreakdown: JSON.stringify([{ color: 'Green', qty: 5 }]) }));
  assert(badColor.success === false, 'unconfigured Color ("Green") is rejected');
  assert(/not a configured color sub-group/i.test(badColor.message), 'rejection message names the bad color (got "' + badColor.message + '")');

  const goodColor = saveProduction(Object.assign({}, baseLot, { colorBreakdown: JSON.stringify([{ color: 'Red', qty: 5 }]) }));
  assert(goodColor.success, 'valid Color ("Red") is accepted: ' + goodColor.message);
  const redLotNumber = goodColor.data && goodColor.data.lotNumber;
  assert(!!redLotNumber, 'lot number generated (got "' + redLotNumber + '")');

  const allLots = getProductionData();
  assert(allLots.success, 'getProductionData succeeds');
  const redLot = (allLots.data || []).find(l => l.lotNumber === redLotNumber);
  assert(!!redLot, 'saved lot found by lot number');
  assert(redLot.color === 'Red', 'saved lot carries Color "Red" (got "' + (redLot && redLot.color) + '")');

  const prodSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PRODUCTION);
  assert(prodSheet.getLastColumn() >= PRODUCTION_COL.COLOR, 'Production sheet has the Color column (got ' + prodSheet.getLastColumn() + ' cols)');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 5: a process with no color sub-groups never requires Color
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 5: saveProduction does not require Color for a Common-only process ===');
{
  const res = saveProduction({
    processId: packingId,
    qty: 3,
    assignedTo: 'Test Contractor',
    status: 'Pending',
    componentsConsumed: JSON.stringify([
      { itemName: 'Packing Box', sourceType: 'ITEM', qty: 3 }
    ])
    // no color field at all
  });
  assert(res.success, 'lot for a Common-only process saves without Color: ' + res.message);

  const allLots = getProductionData();
  const lot = (allLots.data || []).find(l => l.lotNumber === (res.data && res.data.lotNumber));
  assert(!!lot, 'packing lot found');
  assert(lot.color === '', 'packing lot has blank Color (got "' + (lot && lot.color) + '")');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 6: editing a lot preserves its Color
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 6: editing a lot preserves its Color ===');
{
  const allLots = getProductionData();
  const redLot = allLots.data.find(l => l.color === 'Red');
  assert(!!redLot, 'red lot present for edit test');

  const editRes = saveProduction({
    rowIdx: redLot.rowIdx,
    processId: framePaintingId,
    qty: 7, // bump quantity
    assignedTo: 'Test Contractor',
    status: 'Completed',
    colorBreakdown: JSON.stringify([{ color: 'Red', qty: 7 }]),
    componentsConsumed: JSON.stringify([
      { itemName: 'Brush', sourceType: 'ITEM', qty: 7 },
      { itemName: 'Red Paint', sourceType: 'ITEM', qty: 14 }
    ])
  });
  assert(editRes.success, 'edit saves successfully: ' + editRes.message);

  const after = getProductionData();
  const updated = after.data.find(l => l.rowIdx === redLot.rowIdx);
  assert(!!updated, 'edited row still found');
  assert(updated.qty === 7, 'qty updated to 7 (got ' + (updated && updated.qty) + ')');
  assert(updated.color === 'Red', 'color preserved as Red after edit (got "' + (updated && updated.color) + '")');
}

// ─────────────────────────────────────────────────────────────────────────
console.log('\n' + (failures === 0 ? 'ALL TESTS PASSED' : failures + ' TEST(S) FAILED'));
process.exit(failures === 0 ? 0 : 1);
