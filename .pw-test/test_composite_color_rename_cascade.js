/**
 * Standalone Node harness (same mock-SpreadsheetApp pattern as
 * test_process_and_tags_fixes.js) covering the composite-token gap in
 * module_tags.js#_renameColorEverywhere found while auditing the Warehouse
 * Pool color-combination algorithm (2026-07-22): a color cell isn't always a
 * single literal name — computeColorAxesForProcess/_mergeLinkedAxes (Process
 * Color Links) can produce a COMPOSITE value joining 2+ independent axes
 * with COLOR_COMBO_DELIMITER (" / "), e.g. "BCP / Blue-White". The old
 * cascade did a whole-cell exact-string compare, so renaming "Blue-White"
 * left every composite bucket referencing it (Production's COLOR_BREAKDOWN,
 * a Process Component's colorGroup, a Process Color Links row) silently
 * stale — and since recalculateWarehousePool() rebuilds Warehouse Pool
 * bucket keys straight from that same un-renamed data, the stale composite
 * bucket persisted through every future rebuild too.
 *
 * Run: node .pw-test/test_composite_color_rename_cascade.js
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

['config.js', 'utils.js', 'module_units.js', 'module_process.js', 'module_tags.js', 'module_warehouse.js', 'module_production.js']
  .forEach(f => vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f }));

vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.COLOR_COMBO_DELIMITER = COLOR_COMBO_DELIMITER;
  global.COMPONENT_SOURCE_TYPES = COMPONENT_SOURCE_TYPES;
  global.COMPONENT_COLOR_GROUP_COMMON = COMPONENT_COLOR_GROUP_COMMON;
  global.PROCESS_COLOR_LINKS_COL = PROCESS_COLOR_LINKS_COL;
`, ctx, { filename: 'expose.js' });

const {
  APP_CONFIG, COLOR_COMBO_DELIMITER, COMPONENT_SOURCE_TYPES, COMPONENT_COLOR_GROUP_COMMON, PROCESS_COLOR_LINKS_COL,
  saveColor, saveProcess, saveProduction, getProductionData, recalculateWarehousePool, getWarehousePoolData
} = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); }
}

console.log('\n=== Setup: Color Master entries + a process producing a composite ("BCP / Blue-White") lot ===');
assert(saveColor({ name: 'BCP', remarks: '' }).success, 'Color Master "BCP" created');
assert(saveColor({ name: 'Blue-White', remarks: '' }).success, 'Color Master "Blue-White" created');

const procRes = saveProcess({
  processName: 'Composite Color Process',
  sequence: 1,
  lotPrefix: 'CCP',
  outputItemName: 'Composite Output',
  isFinalStage: false,
  active: true,
  remarks: '',
  components: JSON.stringify([
    // A legacy-style cross-multiplied recipe row: colorGroup itself is a
    // composite string (same shape test_color_axes.js Test 7 exercises).
    { itemName: 'Composite-Colored Part', sourceType: COMPONENT_SOURCE_TYPES.ITEM, qtyPerUnit: 1, colorGroup: `BCP${COLOR_COMBO_DELIMITER}Blue-White` }
  ])
});
assert(procRes.success, 'saveProcess succeeds: ' + procRes.message);
const processId = procRes.data && procRes.data.processId;
assert(!!processId, 'processId returned');

const lotRes = saveProduction({
  processId,
  assignedTo: 'Test Contractor',
  status: 'Completed',
  colorBreakdown: JSON.stringify([
    { color: `BCP${COLOR_COMBO_DELIMITER}Blue-White`, qty: 5, isCustom: true },
    { color: 'BCP', qty: 3, isCustom: true } // a plain (non-composite) sibling entry - must stay untouched by the coming rename
  ]),
  componentsConsumed: JSON.stringify([
    { itemName: 'Composite-Colored Part', sourceType: COMPONENT_SOURCE_TYPES.ITEM, qty: 8, colorGroup: `BCP${COLOR_COMBO_DELIMITER}Blue-White` }
  ])
});
assert(lotRes.success, 'saveProduction (Completed, composite colorBreakdown) succeeds: ' + lotRes.message);
const lotNumber = lotRes.data && lotRes.data.lotNumber;
assert(!!lotNumber, 'lot number generated');

recalculateWarehousePool();
{
  const poolBefore = (getWarehousePoolData().data || []);
  const compositeBucketBefore = poolBefore.find(r => r.color === `BCP${COLOR_COMBO_DELIMITER}Blue-White`);
  assert(!!compositeBucketBefore && compositeBucketBefore.producedQty === 5, `composite bucket "BCP / Blue-White" exists pre-rename with producedQty 5 (got ${JSON.stringify(compositeBucketBefore)})`);
}

console.log('\n=== Rename "Blue-White" -> "Sky Blue-White" ===');
const renameRes = saveColor({ name: 'Sky Blue-White', remarks: '', originalName: 'Blue-White' });
assert(renameRes.success, 'saveColor rename succeeds: ' + renameRes.message);

console.log('\n=== Test 1: Production COLOR_BREAKDOWN composite entry renamed, sibling non-composite entry untouched ===');
{
  const lot = (getProductionData().data || []).find(l => l.lotNumber === lotNumber);
  assert(!!lot, 'lot found after rename');
  const composite = (lot.colorBreakdown || []).find(c => c.color.includes('Sky Blue-White') || c.color === `BCP${COLOR_COMBO_DELIMITER}Blue-White`);
  assert(!!composite && composite.color === `BCP${COLOR_COMBO_DELIMITER}Sky Blue-White`,
    `composite colorBreakdown entry renamed to "BCP / Sky Blue-White" (got "${composite && composite.color}")`);
  assert(composite.qty === 5, 'renamed entry keeps its qty (5)');
  const plain = (lot.colorBreakdown || []).find(c => c.color === 'BCP');
  assert(!!plain && plain.qty === 3, 'sibling plain "BCP" entry untouched (still "BCP", qty 3)');
}

console.log('\n=== Test 2: Warehouse Pool bucket rebuilds under the renamed composite key, not stuck on the old one ===');
{
  recalculateWarehousePool();
  const pool = (getWarehousePoolData().data || []);
  const stale = pool.find(r => r.color === `BCP${COLOR_COMBO_DELIMITER}Blue-White`);
  const renamed = pool.find(r => r.color === `BCP${COLOR_COMBO_DELIMITER}Sky Blue-White`);
  assert(!stale, `no bucket left under the stale composite key "BCP / Blue-White" (got ${JSON.stringify(stale)})`);
  assert(!!renamed && renamed.producedQty === 5, `bucket now exists under the renamed composite key "BCP / Sky Blue-White" with producedQty 5 (got ${JSON.stringify(renamed)})`);
}

console.log('\n=== Test 3: Process Components colorGroup composite entry renamed ===');
{
  const compSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);
  const lastRow = compSheet.getLastRow();
  const rows = compSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const row = rows.find(r => r[0] === processId);
  assert(!!row, 'component row found');
  assert(row[7] === `BCP${COLOR_COMBO_DELIMITER}Sky Blue-White`, `colorGroup cell renamed to "BCP / Sky Blue-White" (got "${row && row[7]}")`);
}

console.log('\n=== Setup 2: a composite value stored in a Process Color Links row ===');
{
  const secondProc = saveProcess({
    processName: 'Second Process For Link',
    sequence: 2,
    lotPrefix: 'SPL',
    outputItemName: 'Second Output',
    isFinalStage: false,
    active: true,
    remarks: '',
    components: JSON.stringify([])
  });
  const secondId = secondProc.data && secondProc.data.processId;
  assert(!!secondId, 'second process created');

  // Directly seed a Process Color Links row whose COLOR_A is itself already
  // a composite (e.g. a prior 3-way chain) — exercises the same rename gap
  // this fix closes, independent of Production.
  const linksSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PROCESS_COLOR_LINKS) || ss.insertSheet(APP_CONFIG.SHEETS.PROCESS_COLOR_LINKS);
  const startRow = linksSheet.getLastRow() + 1;
  linksSheet.getRange(startRow, 1, 1, 6).setValues([[processId, `BCP${COLOR_COMBO_DELIMITER}Sky Blue-White`, secondId, 'Solo Color', '', '']]);

  assert(saveColor({ name: 'Solo Color', remarks: '' }).success, 'Color Master "Solo Color" created');
  assert(saveColor({ name: 'Renamed Solo Color', remarks: '', originalName: 'Solo Color' }).success, 'unrelated plain-color rename succeeds (sanity)');
  assert(saveColor({ name: 'Sunrise Blue-White', remarks: '', originalName: 'Sky Blue-White' }).success, 'second rename (Sky Blue-White -> Sunrise Blue-White) succeeds');

  const linkRow = linksSheet.getRange(startRow, 1, 1, 6).getValues()[0];
  assert(linkRow[1] === `BCP${COLOR_COMBO_DELIMITER}Sunrise Blue-White`, `Process Color Links COLOR_A composite token renamed (got "${linkRow[1]}")`);
  assert(linkRow[3] === 'Renamed Solo Color', `Process Color Links COLOR_B plain value renamed by its own rename (got "${linkRow[3]}")`);
}

console.log('\n' + (failures === 0 ? 'ALL TESTS PASSED' : failures + ' TEST(S) FAILED'));
process.exit(failures === 0 ? 0 : 1);
