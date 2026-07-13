/**
 * Standalone Node harness (same mock-SpreadsheetApp pattern as
 * test_bulk_delete_row_guard.js) covering two of the 9 architectural gaps
 * verified+fixed on 2026-07-13 (see verification_2026_07_13_architectural_gaps
 * in project memory):
 *
 * Fix #1 — saveProcess now rejects a second ACTIVE process claiming an
 * Output Item Name already used by another active process, closing the
 * Warehouse Pool bucket-merging ambiguity at its root (module_warehouse.js
 * #_poolKey has no Process ID in its key, by design — see that function's
 * own comment).
 *
 * Fix #2 — Color Master and Process Type Master renames now cascade to
 * every sheet that stores the old name as a plain string reference
 * (Process Components, BOM, Warehouse Pool Opening, Process Color Links,
 * Production's COLOR/COLOR_BREAKDOWN for colors; Process Master for process
 * types) — previously only Vendor/Contractor/Unit had this.
 *
 * Run: node .pw-test/test_process_and_tags_fixes.js
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
  insertColumnsBefore(beforePosition, howMany) {
    this.rows.forEach(row => { const blanks = new Array(howMany).fill(''); row.splice(beforePosition - 1, 0, ...blanks); });
  }
}

class FakeSpreadsheet {
  constructor() { this.sheets = {}; }
  getSheetByName(name) { return this.sheets[name] || null; }
  addSheet(name) { const s = new FakeSheet(name); this.sheets[name] = s; return s; }
  insertSheet(name) { return this.addSheet(name); }
}

const ss = new FakeSpreadsheet();
const cacheStore = {};

const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush: () => {} },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  CacheService: { getScriptCache: () => ({ get: k => (k in cacheStore ? cacheStore[k] : null), put: (k, v) => { cacheStore[k] = v; }, remove: k => { delete cacheStore[k]; } }) },
  console,
  Logger: { log: () => {} },
  Utilities: { getUuid: () => 'uuid-' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);

['config.js', 'utils.js', 'module_units.js', 'module_process.js', 'module_tags.js', 'module_warehouse.js', 'module_production.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});

vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.PROCESS_COL = PROCESS_COL;
  global.PROCESS_COMPONENTS_COL = PROCESS_COMPONENTS_COL;
  global.BOM_COL = BOM_COL;
  global.WAREHOUSE_POOL_OPENING_COL = WAREHOUSE_POOL_OPENING_COL;
  global.PROCESS_COLOR_LINKS_COL = PROCESS_COLOR_LINKS_COL;
  global.PRODUCTION_COL = PRODUCTION_COL;
  global.TAG_COL = TAG_COL;
`, ctx, { filename: 'expose.js' });

const {
  APP_CONFIG, PROCESS_COL, PROCESS_COMPONENTS_COL, BOM_COL, WAREHOUSE_POOL_OPENING_COL,
  PROCESS_COLOR_LINKS_COL, PRODUCTION_COL,
  saveProcess, saveColor, saveProcessType
} = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); }
}

console.log('\n=== Fix #1: Output Item Name uniqueness across ACTIVE processes ===');
{
  // Lot prefixes/output names deliberately avoid initProcessMasterSheet's
  // own auto-seed defaults (FP/RA/FF/PK, "Painted Frame"/"Fitted Rim"/
  // "Fitted Frame"/"Packed Bicycle" — see that function) so this test's
  // own rows don't collide with them.
  const res1 = saveProcess({
    processName: 'Test Process A', sequence: 11, lotPrefix: 'ZZ1',
    outputItemName: 'Test Widget', isFinalStage: false, active: true, remarks: '', components: '[]'
  });
  assert(res1.success, 'first process with "Test Widget" saves fine: ' + res1.message);

  const res2 = saveProcess({
    processName: 'Test Process B', sequence: 12, lotPrefix: 'ZZ2',
    outputItemName: 'Test Widget', isFinalStage: false, active: true, remarks: '', components: '[]'
  });
  assert(res2.success === false, 'a SECOND active process claiming the same Output Item Name is rejected: ' + res2.message);
  assert(/already used by another active process/i.test(res2.message || ''), `rejection message explains why (got "${res2.message}")`);

  const res3 = saveProcess({
    processName: 'Test Process C', sequence: 13, lotPrefix: 'ZZ3',
    outputItemName: 'Test Widget', isFinalStage: false, active: false, remarks: '', components: '[]'
  });
  assert(res3.success, 'an INACTIVE process may still share the name (no new active ambiguity created): ' + res3.message);

  const res4 = saveProcess({
    processName: 'Test Process D', sequence: 14, lotPrefix: 'ZZ4',
    outputItemName: 'Test Widget Two', isFinalStage: false, active: true, remarks: '', components: '[]'
  });
  assert(res4.success, 'a distinct Output Item Name is unaffected: ' + res4.message);
}

console.log('\n=== Fix #2: Color Master rename cascades to Process Components, BOM, Warehouse Pool Opening, Process Color Links, Production ===');
{
  // Seed Process Components: one COMMON-scoped ITEM row tagged 'Purple'.
  const compSheet = ss.addSheet(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);
  compSheet._set(2, PROCESS_COMPONENTS_COL.PROCESS_ID, 'PRC-1');
  compSheet._set(2, PROCESS_COMPONENTS_COL.ITEM_NAME, 'Paint Tin');
  compSheet._set(2, PROCESS_COMPONENTS_COL.QTY_PER_UNIT, 1);
  compSheet._set(2, PROCESS_COMPONENTS_COL.SOURCE_TYPE, 'ITEM');
  compSheet._set(2, PROCESS_COMPONENTS_COL.COLOR_GROUP, 'Purple');

  // Seed BOM: one row scoped to color 'Purple'.
  const bomSheet = ss.addSheet(APP_CONFIG.SHEETS.BOM);
  bomSheet._set(2, BOM_COL.PRODUCT_ID, 'PRD-1');
  bomSheet._set(2, BOM_COL.PRODUCT_NAME, 'Test Bike');
  bomSheet._set(2, BOM_COL.ITEM_NAME, 'Paint Tin');
  bomSheet._set(2, BOM_COL.QTY_PER_PRODUCT, 1);
  bomSheet._set(2, BOM_COL.COLOR, 'Purple');

  // Seed Warehouse Pool Opening: a balance recorded under 'Purple'.
  const openingSheet = ss.addSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
  openingSheet._set(2, WAREHOUSE_POOL_OPENING_COL.OUTPUT_ITEM_NAME, 'Painted Frame');
  openingSheet._set(2, WAREHOUSE_POOL_OPENING_COL.COLOR, 'Purple');
  openingSheet._set(2, WAREHOUSE_POOL_OPENING_COL.QTY, 5);

  // Seed Process Color Links: 'Purple' in BOTH the A and B slot on different rows.
  const linksSheet = ss.addSheet(APP_CONFIG.SHEETS.PROCESS_COLOR_LINKS);
  linksSheet._set(2, PROCESS_COLOR_LINKS_COL.PROCESS_A_ID, 'PRC-1');
  linksSheet._set(2, PROCESS_COLOR_LINKS_COL.COLOR_A, 'Purple');
  linksSheet._set(2, PROCESS_COLOR_LINKS_COL.PROCESS_B_ID, 'PRC-2');
  linksSheet._set(2, PROCESS_COLOR_LINKS_COL.COLOR_B, 'Navy');
  linksSheet._set(3, PROCESS_COLOR_LINKS_COL.PROCESS_A_ID, 'PRC-3');
  linksSheet._set(3, PROCESS_COLOR_LINKS_COL.COLOR_A, 'Teal');
  linksSheet._set(3, PROCESS_COLOR_LINKS_COL.PROCESS_B_ID, 'PRC-1');
  linksSheet._set(3, PROCESS_COLOR_LINKS_COL.COLOR_B, 'Purple');

  // Seed Production: a completed lot with a comma-joined COLOR display
  // string and a COLOR_BREAKDOWN JSON array, both referencing 'Purple'.
  const prodSheet = ss.addSheet(APP_CONFIG.SHEETS.PRODUCTION);
  prodSheet._set(2, PRODUCTION_COL.DATE, '01/01/2026');
  prodSheet._set(2, PRODUCTION_COL.QTY, 10);
  prodSheet._set(2, PRODUCTION_COL.STATUS, 'Completed');
  prodSheet._set(2, PRODUCTION_COL.PROCESS_ID, 'PRC-1');
  prodSheet._set(2, PRODUCTION_COL.LOT_NUMBER, 'LOT-1');
  prodSheet._set(2, PRODUCTION_COL.OUTPUT_ITEM_NAME, 'Painted Frame');
  prodSheet._set(2, PRODUCTION_COL.COMPONENTS_CONSUMED, '[]');
  prodSheet._set(2, PRODUCTION_COL.COLOR, 'Purple, Navy');
  prodSheet._set(2, PRODUCTION_COL.COLOR_BREAKDOWN, JSON.stringify([{ color: 'Purple', qty: 6 }, { color: 'Navy', qty: 4 }]));

  const createRes = saveColor({ name: 'Purple' });
  assert(createRes.success, 'Color Master entry "Purple" created first: ' + createRes.message);

  const res = saveColor({ name: 'Deep Purple', originalName: 'Purple' });
  assert(res.success, 'renaming Purple -> Deep Purple succeeds: ' + res.message);

  assert(String(compSheet._get(2, PROCESS_COMPONENTS_COL.COLOR_GROUP)) === 'Deep Purple', 'Process Components COLOR_GROUP cascaded');
  assert(String(bomSheet._get(2, BOM_COL.COLOR)) === 'Deep Purple', 'BOM COLOR cascaded');
  assert(String(openingSheet._get(2, WAREHOUSE_POOL_OPENING_COL.COLOR)) === 'Deep Purple', 'Warehouse Pool Opening COLOR cascaded');
  assert(String(linksSheet._get(2, PROCESS_COLOR_LINKS_COL.COLOR_A)) === 'Deep Purple', 'Process Color Links COLOR_A slot cascaded');
  assert(String(linksSheet._get(3, PROCESS_COLOR_LINKS_COL.COLOR_B)) === 'Deep Purple', 'Process Color Links COLOR_B slot cascaded (either slot can hold a given color)');
  assert(String(prodSheet._get(2, PRODUCTION_COL.COLOR)) === 'Deep Purple, Navy', `Production's comma-joined COLOR display string cascaded (got "${prodSheet._get(2, PRODUCTION_COL.COLOR)}")`);

  const breakdown = JSON.parse(String(prodSheet._get(2, PRODUCTION_COL.COLOR_BREAKDOWN)));
  assert(breakdown.find(e => e.color === 'Deep Purple' && e.qty === 6) !== undefined, 'Production COLOR_BREAKDOWN JSON entry renamed with qty preserved');
  assert(breakdown.find(e => e.color === 'Navy') !== undefined, "Production COLOR_BREAKDOWN's other (non-matching) entry untouched");
}

console.log('\n=== Fix #2: Process Type Master rename cascades to Process Master ===');
{
  const procSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PROCESS_MASTER);
  // saveProcess already wrote rows above; tag one with a Process Type.
  procSheet._set(2, PROCESS_COL.PROCESS_TYPE, 'Painting');
  procSheet._set(3, PROCESS_COL.PROCESS_TYPE, 'Painting');
  procSheet._set(4, PROCESS_COL.PROCESS_TYPE, 'Welding');

  const createRes = saveProcessType({ name: 'Painting' });
  assert(createRes.success, 'Process Type Master entry "Painting" created first: ' + createRes.message);

  const res = saveProcessType({ name: 'Coating', originalName: 'Painting' });
  assert(res.success, 'renaming Painting -> Coating succeeds: ' + res.message);
  assert(String(procSheet._get(2, PROCESS_COL.PROCESS_TYPE)) === 'Coating', 'row 2 PROCESS_TYPE cascaded');
  assert(String(procSheet._get(3, PROCESS_COL.PROCESS_TYPE)) === 'Coating', 'row 3 PROCESS_TYPE cascaded');
  assert(String(procSheet._get(4, PROCESS_COL.PROCESS_TYPE)) === 'Welding', 'unrelated row (Welding) left untouched');
}

console.log('\n=== Fix #2: a casing-only rename still cascades (matches the fixed Vendor/Unit convention) ===');
{
  const compSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);
  compSheet._set(3, PROCESS_COMPONENTS_COL.PROCESS_ID, 'PRC-2');
  compSheet._set(3, PROCESS_COMPONENTS_COL.ITEM_NAME, 'Thinner');
  compSheet._set(3, PROCESS_COMPONENTS_COL.QTY_PER_UNIT, 1);
  compSheet._set(3, PROCESS_COMPONENTS_COL.SOURCE_TYPE, 'ITEM');
  compSheet._set(3, PROCESS_COMPONENTS_COL.COLOR_GROUP, 'teal');

  const createRes = saveColor({ name: 'teal' });
  assert(createRes.success, 'Color Master entry "teal" created first: ' + createRes.message);

  const res = saveColor({ name: 'TEAL', originalName: 'teal' });
  assert(res.success, 'casing-only rename (teal -> TEAL) succeeds: ' + res.message);
  assert(String(compSheet._get(3, PROCESS_COMPONENTS_COL.COLOR_GROUP)) === 'TEAL', 'casing-only rename still cascaded (plain string compare, not case-insensitive gate)');
}

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
