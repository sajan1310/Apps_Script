/**
 * Standalone Node harness (same mock-SpreadsheetApp pattern as
 * test_production_pool_and_unit_fixes.js) covering a deep-dive fix from the
 * 2026-07-18 calculations audit: a Process Component recipe row's Color
 * Sub-Group (colorGroup) is configured manually from Color Master (see
 * Script_Process.html's addColorGroup) — independently of whatever literal
 * string the upstream item's own Warehouse Pool credit actually landed
 * under. When an upstream item is produced under a COMPOSITE color (2+
 * independent pool axes cross-multiplied into one string, e.g.
 * "BCP / Blue-White" — see COLOR_COMBO_DELIMITER in config.js, this occurs
 * whenever that upstream process itself has no Primary Color Axis
 * configured), a downstream recipe row scoped to just ONE token of that
 * composite (e.g. "BCP") used to create/debit a phantom "bcp" bucket that
 * was never credited — going negative — while the real composite bucket
 * never got debited and stayed permanently over-available.
 *
 * Fix: module_warehouse.js#_resolveCompositeColorToken resolves a
 * single-token colorGroup to its one matching composite bucket ONLY when
 * unambiguous (exactly one live composite bucket for that item contains
 * the token); a token shared by 2+ composite buckets is left unresolved
 * (falls back to the old phantom-bucket behavior) rather than guessed.
 * Applied to both recalculateWarehousePool's Pass 2 debit and
 * getPoolAvailableQtyMap's byColor (so pre-save validation in
 * module_production.js#_validatePoolAvailability agrees with what Pass 2
 * will actually do).
 *
 * Run: node .pw-test/test_pool_composite_color_token_resolution.js
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
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {} }) },
  console,
  Logger: { log: () => {} },
  Utilities: { getUuid: () => 'uuid-' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);

['config.js', 'utils.js', 'module_units.js', 'module_items.js', 'module_process.js', 'module_production.js', 'module_warehouse.js', 'module_stock.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});

vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.PRODUCTION_COL = PRODUCTION_COL;
`, ctx, { filename: 'expose.js' });

const { APP_CONFIG, PRODUCTION_COL, recalculateWarehousePool, getPoolAvailableQtyMap } = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); }
}

console.log('\n=== Setup: "Painted Frame" credited under a COMPOSITE color, consumed downstream via a single-token Color Sub-Group ===');
{
  const prodSheet = ss.addSheet(APP_CONFIG.SHEETS.PRODUCTION);

  // Upstream lot: Painted Frame process has no Primary Color Axis, so its
  // own multi-axis output checklist offered composite combos — operator
  // picked "BCP / Blue-White" at qty 100.
  // Row 2, not getLastRow()+1 — the sheet is brand new (getLastRow()==0),
  // and data is read starting at row 2 (row 1 reserved for headers).
  let row = 2;
  prodSheet._set(row, PRODUCTION_COL.DATE, '01/01/2026');
  prodSheet._set(row, PRODUCTION_COL.QTY, 100);
  prodSheet._set(row, PRODUCTION_COL.STATUS, 'Completed');
  prodSheet._set(row, PRODUCTION_COL.PROCESS_ID, 'PRC-PAINT');
  prodSheet._set(row, PRODUCTION_COL.LOT_NUMBER, 'LOT-1');
  prodSheet._set(row, PRODUCTION_COL.OUTPUT_ITEM_NAME, 'Painted Frame');
  prodSheet._set(row, PRODUCTION_COL.COMPONENTS_CONSUMED, '[]');
  prodSheet._set(row, PRODUCTION_COL.COLOR_BREAKDOWN, JSON.stringify([{ color: 'BCP / Blue-White', qty: 100 }]));

  // Downstream lot: recipe's Color Sub-Group for the POOL-sourced "Painted
  // Frame" component was manually configured as just "BCP" (a Color Master
  // name), independent of the upstream item's actual composite bucket key.
  row = prodSheet.getLastRow() + 1;
  prodSheet._set(row, PRODUCTION_COL.DATE, '02/01/2026');
  prodSheet._set(row, PRODUCTION_COL.QTY, 10);
  prodSheet._set(row, PRODUCTION_COL.STATUS, 'Completed');
  prodSheet._set(row, PRODUCTION_COL.PROCESS_ID, 'PRC-FIT');
  prodSheet._set(row, PRODUCTION_COL.LOT_NUMBER, 'LOT-2');
  prodSheet._set(row, PRODUCTION_COL.OUTPUT_ITEM_NAME, 'Fitted Frame');
  prodSheet._set(row, PRODUCTION_COL.COMPONENTS_CONSUMED, JSON.stringify([
    { itemName: 'Painted Frame', sourceType: 'POOL', qty: 10, colorGroup: 'BCP' }
  ]));
  prodSheet._set(row, PRODUCTION_COL.COLOR_BREAKDOWN, '');

  const recalc = recalculateWarehousePool();
  assert(recalc.success, 'recalculateWarehousePool succeeds: ' + recalc.message);

  const poolMap = getPoolAvailableQtyMap();
  const entry = poolMap['painted frame'];
  assert(!!entry, 'Painted Frame bucket exists');

  console.log('\n=== Fix: the single-token debit resolved to the real composite bucket, not a phantom "bcp" one ===');
  assert(entry && entry.byColor['bcp / blue-white'] === 90, `composite bucket correctly debited: 100 - 10 = 90 (got ${entry && entry.byColor['bcp / blue-white']})`);
  assert(entry && entry.total === 90, `item total reflects the real debit: 90, not double-counted or phantom-negative (got ${entry && entry.total})`);
}

console.log('\n=== Regression: getPoolAvailableQtyMap.byColor also exposes the token as a resolved alias (pre-save validation agrees with the actual debit) ===');
{
  const poolMap = getPoolAvailableQtyMap();
  const entry = poolMap['painted frame'];
  assert(entry && entry.byColor['bcp'] === 90, `byColor['bcp'] alias resolves to the same composite bucket's balance (got ${entry && entry.byColor['bcp']})`);
}

console.log('\n=== Ambiguous case: a token shared by 2+ composite buckets is left unresolved, not guessed ===');
{
  const prodSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PRODUCTION);

  // A second upstream color combo also containing "BCP".
  let row = prodSheet.getLastRow() + 1;
  prodSheet._set(row, PRODUCTION_COL.DATE, '03/01/2026');
  prodSheet._set(row, PRODUCTION_COL.QTY, 50);
  prodSheet._set(row, PRODUCTION_COL.STATUS, 'Completed');
  prodSheet._set(row, PRODUCTION_COL.PROCESS_ID, 'PRC-PAINT');
  prodSheet._set(row, PRODUCTION_COL.LOT_NUMBER, 'LOT-3');
  prodSheet._set(row, PRODUCTION_COL.OUTPUT_ITEM_NAME, 'Painted Frame');
  prodSheet._set(row, PRODUCTION_COL.COMPONENTS_CONSUMED, '[]');
  prodSheet._set(row, PRODUCTION_COL.COLOR_BREAKDOWN, JSON.stringify([{ color: 'BCP / Orange-White', qty: 50 }]));

  // A new downstream lot consuming "BCP" again -- now ambiguous between
  // "BCP / Blue-White" and "BCP / Orange-White".
  row = prodSheet.getLastRow() + 1;
  prodSheet._set(row, PRODUCTION_COL.DATE, '04/01/2026');
  prodSheet._set(row, PRODUCTION_COL.QTY, 5);
  prodSheet._set(row, PRODUCTION_COL.STATUS, 'Completed');
  prodSheet._set(row, PRODUCTION_COL.PROCESS_ID, 'PRC-FIT');
  prodSheet._set(row, PRODUCTION_COL.LOT_NUMBER, 'LOT-4');
  prodSheet._set(row, PRODUCTION_COL.OUTPUT_ITEM_NAME, 'Fitted Frame 2');
  prodSheet._set(row, PRODUCTION_COL.COMPONENTS_CONSUMED, JSON.stringify([
    { itemName: 'Painted Frame', sourceType: 'POOL', qty: 5, colorGroup: 'BCP' }
  ]));
  prodSheet._set(row, PRODUCTION_COL.COLOR_BREAKDOWN, '');

  const recalc = recalculateWarehousePool();
  assert(recalc.success, 'recalculateWarehousePool succeeds: ' + recalc.message);

  // recalculateWarehousePool fully rebuilds from every row on each call —
  // so this run re-evaluates LOT-2's debit (qty 10) too, which is now ALSO
  // ambiguous (2 composite candidates exist), same as LOT-4's (qty 5).
  // Neither real composite bucket is touched (no wrong guess); both
  // debits fall back to the same phantom "bcp" bucket instead.
  const poolMap = getPoolAvailableQtyMap();
  const entry = poolMap['painted frame'];
  assert(entry && entry.byColor['bcp / blue-white'] === 100, `first composite bucket untouched by either ambiguous debit, stays at its full credit (got ${entry && entry.byColor['bcp / blue-white']})`);
  assert(entry && entry.byColor['bcp / orange-white'] === 50, `second composite bucket also untouched (got ${entry && entry.byColor['bcp / orange-white']})`);
  assert(entry && entry.byColor['bcp'] === -15, `both ambiguous debits (10 + 5) fall back to the same phantom bucket instead of guessing which composite to draw from (got ${entry && entry.byColor['bcp']})`);
}

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
