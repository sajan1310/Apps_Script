/**
 * Verifies repairMissingPoolColorConsumption:
 * 1. Detects a real gap (colorBreakdown has a color, componentsConsumed
 *    has NO matching POOL entry for that color) and backfills it correctly
 *    using recipe qtyPerUnit x breakdown qty.
 * 2. Does NOT flag/backfill a lot with no gap (already has a matching entry).
 * 3. Does NOT produce a false positive when a lot's breakdown includes a
 *    color from a DIFFERENT pool item's own axis (e.g. Mudguard's own
 *    color alongside Rim's) -- must not expect Rim to have an entry for
 *    Mudguard's color.
 * 4. Correctly handles a composite breakdown color ("BCP / Blue-White"),
 *    splitting into tokens and only backfilling the token that actually
 *    belongs to this item.
 * 5. dryRun writes nothing; dryRun:false actually patches the sheet and
 *    triggers recalculateWarehousePool.
 *
 * Run: node .pw-test/test_repair_missing_pool_color_consumption.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = 'c:\\Users\\erkar\\my-app-script-project';

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) { this.sheet = sheet; this.row = row; this.col = col; this.numRows = numRows; this.numCols = numCols; }
  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) { const rowArr = []; for (let c = 0; c < this.numCols; c++) rowArr.push(this.sheet._get(this.row + r, this.col + c)); out.push(rowArr); }
    return out;
  }
  getValue() { return this.sheet._get(this.row, this.col); }
  setValues(values) { values.forEach((rowArr, r) => rowArr.forEach((val, c) => this.sheet._set(this.row + r, this.col + c, val))); return this; }
  setValue(v) { this.sheet._set(this.row, this.col, v); return this; }
  clearContent() { for (let r = 0; r < this.numRows; r++) for (let c = 0; c < this.numCols; c++) this.sheet._set(this.row + r, this.col + c, ''); return this; }
  setFontWeight() { return this; } setBackground() { return this; } setNumberFormat() { return this; }
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
  deleteRow(r) { this.rows.splice(r - 1, 1); } deleteRows(r, n) { this.rows.splice(r - 1, n); }
  insertRows(r, n) { for (let i = 0; i < n; i++) this.rows.splice(r - 1, 0, []); }
  insertColumnsAfter(afterPosition, howMany) { this.rows.forEach(row => { const blanks = new Array(howMany).fill(''); row.splice(afterPosition, 0, ...blanks); }); }
}
class FakeSpreadsheet { constructor() { this.sheets = {}; } getSheetByName(name) { return this.sheets[name] || null; } addSheet(name) { const s = new FakeSheet(name); this.sheets[name] = s; return s; } insertSheet(name) { return this.addSheet(name); } }
const ss = new FakeSpreadsheet();
const logs = [];
const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush: () => {} },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
  console, Logger: { log: (...a) => { logs.push(a.join(' ')); } },
  Utilities: { getUuid: () => 'uuid-' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);
['config.js', 'utils.js', 'module_units.js', 'module_items.js', 'module_process.js', 'module_production.js', 'module_warehouse.js', 'module_stock.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});
vm.runInContext('global.APP_CONFIG=APP_CONFIG; global.PRODUCTION_COL=PRODUCTION_COL; global.PROCESS_COL=PROCESS_COL; global.PROCESS_COMPONENTS_COL=PROCESS_COMPONENTS_COL; global.WAREHOUSE_POOL_COL=WAREHOUSE_POOL_COL;', ctx, { filename: 'expose.js' });
const { APP_CONFIG, PRODUCTION_COL, PROCESS_COL, PROCESS_COMPONENTS_COL, repairMissingPoolColorConsumption, getWarehousePoolData, recalculateWarehousePool } = ctx;

let failures = 0;
function assert(cond, msg) { if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); } }

// --- Setup: PRC-FIT process, recipe with 2 pool axes: Rim (BCP/Black) and Mudguard (Blue/Red) ---
const processSheet = ss.addSheet(APP_CONFIG.SHEETS.PROCESS_MASTER);
processSheet._set(2, PROCESS_COL.PROCESS_ID, 'PRC-FIT');
processSheet._set(2, PROCESS_COL.PROCESS_NAME, 'Fitting Frame');
processSheet._set(2, PROCESS_COL.SEQUENCE, 5);
processSheet._set(2, PROCESS_COL.ACTIVE, true);
processSheet._set(2, PROCESS_COL.OUTPUT_ITEM_NAME, 'Fitted Frame Crysta 16 inch');

const compSheet = ss.addSheet(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);
compSheet._set(2, PROCESS_COMPONENTS_COL.PROCESS_ID, 'PRC-FIT');
compSheet._set(2, PROCESS_COMPONENTS_COL.ITEM_NAME, 'Rim');
compSheet._set(2, PROCESS_COMPONENTS_COL.SOURCE_TYPE, 'POOL');
compSheet._set(2, PROCESS_COMPONENTS_COL.COLOR_GROUP, 'COMMON');
compSheet._set(2, PROCESS_COMPONENTS_COL.QTY_PER_UNIT, 1);
compSheet._set(3, PROCESS_COMPONENTS_COL.PROCESS_ID, 'PRC-FIT');
compSheet._set(3, PROCESS_COMPONENTS_COL.ITEM_NAME, 'Mudguard');
compSheet._set(3, PROCESS_COMPONENTS_COL.SOURCE_TYPE, 'POOL');
compSheet._set(3, PROCESS_COMPONENTS_COL.COLOR_GROUP, 'COMMON');
compSheet._set(3, PROCESS_COMPONENTS_COL.QTY_PER_UNIT, 2);

const prodSheet = ss.addSheet(APP_CONFIG.SHEETS.PRODUCTION);
function setLot(row, { qty, lotNumber, colorBreakdown, componentsConsumed }) {
  prodSheet._set(row, PRODUCTION_COL.DATE, '01/01/2026');
  prodSheet._set(row, PRODUCTION_COL.QTY, qty);
  prodSheet._set(row, PRODUCTION_COL.STATUS, 'Completed');
  prodSheet._set(row, PRODUCTION_COL.PROCESS_ID, 'PRC-FIT');
  prodSheet._set(row, PRODUCTION_COL.LOT_NUMBER, lotNumber);
  prodSheet._set(row, PRODUCTION_COL.OUTPUT_ITEM_NAME, 'Fitted Frame Crysta 16 inch');
  prodSheet._set(row, PRODUCTION_COL.COLOR_BREAKDOWN, JSON.stringify(colorBreakdown));
  prodSheet._set(row, PRODUCTION_COL.COMPONENTS_CONSUMED, JSON.stringify(componentsConsumed));
}

// Upstream Rim (BCP/Black) and Mudguard (Blue/Red) pool credit, so both are detected as pool-color-aware.
setLot(2, { qty: 100, lotNumber: 'RIM-SEED-1', colorBreakdown: [{ color: 'BCP', qty: 100 }], componentsConsumed: [] });
setLot(3, { qty: 100, lotNumber: 'RIM-SEED-2', colorBreakdown: [{ color: 'Black', qty: 100 }], componentsConsumed: [] });
processSheet._set(3, PROCESS_COL.PROCESS_ID, 'PRC-RIM'); // give seed lots their own upstream process id (unused beyond crediting)
prodSheet._set(2, PRODUCTION_COL.PROCESS_ID, 'PRC-RIM'); prodSheet._set(2, PRODUCTION_COL.OUTPUT_ITEM_NAME, 'Rim');
prodSheet._set(3, PRODUCTION_COL.PROCESS_ID, 'PRC-RIM'); prodSheet._set(3, PRODUCTION_COL.OUTPUT_ITEM_NAME, 'Rim');
setLot(4, { qty: 100, lotNumber: 'MUD-SEED-1', colorBreakdown: [{ color: 'Blue', qty: 100 }], componentsConsumed: [] });
setLot(5, { qty: 100, lotNumber: 'MUD-SEED-2', colorBreakdown: [{ color: 'Red', qty: 100 }], componentsConsumed: [] });
prodSheet._set(4, PRODUCTION_COL.PROCESS_ID, 'PRC-MUD'); prodSheet._set(4, PRODUCTION_COL.OUTPUT_ITEM_NAME, 'Mudguard');
prodSheet._set(5, PRODUCTION_COL.PROCESS_ID, 'PRC-MUD'); prodSheet._set(5, PRODUCTION_COL.OUTPUT_ITEM_NAME, 'Mudguard');

// Lot A (row 6): the REAL bug -- colorBreakdown says BCP qty=40, but
// componentsConsumed has NO Rim/BCP entry at all (silently dropped).
setLot(6, {
  qty: 40, lotNumber: 'FIT-0001',
  colorBreakdown: [{ color: 'BCP', qty: 40 }],
  componentsConsumed: [] // gap: no Rim entry for BCP at all
});

// Lot B (row 7): NO gap -- already has a correct Rim/Black entry, must not be touched.
setLot(7, {
  qty: 20, lotNumber: 'FIT-0002',
  colorBreakdown: [{ color: 'Black', qty: 20 }],
  componentsConsumed: [{ itemName: 'Rim', sourceType: 'POOL', qty: 20, colorGroup: 'Black' }]
});

// Lot C (row 8): composite breakdown "BCP / Blue" (Rim=BCP paired with Mudguard=Blue on
// one lot) -- Rim has an entry for BCP, but Mudguard's Blue entry is MISSING.
setLot(8, {
  qty: 15, lotNumber: 'FIT-0003',
  colorBreakdown: [{ color: 'BCP / Blue', qty: 15 }],
  componentsConsumed: [{ itemName: 'Rim', sourceType: 'POOL', qty: 15, colorGroup: 'BCP' }]
});

console.log('\n=== Seed the Warehouse Pool sheet from the seed lots above ===');
recalculateWarehousePool();

console.log('\n=== DRY RUN ===');
repairMissingPoolColorConsumption(true);
console.log(logs.join('\n'));

const dryRunLog = logs.join('\n');
assert(dryRunLog.includes('FIT-0001'), 'Lot FIT-0001 (real gap) is flagged in dry run');
assert(!dryRunLog.includes('FIT-0002'), 'Lot FIT-0002 (no gap) is NOT flagged');
assert(dryRunLog.includes('FIT-0003'), 'Lot FIT-0003 (composite color gap) is flagged in dry run');

// Confirm the FIT-0001 backfill entry is Rim/BCP qty=40 (qtyPerUnit=1 x 40), not touched by Mudguard at all.
const fit1Match = dryRunLog.match(/FIT-0001.*missing entr[^:]*: (\[.*\])/);
assert(!!fit1Match, 'FIT-0001 missing-entries JSON found in log');
if (fit1Match) {
  const missing = JSON.parse(fit1Match[1]);
  assert(missing.length === 1, `FIT-0001 has exactly 1 missing entry (got ${missing.length})`);
  assert(missing[0].itemName === 'Rim' && missing[0].colorGroup === 'BCP' && missing[0].qty === 40,
    `FIT-0001 missing entry is Rim/BCP qty=40 (got ${JSON.stringify(missing[0])})`);
}

// Confirm FIT-0003's missing entry is Mudguard/Blue qty=30 (qtyPerUnit=2 x 15), NOT a Rim/Blue false positive.
const fit3Match = dryRunLog.match(/FIT-0003.*missing entr[^:]*: (\[.*\])/);
assert(!!fit3Match, 'FIT-0003 missing-entries JSON found in log');
if (fit3Match) {
  const missing = JSON.parse(fit3Match[1]);
  assert(missing.length === 1, `FIT-0003 has exactly 1 missing entry (got ${missing.length})`);
  assert(missing[0].itemName === 'Mudguard' && missing[0].colorGroup === 'Blue' && missing[0].qty === 30,
    `FIT-0003 missing entry is Mudguard/Blue qty=30, i.e. 2 x 15 (got ${JSON.stringify(missing[0])})`);
}

// Confirm dry run wrote NOTHING to the sheet.
const fit1RawAfterDryRun = prodSheet._get(6, PRODUCTION_COL.COMPONENTS_CONSUMED);
assert(fit1RawAfterDryRun === '[]', `Dry run did not modify FIT-0001's stored componentsConsumed (got ${fit1RawAfterDryRun})`);

console.log('\n=== APPLY (dryRun: false) ===');
logs.length = 0;
repairMissingPoolColorConsumption(false);
console.log(logs.join('\n'));

const fit1RawAfterApply = JSON.parse(prodSheet._get(6, PRODUCTION_COL.COMPONENTS_CONSUMED));
assert(fit1RawAfterApply.length === 1 && fit1RawAfterApply[0].itemName === 'Rim' && fit1RawAfterApply[0].colorGroup === 'BCP' && fit1RawAfterApply[0].qty === 40,
  `Apply mode actually wrote the backfilled entry to the sheet (got ${JSON.stringify(fit1RawAfterApply)})`);

const poolAfter = getWarehousePoolData().data;
const rimBcp = poolAfter.find(r => r.outputItemName.toLowerCase() === 'rim' && r.color === 'BCP');
console.log('Rim/BCP bucket after repair+recalculate:', JSON.stringify(rimBcp));
// 40 (FIT-0001's newly-backfilled entry) + 15 (FIT-0003's own pre-existing Rim/BCP entry) = 55.
assert(!!rimBcp && rimBcp.consumedQty === 55, `Rim/BCP bucket now correctly shows consumedQty=55 (40 backfilled + 15 pre-existing) after repair (got ${rimBcp && rimBcp.consumedQty})`);

const mudguardBlue = poolAfter.find(r => r.outputItemName.toLowerCase() === 'mudguard' && r.color === 'Blue');
console.log('Mudguard/Blue bucket after repair+recalculate:', JSON.stringify(mudguardBlue));
assert(!!mudguardBlue && mudguardBlue.consumedQty === 30, `Mudguard/Blue bucket now correctly shows consumedQty=30 (FIT-0003's backfilled entry) (got ${mudguardBlue && mudguardBlue.consumedQty})`);

console.log('\n=== Re-run dry run: should now be idempotent (no more gaps for FIT-0001/FIT-0003) ===');
logs.length = 0;
repairMissingPoolColorConsumption(true);
console.log(logs.join('\n'));
const secondDryRun = logs.join('\n');
assert(!secondDryRun.includes('FIT-0001') && !secondDryRun.includes('FIT-0003'), 'Second dry run finds no more gaps for the lots just fixed (idempotent)');

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
