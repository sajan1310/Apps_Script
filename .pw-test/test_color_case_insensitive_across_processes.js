/**
 * Colors are one real thing regardless of casing, everywhere two processes
 * meet: "Blue" / "BLUE" / "blue" / "bLUE" must resolve to the same color
 * whether it arrives from Color Master, a recipe's Color Sub-Group, a
 * Warehouse Pool bucket credited upstream, or a saved lot's Color Breakdown.
 *
 * Most of this chain was already case-insensitive (_poolKey, _colorNamesMatch,
 * _addUniqueCaseInsensitive, saveProduction validation); these tests pin that
 * down so it can't regress, and cover the two gaps that were NOT:
 *   - a Color Sub-Group of "Common"/"common" was compared exactly against the
 *     COMMON sentinel, so it became a phantom color sub-group named "Common"
 *     (and its component row dropped out of Common Components);
 *   - _axisLinkRef lowercased only the axis key, not the Process ID, so a
 *     Process Color Link whose stored Process ID differed in case from the
 *     pool row's silently failed to pair (axes cross-multiplied instead).
 *
 * Run: node .pw-test/test_color_case_insensitive_across_processes.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = 'c:\\Users\\erkar\\my-app-script-project';

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) { this.sheet = sheet; this.row = row; this.col = col; this.numRows = numRows; this.numCols = numCols; }
  getValues() { const out = []; for (let r = 0; r < this.numRows; r++) { const a = []; for (let c = 0; c < this.numCols; c++) a.push(this.sheet._get(this.row + r, this.col + c)); out.push(a); } return out; }
  getValue() { return this.sheet._get(this.row, this.col); }
  setValues(v) { v.forEach((ra, r) => ra.forEach((val, c) => this.sheet._set(this.row + r, this.col + c, val))); return this; }
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
  getLastColumn() { let m = 0; this.rows.forEach(row => { for (let c = row.length; c >= 1; c--) { if (row[c - 1] !== '' && row[c - 1] !== undefined && row[c - 1] !== null) { m = Math.max(m, c); break; } } }); return m; }
  getRange(row, col, numRows = 1, numCols = 1) { return new FakeRange(this, row, col, numRows, numCols); }
  appendRow(arr) { const r = this.getLastRow() + 1; arr.forEach((v, i) => this._set(r, i + 1, v)); }
  deleteRow(r) { this.rows.splice(r - 1, 1); } deleteRows(r, n) { this.rows.splice(r - 1, n); }
  insertRows(r, n) { for (let i = 0; i < n; i++) this.rows.splice(r - 1, 0, []); }
  insertColumnsAfter(a, h) { this.rows.forEach(row => { row.splice(a, 0, ...new Array(h).fill('')); }); }
}
class FakeSpreadsheet { constructor() { this.sheets = {}; } getSheetByName(n) { return this.sheets[n] || null; } addSheet(n) { const s = new FakeSheet(n); this.sheets[n] = s; return s; } insertSheet(n) { return this.addSheet(n); } }
const ss = new FakeSpreadsheet();
const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush: () => {} },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
  console, Logger: { log: () => {} },
  Utilities: { getUuid: () => 'uuid-' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);
['config.js', 'utils.js', 'module_units.js', 'module_items.js', 'module_process.js', 'module_production.js', 'module_warehouse.js', 'module_stock.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});
vm.runInContext('global.APP_CONFIG=APP_CONFIG; global.PRODUCTION_COL=PRODUCTION_COL;', ctx, { filename: 'expose.js' });
const C = ctx;
const { APP_CONFIG, PRODUCTION_COL } = C;

let failures = 0;
function assert(cond, msg) { if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); } }

const prodSheet = ss.addSheet(APP_CONFIG.SHEETS.PRODUCTION);
let nextRow = 2;
function addLot(o) {
  const row = nextRow++;
  prodSheet._set(row, PRODUCTION_COL.DATE, '01/01/2026');
  prodSheet._set(row, PRODUCTION_COL.QTY, o.qty);
  prodSheet._set(row, PRODUCTION_COL.STATUS, 'Completed');
  prodSheet._set(row, PRODUCTION_COL.PROCESS_ID, o.processId);
  prodSheet._set(row, PRODUCTION_COL.LOT_NUMBER, o.lotNumber);
  prodSheet._set(row, PRODUCTION_COL.OUTPUT_ITEM_NAME, o.outputItemName);
  prodSheet._set(row, PRODUCTION_COL.COMPONENTS_CONSUMED, JSON.stringify(o.componentsConsumed || []));
  prodSheet._set(row, PRODUCTION_COL.COLOR_BREAKDOWN, JSON.stringify(o.colorBreakdown));
}
const bucketsFor = (item) => C.getWarehousePoolData().data.filter(r => r.outputItemName === item)
  .map(r => ({ color: r.color, produced: r.producedQty, consumed: r.consumedQty }))
  .sort((a, b) => a.color.localeCompare(b.color));

console.log('=== Test 1: sameColor / isCommonColorGroup helpers ===');
assert(C.sameColor('Blue', 'BLUE') && C.sameColor('bLUE', ' blue ') && C.sameColor('Blue-White', 'BLUE-WHITE'),
  'sameColor ignores case and surrounding whitespace');
assert(!C.sameColor('Blue', 'Blue-White') && !C.sameColor('Blue', ''), 'sameColor stays an exact-name test (not the segment heuristic)');
assert(C.isCommonColorGroup('COMMON') && C.isCommonColorGroup('Common') && C.isCommonColorGroup(' common '),
  'isCommonColorGroup accepts any casing of the sentinel');
assert(!C.isCommonColorGroup('Blue') && !C.isCommonColorGroup(''), 'isCommonColorGroup rejects real colors and blanks');

console.log('\n=== Test 2: same color in different casing credits ONE pool bucket ===');
addLot({ processId: 'PRC-A', qty: 10, lotNumber: 'A1', outputItemName: 'Frame',
  colorBreakdown: [{ color: 'Blue-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:f' },
                   { color: 'Black', qty: 10, countsTowardTotal: false, axisKey: 'pool:r' }] });
addLot({ processId: 'PRC-A', qty: 10, lotNumber: 'A2', outputItemName: 'Frame',
  colorBreakdown: [{ color: 'BLUE-WHITE', qty: 10, countsTowardTotal: true, axisKey: 'pool:f' },
                   { color: 'black', qty: 10, countsTowardTotal: false, axisKey: 'pool:r' }] });
C.recalculateWarehousePool();
let pool = bucketsFor('Frame');
console.log('  ', JSON.stringify(pool));
assert(pool.length === 1 && pool[0].produced === 20,
  `"Blue-White / Black" and "BLUE-WHITE / black" merge into one 20-unit bucket (got ${JSON.stringify(pool)})`);

console.log('\n=== Test 3: mirror axis recognised across casing (Frame "Blue-White" vs Mudguard "BLUE") ===');
prodSheet.rows = []; nextRow = 2;
addLot({ processId: 'PRC-A', qty: 10, lotNumber: 'A3', outputItemName: 'Frame2',
  colorBreakdown: [{ color: 'Blue-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:f' },
                   { color: 'BLUE', qty: 10, countsTowardTotal: false, axisKey: 'pool:mud' },
                   { color: 'bcp', qty: 10, countsTowardTotal: false, axisKey: 'pool:rim' }] });
C.recalculateWarehousePool();
pool = bucketsFor('Frame2');
console.log('  ', JSON.stringify(pool));
assert(pool.length === 1 && pool[0].color === 'Blue-White / bcp',
  `uppercase "BLUE" is still recognised as the Frame's mirror and excluded (got ${JSON.stringify(pool.map(p => p.color))})`);

console.log('\n=== Test 4: a recipe consuming the bucket in a different case debits the SAME bucket ===');
prodSheet.rows = []; nextRow = 2;
addLot({ processId: 'PRC-A', qty: 10, lotNumber: 'A4', outputItemName: 'Frame3',
  colorBreakdown: [{ color: 'Blue-White', qty: 10, countsTowardTotal: true, axisKey: 'pool:f' }] });
addLot({ processId: 'PRC-B', qty: 4, lotNumber: 'B1', outputItemName: 'Cycle',
  colorBreakdown: [{ color: 'Blue-White', qty: 4, countsTowardTotal: true, axisKey: 'pool:frame3' }],
  componentsConsumed: [{ itemName: 'Frame3', sourceType: 'POOL', qty: 4, colorGroup: 'bLUE-wHITE' }] });
C.recalculateWarehousePool();
pool = bucketsFor('Frame3');
console.log('  ', JSON.stringify(pool));
assert(pool.length === 1 && pool[0].consumed === 4,
  `colorGroup "bLUE-wHITE" debits the "Blue-White" bucket, no phantom second bucket (got ${JSON.stringify(pool)})`);

console.log('\n=== Test 5: a differently-cased color passes saveProduction validation ===');
prodSheet.rows = []; nextRow = 2;
addLot({ processId: 'PRC-A', qty: 5, lotNumber: 'A5', outputItemName: 'Frame4',
  colorBreakdown: [{ color: 'Blue', qty: 5, countsTowardTotal: true, axisKey: 'pool:f' }] });
addLot({ processId: 'PRC-A', qty: 5, lotNumber: 'A6', outputItemName: 'Frame4',
  colorBreakdown: [{ color: 'RED', qty: 5, countsTowardTotal: true, axisKey: 'pool:f' }] });
const asmRes = C.saveProcess({ processName: 'Assembly', sequence: 2, lotPrefix: 'ZAS',
  outputItemName: 'Assembled', isFinalStage: false, active: true, remarks: '',
  components: JSON.stringify([{ itemName: 'Frame4', sourceType: 'POOL', qtyPerUnit: 1, colorGroup: 'COMMON' }]) });
assert(asmRes.success, 'Assembly process saved: ' + (asmRes.message || ''));
C.recalculateWarehousePool();
const groups = C.getProcessColorGroups(asmRes.data.processId).data || [];
assert(groups.length === 2, `checklist offers 2 colors, not 4 case-variants (got ${JSON.stringify(groups)})`);
const saveRes = C.saveProduction({ processId: asmRes.data.processId, qty: 2, assignedTo: 'X', status: 'Pending',
  colorBreakdown: JSON.stringify([{ color: 'bLUE', qty: 2 }]),
  componentsConsumed: JSON.stringify([{ itemName: 'Frame4', sourceType: 'POOL', qty: 2, colorGroup: 'bLUE' }]) });
assert(saveRes.success, 'a lot logged as "bLUE" against a pool holding "Blue" is accepted: ' + (saveRes.message || ''));

console.log('\n=== Test 6: Color Sub-Group "Common"/"common" is the sentinel, not a color named Common ===');
const ccRes = C.saveProcess({ processName: 'CaseCommon', sequence: 3, lotPrefix: 'ZCC',
  outputItemName: 'CC Out', isFinalStage: false, active: true, remarks: '',
  components: JSON.stringify([
    { itemName: 'Widget', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Common' },
    { itemName: 'Bolt', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'common' },
    { itemName: 'Nut', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'COMMON' }
  ]) });
assert(ccRes.success, 'process with mixed-case COMMON rows saved: ' + (ccRes.message || ''));
const ccGroups = C.getProcessColorGroups(ccRes.data.processId).data || [];
assert(ccGroups.length === 0,
  `no phantom "Common" color sub-group is invented (got ${JSON.stringify(ccGroups)})`);

console.log('\n=== Test 7: axis identity survives a Process ID casing difference ===');
assert(C._axisLinkRef('PRC-001', 'tag:Rim Color') === C._axisLinkRef('prc-001', 'tag:rim color'),
  'the same axis referenced with different Process ID / axis key casing resolves to one ref');

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
