/**
 * Verifies refreshProductionComponentsFromItemsMaster (module_production.js):
 * the repair pass that rewrites narration, unit, and item name (casing only)
 * STORED on every already-logged Production lot to match Items Master's
 * current values. Supersedes the old, narration-only
 * backfillProductionNarrationFromItems (see verify_production_narration_backfill.js,
 * now retired — its narration coverage is folded in here).
 *
 * narration/unit/name are all item metadata the operator maintains in Items
 * Master, but each lot copies them into its own COMPONENTS_CONSUMED /
 * CUSTOM_COMPONENTS JSON at log time. Every lot logged before a value was
 * set/corrected, or before an item was renamed/re-cased or its Base Unit
 * changed, therefore carries stale data forever unless refreshed.
 *
 * Covered:
 *   - a stale narration is replaced by the current Items Master one
 *   - a BLANK stored narration is filled in
 *   - an item Items Master doesn't know keeps its stored note/unit/name
 *     entirely (not touched at all)
 *   - an item whose Items Master narration is blank keeps its stored note
 *     (even while its unit DOES get corrected — narration and unit are
 *     independent per-field decisions, not all-or-nothing per component)
 *   - itemName is corrected to Items Master's current casing/spelling for
 *     the SAME identity (matched case-insensitively) — never repoints to a
 *     different item
 *   - an EXPLICIT stored unit that no longer matches the item's current
 *     Base Unit is corrected
 *   - a BLANK stored unit is left blank (it already means "this item's Base
 *     Unit" — rewriting it to an explicit string would be pure churn)
 *   - a component already fully in sync (name/narration/unit) is untouched
 *   - the CUSTOM_COMPONENTS ("sheet customization") snapshot is repaired too
 *   - POOL-sourced rows are resolved by name like any other (metadata, not
 *     identity — contrast backfillProductionConsumedItemRefs)
 *   - NOTHING but name/narration/unit changes: qty/color/colorGroup/
 *     sourceType/requiredQty and every other Production column round-trip
 *     untouched
 *   - idempotent: an immediate re-run reports zero changes
 *   - unparseable JSON is skipped, not corrupted
 *   - saveProduction writes narration fresh from Items Master, so a lot saved
 *     with a stale client value can't undo the refresh
 *
 * Uses the STRICT FakeRange from verify_process_save_row_width.js (setValues
 * enforces the range's declared shape exactly as Apps Script does), so a
 * mismatched write width fails here instead of in production.
 *
 * Run: node .pw-test/verify_production_refresh_from_items_master.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = 'c:\\Users\\erkar\\my-app-script-project';

class FakeRange {
  constructor(s, r, c, nr, nc) { this.sheet = s; this.row = r; this.col = c; this.numRows = nr; this.numCols = nc; }
  getValues() { const o = []; for (let r = 0; r < this.numRows; r++) { const a = []; for (let c = 0; c < this.numCols; c++) a.push(this.sheet._get(this.row + r, this.col + c)); o.push(a); } return o; }
  getValue() { return this.sheet._get(this.row, this.col); }
  setValues(v) {
    if (!Array.isArray(v) || v.length !== this.numRows) {
      throw new Error(`The number of rows in the data does not match the number of rows in the range. The data has ${Array.isArray(v) ? v.length : 0} but the range has ${this.numRows}.`);
    }
    v.forEach(ra => {
      if (!Array.isArray(ra) || ra.length !== this.numCols) {
        throw new Error(`The number of columns in the data does not match the number of columns in the range. The data has ${Array.isArray(ra) ? ra.length : 0} but the range has ${this.numCols}.`);
      }
    });
    v.forEach((ra, r) => ra.forEach((val, c) => this.sheet._set(this.row + r, this.col + c, val)));
    return this;
  }
  setValue(v) {
    if (this.numRows !== 1 || this.numCols !== 1) { for (let r = 0; r < this.numRows; r++) for (let c = 0; c < this.numCols; c++) this.sheet._set(this.row + r, this.col + c, v); return this; }
    this.sheet._set(this.row, this.col, v); return this;
  }
  clearContent() { for (let r = 0; r < this.numRows; r++) for (let c = 0; c < this.numCols; c++) this.sheet._set(this.row + r, this.col + c, ''); return this; }
  setFontWeight() { return this; } setBackground() { return this; } setNumberFormat() { return this; }
  setHorizontalAlignment() { return this; } setFontColor() { return this; } setWrap() { return this; }
  setFontSize() { return this; } setBorder() { return this; } setValuesUnchecked() { return this; }
}
class FakeSheet {
  constructor(n) { this.name = n; this.rows = []; }
  getName() { return this.name; }
  _ensureRow(r) { while (this.rows.length < r) this.rows.push([]); }
  _get(r, c) { this._ensureRow(r); const w = this.rows[r - 1]; return w[c - 1] === undefined ? '' : w[c - 1]; }
  _set(r, c, v) { this._ensureRow(r); const w = this.rows[r - 1]; while (w.length < c) w.push(''); w[c - 1] = v; }
  getLastRow() { for (let r = this.rows.length; r >= 1; r--) { if (this.rows[r - 1].some(v => v !== '' && v !== undefined && v !== null)) return r; } return 0; }
  getLastColumn() { let m = 0; this.rows.forEach(w => { for (let c = w.length; c >= 1; c--) { if (w[c - 1] !== '' && w[c - 1] !== undefined && w[c - 1] !== null) { m = Math.max(m, c); break; } } }); return m; }
  getRange(r, c, nr = 1, nc = 1) { return new FakeRange(this, r, c, nr, nc); }
  getMaxColumns() { return Math.max(this.getLastColumn(), 30); }
  getMaxRows() { return Math.max(this.rows.length, 100); }
  appendRow(a) { const r = this.getLastRow() + 1; a.forEach((v, i) => this._set(r, i + 1, v)); }
  deleteRow(r) { this.rows.splice(r - 1, 1); } deleteRows(r, n) { this.rows.splice(r - 1, n); }
  insertRows(r, n) { for (let i = 0; i < n; i++) this.rows.splice(r - 1, 0, []); }
  insertColumnsAfter(a, h) { this.rows.forEach(w => { w.splice(a, 0, ...new Array(h).fill('')); }); }
  setFrozenRows() { return this; } setColumnWidth() { return this; } autoResizeColumn() { return this; }
  getFilter() { return null; } clear() { this.rows = []; return this; }
}
class FakeSpreadsheet {
  constructor() { this.sheets = {}; }
  getSheetByName(n) { return this.sheets[n] || null; }
  addSheet(n) { const s = new FakeSheet(n); this.sheets[n] = s; return s; }
  insertSheet(n) { return this.addSheet(n); }
  getSheets() { return Object.keys(this.sheets).map(k => this.sheets[k]); }
}
const ss = new FakeSpreadsheet();
const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush: () => {} },
  // waitLock, not just tryLock — refreshProductionComponentsFromItemsMaster
  // is a top-level entry point and takes the lock itself.
  LockService: { getDocumentLock: () => ({ tryLock: () => true, waitLock: () => true, releaseLock: () => {} }) },
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
  console, Logger: { log: () => {} },
  Utilities: { getUuid: () => 'u' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);
['config.js', 'utils.js', 'module_units.js', 'module_items.js', 'module_process.js', 'module_production.js',
 'module_warehouse.js', 'module_stock.js', 'module_clients.js', 'module_dispatch.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});
vm.runInContext('global.APP_CONFIG=APP_CONFIG; global.PRODUCTION_COL=PRODUCTION_COL; global.ITEMS_COL=ITEMS_COL;',
  ctx, { filename: 'expose.js' });
const C = ctx;
const { APP_CONFIG, PRODUCTION_COL, ITEMS_COL } = C;

let failures = 0;
function assert(cond, msg) { if (!cond) { failures++; console.error('  FAIL:', msg); } else { console.log('  PASS:', msg); } }

// ── Seed Items Master ────────────────────────────────────────────────────
const items = ss.addSheet(APP_CONFIG.SHEETS.ITEMS);
items.appendRow(['Item Name', 'Size', 'Remarks', 'Narration', 'Specification', 'Base Unit', 'Purchase Unit', 'Weight/Base Unit', 'Vendors']);
[
  ['Carton Box 16 inch', '', '', 'Corrugated 5-ply (revised)', '', 'Pcs', 'Pcs', 0, ''],
  ['Adhesive Tape', '', '', '', '', 'Kg', 'Kg', 0, ''],
  ['Frame Sticker---Blue', '', '', 'Sticker kit A', '', 'Set', 'Set', 0, ''],
  ['Bearing 6203', '20 inch', '', 'Sealed, size-specific note', '', 'Pcs', 'Pcs', 0, ''],
  ['Bearing 6203', '', '', 'Generic, no size', '', 'Pcs', 'Pcs', 0, ''],
  ['Fitted Frame 16 inch', '', '', 'WIP frame from Fitting', '', 'Pcs', 'Pcs', 0, ''],
  ['CURVY STICKER Backrest', '', '', 'Sticker note', '', 'Pcs', 'Pcs', 0, ''],
  ['Bolt M6', '', '', 'Bolt note', '', 'Pcs', 'Pcs', 0, '']
].forEach(r => items.appendRow(r));

// ── Seed Production ──────────────────────────────────────────────────────
const prod = ss.addSheet(APP_CONFIG.SHEETS.PRODUCTION);
const header = new Array(PRODUCTION_COL.ORDER_NUMBER).fill('');
header[PRODUCTION_COL.DATE - 1] = 'Date';
header[PRODUCTION_COL.LOT_NUMBER - 1] = 'Lot Number';
header[PRODUCTION_COL.COMPONENTS_CONSUMED - 1] = 'Components Consumed';
prod.appendRow(header);

function lotRow(opts) {
  const row = new Array(PRODUCTION_COL.ORDER_NUMBER).fill('');
  row[PRODUCTION_COL.DATE - 1] = '29/07/2026';
  row[PRODUCTION_COL.QTY - 1] = opts.qty;
  row[PRODUCTION_COL.ASSIGNED_TO - 1] = 'Sanjay';
  row[PRODUCTION_COL.STATUS - 1] = 'Completed';
  row[PRODUCTION_COL.REMARKS - 1] = 'do not touch me';
  row[PRODUCTION_COL.PROCESS_ID - 1] = 'PRC-1151';
  row[PRODUCTION_COL.LOT_NUMBER - 1] = opts.lot;
  row[PRODUCTION_COL.OUTPUT_ITEM_NAME - 1] = '16 inch Crysta';
  row[PRODUCTION_COL.COMPONENTS_CONSUMED - 1] = JSON.stringify(opts.consumed);
  if (opts.custom !== undefined) row[PRODUCTION_COL.CUSTOM_COMPONENTS - 1] = JSON.stringify(opts.custom);
  row[PRODUCTION_COL.COLOR - 1] = opts.color || '';
  return row;
}

const LOT_A = [
  // stale narration -> must be replaced; name+unit already in sync
  { itemName: 'Carton Box 16 inch', size: '', narration: 'OLD 3-ply note', color: '', sourceType: 'ITEM', qty: 40, colorGroup: 'COMMON', unit: '' },
  // blank narration -> must be filled; name+unit already in sync
  { itemName: 'Frame Sticker---Blue', size: '', narration: '', color: '', sourceType: 'ITEM', qty: 20, colorGroup: 'Blue', unit: '' },
  // Items Master narration is blank -> keep hand-typed note; EXPLICIT stale
  // unit ('Gross', item renamed/reconfigured to base unit 'Kg') -> corrected
  { itemName: 'Adhesive Tape', size: '', narration: 'Hand-typed tape note', color: '', sourceType: 'ITEM', qty: 2, colorGroup: 'COMMON', unit: 'Gross' },
  // not in Items Master at all -> keep everything stored, untouched
  { itemName: 'Rework Charge', size: '', narration: 'Ad-hoc labour', color: '', sourceType: 'ITEM', qty: 1, colorGroup: 'COMMON', unit: 'Nos' },
  // POOL-sourced, and its name IS an Items Master row -> resolved like any other
  { itemName: 'Fitted Frame 16 inch', size: '', narration: '', color: '', sourceType: 'POOL', qty: 40, colorGroup: 'Blue', unit: '' },
  // casing/spelling drift on the SAME identity -> corrected to canonical; its
  // own narration already matches, and unit is blank (must stay blank)
  { itemName: 'curvy sticker backrest', size: '', narration: 'Sticker note', color: '', sourceType: 'ITEM', qty: 5, colorGroup: 'COMMON', unit: '' },
  // fully in sync already (name/narration exact, unit blank) -> untouched
  { itemName: 'Bolt M6', size: '', narration: 'Bolt note', color: '', sourceType: 'ITEM', qty: 12, colorGroup: 'COMMON', unit: '' }
];
// Size is part of the key: the '20 inch' row must NOT pick up the sizeless
// row's narration, and vice versa.
const LOT_B = [
  { itemName: 'Bearing 6203', size: '20 inch', narration: 'stale', color: '', sourceType: 'ITEM', qty: 4, colorGroup: 'COMMON', unit: '' },
  { itemName: 'Bearing 6203', size: '', narration: 'stale', color: '', sourceType: 'ITEM', qty: 8, colorGroup: 'COMMON', unit: '' }
];
const LOT_C_CUSTOM = [
  { itemName: 'Carton Box 16 inch', size: '', narration: 'OLD 3-ply note', colorGroup: 'COMMON', requiredQty: 40 }
];

prod.appendRow(lotRow({ lot: 'LOT-A', qty: 40, consumed: LOT_A, color: 'Blue' }));
prod.appendRow(lotRow({ lot: 'LOT-B', qty: 12, consumed: LOT_B }));
prod.appendRow(lotRow({ lot: 'LOT-C', qty: 40, consumed: LOT_A, custom: LOT_C_CUSTOM, color: 'Blue' }));
// Corrupt JSON must be left exactly as-is, not rewritten or blanked.
const badRow = lotRow({ lot: 'LOT-D', qty: 5, consumed: [] });
badRow[PRODUCTION_COL.COMPONENTS_CONSUMED - 1] = '{not valid json';
prod.appendRow(badRow);

const ROW = { A: 2, B: 3, C: 4, D: 5 };
const readConsumed = row => JSON.parse(prod._get(row, PRODUCTION_COL.COMPONENTS_CONSUMED));
const readCustom = row => JSON.parse(prod._get(row, PRODUCTION_COL.CUSTOM_COMPONENTS));
const byName = (list, name, size) => list.find(c => c.itemName === name && (size === undefined || c.size === size));

// Snapshot every field so we can prove nothing outside name/narration/unit moved.
const beforeFull = JSON.parse(JSON.stringify(prod.rows));

console.log('=== Test 1: the refresh runs and reports what it changed ===');
const res = C.refreshProductionComponentsFromItemsMaster();
assert(res.success, 'returns success: ' + (res.message || ''));
console.log('  message: ' + res.message);
assert(res.data && res.data.lotsScanned === 4, `scanned all 4 lots (got ${res.data && res.data.lotsScanned})`);
// LOT-A consumed: Carton Box narration(1) + Frame Sticker narration(1) +
// Adhesive Tape unit(1) + Rework Charge(0, unknown) + Fitted Frame narration(1)
// + curvy sticker backrest name(1) + Bolt M6(0, already in sync) = 5.
// LOT-B consumed: both Bearing variants' narration = 2.
// LOT-C consumed is another copy of LOT_A = 5. LOT-C custom: Carton Box narration = 1.
assert(res.data && res.data.fieldsUpdated === 13,
  `refreshed exactly 13 fields: 5 (LOT-A) + 2 (LOT-B) + 5 (LOT-C consumed) + 1 (LOT-C custom) ` +
  `(got ${res.data && res.data.fieldsUpdated})`);
assert(res.data && res.data.lotsUpdated === 3,
  `3 of the 4 lots changed — LOT-D's corrupt JSON is skipped (got ${res.data && res.data.lotsUpdated})`);

console.log('\n=== Test 2: stale and blank narrations are refreshed from Items Master ===');
const a = readConsumed(ROW.A);
assert(byName(a, 'Carton Box 16 inch').narration === 'Corrugated 5-ply (revised)',
  `stale note replaced (got "${byName(a, 'Carton Box 16 inch').narration}")`);
assert(byName(a, 'Frame Sticker---Blue').narration === 'Sticker kit A',
  `blank note filled in (got "${byName(a, 'Frame Sticker---Blue').narration}")`);

console.log('\n=== Test 3: nothing is blanked/touched where Items Master has nothing to say ===');
const adhesive = byName(a, 'Adhesive Tape');
assert(adhesive.narration === 'Hand-typed tape note',
  `blank-in-Items-Master keeps its stored note (got "${adhesive.narration}")`);
const rework = byName(a, 'Rework Charge');
assert(rework.narration === 'Ad-hoc labour' && rework.unit === 'Nos' && rework.itemName === 'Rework Charge',
  `unknown item is fully untouched (got narration="${rework.narration}", unit="${rework.unit}", itemName="${rework.itemName}")`);

console.log('\n=== Test 4: explicit stale unit is corrected; blank unit stays blank ===');
assert(adhesive.unit === 'Kg', `Adhesive Tape's explicit stale unit 'Gross' corrected to current Base Unit 'Kg' (got "${adhesive.unit}")`);
assert(byName(a, 'Carton Box 16 inch').unit === '', 'Carton Box\'s blank unit is left blank, not stamped to "Pcs"');
const bolt = byName(a, 'Bolt M6');
assert(bolt.unit === '' && bolt.narration === 'Bolt note' && bolt.itemName === 'Bolt M6',
  `already-in-sync component (blank unit) is completely untouched (got unit="${bolt.unit}", narration="${bolt.narration}", itemName="${bolt.itemName}")`);

console.log('\n=== Test 5: item name is corrected to current Items Master casing, same identity only ===');
const sticker = a.find(c => c.itemName.toLowerCase() === 'curvy sticker backrest');
assert(sticker && sticker.itemName === 'CURVY STICKER Backrest',
  `casing-drifted name corrected to canonical (got "${sticker && sticker.itemName}")`);
assert(sticker && sticker.narration === 'Sticker note' && sticker.unit === '',
  'narration/unit for that same component are untouched (already in sync)');

console.log('\n=== Test 6: POOL rows are resolved by name (metadata, not identity) ===');
assert(byName(a, 'Fitted Frame 16 inch').narration === 'WIP frame from Fitting',
  `POOL row picked up its Items Master narration (got "${byName(a, 'Fitted Frame 16 inch').narration}")`);
assert(byName(a, 'Fitted Frame 16 inch').sourceType === 'POOL',
  'POOL row keeps sourceType POOL (identity untouched)');

console.log('\n=== Test 7: size is part of the lookup key ===');
const b = readConsumed(ROW.B);
assert(byName(b, 'Bearing 6203', '20 inch').narration === 'Sealed, size-specific note',
  `'20 inch' variant got its own narration (got "${byName(b, 'Bearing 6203', '20 inch').narration}")`);
assert(byName(b, 'Bearing 6203', '').narration === 'Generic, no size',
  `sizeless variant got its own narration (got "${byName(b, 'Bearing 6203', '').narration}")`);

console.log('\n=== Test 8: the CUSTOM_COMPONENTS sheet snapshot is repaired too ===');
const cCustom = readCustom(ROW.C);
assert(cCustom[0].narration === 'Corrugated 5-ply (revised)',
  `customComponents narration refreshed (got "${cCustom[0].narration}")`);
assert(cCustom[0].requiredQty === 40, `customComponents requiredQty intact (got ${cCustom[0].requiredQty})`);

console.log('\n=== Test 9: unparseable JSON is skipped, not corrupted ===');
assert(prod._get(ROW.D, PRODUCTION_COL.COMPONENTS_CONSUMED) === '{not valid json',
  'corrupt cell left byte-for-byte unchanged');

console.log('\n=== Test 10: ONLY itemName/narration/unit changed — every other field round-trips ===');
const afterFull = JSON.parse(JSON.stringify(prod.rows));
let strayDiffs = [];
afterFull.forEach((row, ri) => {
  (row || []).forEach((val, ci) => {
    const before = (beforeFull[ri] || [])[ci];
    if (before === val) return;
    const col = ci + 1;
    if (col === PRODUCTION_COL.COMPONENTS_CONSUMED || col === PRODUCTION_COL.CUSTOM_COMPONENTS) {
      let pb, pa;
      try { pb = JSON.parse(before); pa = JSON.parse(val); } catch (e) { strayDiffs.push(`row ${ri + 1} col ${col} unparseable`); return; }
      if (pb.length !== pa.length) { strayDiffs.push(`row ${ri + 1} col ${col} length changed`); return; }
      pb.forEach((cb, i) => {
        const ca = pa[i];
        Object.keys(cb).forEach(k => {
          if (k === 'narration' || k === 'unit' || k === 'itemName') return;
          if (JSON.stringify(cb[k]) !== JSON.stringify(ca[k])) {
            strayDiffs.push(`row ${ri + 1} col ${col} comp ${i} field "${k}": ${JSON.stringify(cb[k])} -> ${JSON.stringify(ca[k])}`);
          }
        });
        Object.keys(ca).forEach(k => {
          if (!(k in cb)) strayDiffs.push(`row ${ri + 1} col ${col} comp ${i} gained field "${k}"`);
        });
      });
      return;
    }
    strayDiffs.push(`row ${ri + 1} col ${col}: ${JSON.stringify(before)} -> ${JSON.stringify(val)}`);
  });
});
assert(strayDiffs.length === 0, 'no field other than itemName/narration/unit changed anywhere' +
  (strayDiffs.length ? ' -- stray diffs: ' + strayDiffs.join('; ') : ''));

console.log('\n=== Test 11: idempotent — an immediate re-run changes nothing ===');
const again = C.refreshProductionComponentsFromItemsMaster();
assert(again.success, 're-run succeeds');
assert(again.data && again.data.fieldsUpdated === 0,
  `re-run reports 0 changes (got ${again.data && again.data.fieldsUpdated})`);
assert(again.data && again.data.lotsUpdated === 0,
  `re-run reports 0 lots touched (got ${again.data && again.data.lotsUpdated})`);
assert(/already match/i.test(again.message || ''), `no-op message says so: "${again.message}"`);

console.log('\n=== Test 12: a re-run after an Items Master edit picks the new value up ===');
// Row 2 of the Items sheet is 'Carton Box 16 inch' (row 1 is the header).
items._set(2, ITEMS_COL.NARRATION, 'Corrugated 7-ply (2nd revision)');
// Carton Box is stored 3 times: LOT-A consumed, LOT-C consumed, LOT-C custom.
const third = C.refreshProductionComponentsFromItemsMaster();
assert(third.success && third.data.fieldsUpdated === 3,
  `re-run refreshed all 3 stored copies of the edited item (got ${third.data && third.data.fieldsUpdated})`);
assert(byName(readConsumed(ROW.A), 'Carton Box 16 inch').narration === 'Corrugated 7-ply (2nd revision)',
  'LOT-A consumed picked up the 2nd revision');
assert(readCustom(ROW.C)[0].narration === 'Corrugated 7-ply (2nd revision)',
  'LOT-C custom picked up the 2nd revision');

console.log('\n=== Test 13: saveProductionSheet writes narration fresh, not the client value ===');
// A stale client payload must not be able to re-persist the old note and undo
// the refresh (module_production.js#_withMasterNarration).
const saveRes = C.saveProductionSheet(ROW.C, '', 40, JSON.stringify([
  { itemName: 'Carton Box 16 inch', size: '', narration: 'STALE FROM CLIENT', colorGroup: 'COMMON', requiredQty: 40 },
  { itemName: 'Rework Charge', size: '', narration: 'client-only note', colorGroup: 'COMMON', requiredQty: 1 }
]), 'sheet remark');
assert(saveRes.success, 'saveProductionSheet succeeded: ' + (saveRes.message || ''));
const saved = readCustom(ROW.C);
assert(byName(saved, 'Carton Box 16 inch').narration === 'Corrugated 7-ply (2nd revision)',
  `client's stale narration overridden by Items Master (got "${byName(saved, 'Carton Box 16 inch').narration}")`);
assert(byName(saved, 'Rework Charge').narration === 'client-only note',
  `an unknown item still keeps the client's own note (got "${byName(saved, 'Rework Charge').narration}")`);

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
