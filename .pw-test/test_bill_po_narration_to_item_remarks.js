/**
 * Verifies that the narration typed on a PO / Receive-New-Bill line is
 * redirected into the item's Items Master REMARKS column
 * (autoExtractFromPoOrBill, module_vendors.js) — the shared hook both
 * saveBill and savePO call after writing their own rows.
 *
 * Before: a line's narration only ever populated Items Master's NARRATION,
 * and only when that field happened to be blank — so a corrected or
 * vendor-specific description entered on a bill was silently dropped for
 * every item that already had a narration.
 *
 * Now: narration overwrites REMARKS, so Remarks always reflects the most
 * recent purchase-side description. NARRATION is deliberately left alone for
 * items that already exist — that is the hand-maintained note Production
 * prints and resolves against (_getItemNarrationMap / _resolveDisplayNarration),
 * and letting a vendor's wording overwrite it would change printed sheets.
 *
 * Covered:
 *   - narration OVERWRITES an existing, different Remarks
 *   - narration fills a blank Remarks
 *   - an item's existing NARRATION is never touched by a bill/PO line
 *   - a blank line narration clears nothing (existing Remarks survives)
 *   - re-saving the same narration is a no-op (no redundant write)
 *   - two lines for the SAME item in one save: the last one wins, and the
 *     row is written once
 *   - a brand-new item (not yet in Items Master) still seeds NARRATION, per
 *     the documented scope decision, and is created rather than skipped
 *   - the narration still lands on the ledger row itself — it is part of
 *     _buildPoLineKey, so PO<->Bill quantity matching depends on it
 *
 * Run: node .pw-test/test_bill_po_narration_to_item_remarks.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

class FakeRange {
  constructor(s, r, c, nr, nc) { this.sheet = s; this.row = r; this.col = c; this.numRows = nr; this.numCols = nc; }
  getValues() { const o = []; for (let r = 0; r < this.numRows; r++) { const a = []; for (let c = 0; c < this.numCols; c++) a.push(this.sheet._get(this.row + r, this.col + c)); o.push(a); } return o; }
  getValue() { return this.sheet._get(this.row, this.col); }
  setValues(v) { v.forEach((ra, r) => ra.forEach((val, c) => this.sheet._set(this.row + r, this.col + c, val))); return this; }
  setValue(v) { this.sheet._set(this.row, this.col, v); this.sheet._writes.push({ row: this.row, col: this.col, value: v }); return this; }
  clearContent() { for (let r = 0; r < this.numRows; r++) for (let c = 0; c < this.numCols; c++) this.sheet._set(this.row + r, this.col + c, ''); return this; }
  setFontWeight() { return this; } setBackground() { return this; } setNumberFormat() { return this; }
  setHorizontalAlignment() { return this; } setFontColor() { return this; } setWrap() { return this; }
  setFontSize() { return this; } setBorder() { return this; }
}
class FakeSheet {
  constructor(n) { this.name = n; this.rows = []; this._writes = []; }
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
  deleteRow(r) { this.rows.splice(r - 1, 1); }
  insertColumnBefore() { return this; } insertColumnsAfter() { return this; }
  setFrozenRows() { return this; } setColumnWidth() { return this; } autoResizeColumn() { return this; }
  getFilter() { return null; } clear() { this.rows = []; return this; }
}
class FakeSpreadsheet {
  constructor() { this.sheets = {}; }
  getSheetByName(n) { return this.sheets[n] || null; }
  addSheet(n) { const s = new FakeSheet(n); this.sheets[n] = s; return s; }
  insertSheet(n) { return this.addSheet(n); }
  getSheets() { return Object.values(this.sheets); }
}

const ss = new FakeSpreadsheet();
const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush: () => {} },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, waitLock: () => true, releaseLock: () => {} }) },
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
  console, Logger: { log: () => {} },
  Utilities: { getUuid: () => 'u' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);
['config.js', 'utils.js', 'module_units.js', 'module_items.js', 'module_vendors.js',
 'module_stock.js', 'module_bill.js', 'module_po.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});
vm.runInContext('global.APP_CONFIG=APP_CONFIG; global.ITEMS_COL=ITEMS_COL;', ctx, { filename: 'expose.js' });
const C = ctx;
const { APP_CONFIG, ITEMS_COL } = C;

let failures = 0;
function assert(cond, msg) { if (!cond) { failures++; console.error('  FAIL:', msg); } else { console.log('  PASS:', msg); } }

// ── Seed sheets ──────────────────────────────────────────────────────────
const items = ss.addSheet(APP_CONFIG.SHEETS.ITEMS);
items.appendRow(['Item Name', 'Size', 'Remarks', 'Narration', 'Specification',
                 'Base Unit', 'Purchase Unit', 'Weight/Base Unit', 'Vendor 1', 'Rate 1']);
// Remarks already has hand-typed text; Narration is the note Production uses.
items.appendRow(['Carton Box', '', 'OLD remark', 'Production narration - keep me', '', 'Pcs', 'Pcs', 0, '', '']);
// Blank Remarks, to prove a fill also works.
items.appendRow(['Bubble Wrap', '', '', 'Wrap narration', '', 'Kg', 'Kg', 0, '', '']);
// Used for the "same item twice in one save" case.
items.appendRow(['Steel Rod', '12mm', 'rod remark v0', 'Rod narration', '', 'Pcs', 'Pcs', 0, '', '']);

ss.addSheet(APP_CONFIG.SHEETS.VENDORS).appendRow(['Vendor Name', 'Contact', 'Address', 'GST', 'Remarks']);
ss.addSheet(APP_CONFIG.SHEETS.STOCK).appendRow(['Item Name', 'Size', 'Current Stock']);
ss.addSheet(APP_CONFIG.SHEETS.RATE_HISTORY);

const ROW = { CARTON: 2, BUBBLE: 3, ROD: 4 };
const remarksAt = r => String(items._get(r, ITEMS_COL.REMARKS) || '');
const narrationAt = r => String(items._get(r, ITEMS_COL.NARRATION) || '');
const findRow = (name, size) => {
  for (let r = 2; r <= items.getLastRow(); r++) {
    if (String(items._get(r, ITEMS_COL.ITEM_NAME)).toLowerCase() === name.toLowerCase() &&
        String(items._get(r, ITEMS_COL.SIZE) || '').toLowerCase() === String(size || '').toLowerCase()) return r;
  }
  return -1;
};

console.log('=== Test 1: line narration OVERWRITES an existing, different Remarks ===');
C.autoExtractFromPoOrBill('Acme Supplies', '9990001111', [
  { name: 'Carton Box', size: '', narration: '5-ply corrugated (vendor spec)', price: 12, unit: 'Pcs' }
], { date: new Date(2026, 6, 30), poNumber: 'PO-1', billNumber: 'B-1' });

assert(remarksAt(ROW.CARTON) === '5-ply corrugated (vendor spec)',
  `Remarks overwritten with the bill's narration (got "${remarksAt(ROW.CARTON)}")`);

console.log('\n=== Test 2: the item\'s own NARRATION is left untouched ===');
assert(narrationAt(ROW.CARTON) === 'Production narration - keep me',
  `Items Master Narration is NOT overwritten by the bill line (got "${narrationAt(ROW.CARTON)}")`);

console.log('\n=== Test 3: a blank Remarks is filled ===');
C.autoExtractFromPoOrBill('Acme Supplies', '', [
  { name: 'Bubble Wrap', size: '', narration: 'Roll, 1m width', price: 40, unit: 'Kg' }
], { date: new Date(2026, 6, 30), billNumber: 'B-2' });
assert(remarksAt(ROW.BUBBLE) === 'Roll, 1m width',
  `blank Remarks filled from the line narration (got "${remarksAt(ROW.BUBBLE)}")`);
assert(narrationAt(ROW.BUBBLE) === 'Wrap narration',
  `that item's Narration still untouched (got "${narrationAt(ROW.BUBBLE)}")`);

console.log('\n=== Test 4: a BLANK line narration clears nothing ===');
C.autoExtractFromPoOrBill('Acme Supplies', '', [
  { name: 'Carton Box', size: '', narration: '', price: 13, unit: 'Pcs' }
], { date: new Date(2026, 6, 30), billNumber: 'B-3' });
assert(remarksAt(ROW.CARTON) === '5-ply corrugated (vendor spec)',
  `existing Remarks survives a line with no narration (got "${remarksAt(ROW.CARTON)}")`);

console.log('\n=== Test 5: re-saving the SAME narration performs no redundant write ===');
items._writes.length = 0;
C.autoExtractFromPoOrBill('Acme Supplies', '', [
  { name: 'Carton Box', size: '', narration: '5-ply corrugated (vendor spec)', price: 12, unit: 'Pcs' }
], { date: new Date(2026, 6, 30), billNumber: 'B-4' });
const remarkWrites = items._writes.filter(w => w.col === ITEMS_COL.REMARKS);
assert(remarkWrites.length === 0,
  `no Remarks write queued when the value already matches (got ${remarkWrites.length})`);
assert(remarksAt(ROW.CARTON) === '5-ply corrugated (vendor spec)', 'value unchanged, as expected');

console.log('\n=== Test 6: two lines for the SAME item in one save — last wins, one write ===');
items._writes.length = 0;
C.autoExtractFromPoOrBill('Acme Supplies', '', [
  { name: 'Steel Rod', size: '12mm', narration: 'first description', price: 55, unit: 'Pcs' },
  { name: 'Steel Rod', size: '12mm', narration: 'SECOND description', price: 55, unit: 'Pcs' }
], { date: new Date(2026, 6, 30), billNumber: 'B-5' });
const rodWrites = items._writes.filter(w => w.col === ITEMS_COL.REMARKS && w.row === ROW.ROD);
assert(remarksAt(ROW.ROD) === 'SECOND description',
  `the last line's narration wins (got "${remarksAt(ROW.ROD)}")`);
assert(rodWrites.length === 2,
  `each distinct value is written once, not re-queued per line (got ${rodWrites.length} writes)`);

console.log('\n=== Test 7: a brand-new item still seeds NARRATION (documented scope) ===');
const before = items.getLastRow();
C.autoExtractFromPoOrBill('Acme Supplies', '', [
  { name: 'Nylon Strap', size: '25mm', narration: 'Woven, black', price: 8, unit: 'Pcs' }
], { date: new Date(2026, 6, 30), billNumber: 'B-6' });
const newRow = findRow('Nylon Strap', '25mm');
assert(newRow !== -1 && items.getLastRow() === before + 1,
  `the unknown item was created (row ${newRow}, lastRow ${before} -> ${items.getLastRow()})`);
assert(narrationAt(newRow) === 'Woven, black',
  `a brand-new item seeds NARRATION from the line (got "${narrationAt(newRow)}")`);
assert(remarksAt(newRow) === '',
  `and its Remarks starts blank — the redirect is scoped to already-existing items (got "${remarksAt(newRow)}")`);

console.log('\n=== Test 8: narration still reaches the ledger row (PO<->Bill matching key) ===');
// _buildPoLineKey includes narration, so PO/Bill quantity matching breaks if
// the value stops being persisted on the row itself.
const key = C._buildPoLineKey('PO-1', 'Carton Box', '', '5-ply corrugated (vendor spec)');
assert(key.indexOf('5-ply corrugated (vendor spec)'.toLowerCase()) !== -1,
  `narration is part of the PO line key, so it must stay on the row (key: "${key}")`);

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
