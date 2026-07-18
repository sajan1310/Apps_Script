/**
 * Standalone Node harness covering a deep-dive fix from the 2026-07-18
 * calculations audit: backfillBOMItemRefs (module_bom.js) rewrites every
 * BOM row referencing a renamed/merged item's old (name, size) to its new
 * identity, but previously did so blindly — if the SAME product already
 * had a separate row for the rename's TARGET identity, the rename created
 * a second row sharing an identical (Product, Item, Size, Color) key.
 * getBOMData sums every component row's cost as a plain array (no
 * de-duplication), so the resulting duplicate pair silently double-counted
 * that item's cost in the product's Material Cost total.
 *
 * Fix: after renaming, detect any (Product, Item, Size, Color) collision
 * this created, sum QTY_PER_PRODUCT into the first row, and delete the
 * rest — same additive-quantity precedent as the Items Master merge.
 *
 * Run: node .pw-test/test_bom_backfill_duplicate_merge.js
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
}

class FakeSpreadsheet {
  constructor() { this.sheets = {}; }
  getSheetByName(name) { return this.sheets[name] || null; }
  addSheet(name) { const s = new FakeSheet(name); this.sheets[name] = s; return s; }
  insertSheet(name) { return this.addSheet(name); }
}

const ss = new FakeSpreadsheet();

const fakeCache = {
  _store: {},
  get(k) { return Object.prototype.hasOwnProperty.call(this._store, k) ? this._store[k] : null; },
  put(k, v) { this._store[k] = v; },
  remove(k) { delete this._store[k]; }
};

const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush: () => {} },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  CacheService: { getScriptCache: () => fakeCache },
  console,
  Logger: { log: () => {} },
  Utilities: { getUuid: () => 'uuid-' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);

['config.js', 'utils.js', 'module_process.js', 'module_production.js', 'module_bom.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});

vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.BOM_COL = BOM_COL;
  global.BOM_AUTH_CACHE_PREFIX = BOM_AUTH_CACHE_PREFIX;
`, ctx, { filename: 'expose.js' });

const { APP_CONFIG, BOM_COL, BOM_AUTH_CACHE_PREFIX, backfillBOMItemRefs, getBOMData } = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); }
}

const TOKEN = 'test-token';
fakeCache.put(BOM_AUTH_CACHE_PREFIX + TOKEN, '1');

console.log('\n=== Fix: renaming an item into a name that already exists in the SAME product\'s BOM merges instead of duplicating ===');
{
  const bomSheet = ss.addSheet(APP_CONFIG.SHEETS.BOM);
  // Row 2: Product PRD-1 already has a row for "Bolt Std" (the rename's target identity).
  bomSheet._set(2, BOM_COL.PRODUCT_ID, 'PRD-1');
  bomSheet._set(2, BOM_COL.PRODUCT_NAME, 'Test Bike');
  bomSheet._set(2, BOM_COL.ITEM_NAME, 'Bolt Std');
  bomSheet._set(2, BOM_COL.SIZE, '');
  bomSheet._set(2, BOM_COL.RATE, 2);
  bomSheet._set(2, BOM_COL.QTY_PER_PRODUCT, 3);
  bomSheet._set(2, BOM_COL.COLOR, '');

  // Row 3: same product, a DIFFERENT row for "Bolt Old" -- about to be
  // renamed to "Bolt Std", the same identity as row 2.
  bomSheet._set(3, BOM_COL.PRODUCT_ID, 'PRD-1');
  bomSheet._set(3, BOM_COL.PRODUCT_NAME, 'Test Bike');
  bomSheet._set(3, BOM_COL.ITEM_NAME, 'Bolt Old');
  bomSheet._set(3, BOM_COL.SIZE, '');
  bomSheet._set(3, BOM_COL.RATE, 2);
  bomSheet._set(3, BOM_COL.QTY_PER_PRODUCT, 5);
  bomSheet._set(3, BOM_COL.COLOR, '');

  // Row 4: a DIFFERENT product also referencing "Bolt Old" -- must be
  // renamed too, but has no collision (PRD-2 has no other "Bolt Std" row).
  bomSheet._set(4, BOM_COL.PRODUCT_ID, 'PRD-2');
  bomSheet._set(4, BOM_COL.PRODUCT_NAME, 'Other Bike');
  bomSheet._set(4, BOM_COL.ITEM_NAME, 'Bolt Old');
  bomSheet._set(4, BOM_COL.SIZE, '');
  bomSheet._set(4, BOM_COL.RATE, 2);
  bomSheet._set(4, BOM_COL.QTY_PER_PRODUCT, 7);
  bomSheet._set(4, BOM_COL.COLOR, '');

  backfillBOMItemRefs('Bolt Old', '', 'Bolt Std', '');

  const res = getBOMData(TOKEN);
  assert(res.success, 'getBOMData succeeds after backfill: ' + res.message);

  const prd1 = (res.data || []).find(p => p.productId === 'PRD-1');
  assert(!!prd1, 'PRD-1 found');
  assert(prd1 && prd1.components.length === 1, `PRD-1 has exactly ONE "Bolt Std" row after merge, not two (got ${prd1 && prd1.components.length})`);
  assert(prd1 && prd1.components[0].qtyPerProduct === 8, `merged qty is 3 + 5 = 8 (got ${prd1 && prd1.components[0].qtyPerProduct})`);
  assert(prd1 && prd1.totalCost === 16, `PRD-1 totalCost reflects the merged qty (2 x 8 = 16), not double-counted (got ${prd1 && prd1.totalCost})`);

  const prd2 = (res.data || []).find(p => p.productId === 'PRD-2');
  assert(!!prd2, 'PRD-2 found');
  assert(prd2 && prd2.components.length === 1, `PRD-2's single row just got renamed, no collision there (got ${prd2 && prd2.components.length})`);
  assert(prd2 && prd2.components[0].itemName === 'Bolt Std', `PRD-2's row renamed to "Bolt Std" (got ${prd2 && prd2.components[0].itemName})`);
  assert(prd2 && prd2.components[0].qtyPerProduct === 7, `PRD-2's qty untouched by the unrelated PRD-1 merge (got ${prd2 && prd2.components[0].qtyPerProduct})`);
}

console.log('\n=== Regression: a rename with no collision anywhere still behaves exactly as before ===');
{
  const bomSheet = ss.addSheet(APP_CONFIG.SHEETS.BOM);
  bomSheet._set(2, BOM_COL.PRODUCT_ID, 'PRD-3');
  bomSheet._set(2, BOM_COL.PRODUCT_NAME, 'Solo Bike');
  bomSheet._set(2, BOM_COL.ITEM_NAME, 'Chain Old');
  bomSheet._set(2, BOM_COL.SIZE, '');
  bomSheet._set(2, BOM_COL.RATE, 10);
  bomSheet._set(2, BOM_COL.QTY_PER_PRODUCT, 1);
  bomSheet._set(2, BOM_COL.COLOR, '');

  backfillBOMItemRefs('Chain Old', '', 'Chain New', '');

  const res = getBOMData(TOKEN);
  const prd3 = (res.data || []).find(p => p.productId === 'PRD-3');
  assert(!!prd3 && prd3.components.length === 1 && prd3.components[0].itemName === 'Chain New',
    `plain rename with no collision still just renames the one row (got ${prd3 && JSON.stringify(prd3.components)})`);
}

console.log('\n=== Regression: same item+size but DIFFERENT color is not treated as a collision ===');
{
  const bomSheet = ss.getSheetByName(APP_CONFIG.SHEETS.BOM);
  const row1 = bomSheet.getLastRow() + 1;
  bomSheet._set(row1, BOM_COL.PRODUCT_ID, 'PRD-4');
  bomSheet._set(row1, BOM_COL.PRODUCT_NAME, 'Color Bike');
  bomSheet._set(row1, BOM_COL.ITEM_NAME, 'Paint Std');
  bomSheet._set(row1, BOM_COL.SIZE, '');
  bomSheet._set(row1, BOM_COL.RATE, 5);
  bomSheet._set(row1, BOM_COL.QTY_PER_PRODUCT, 1);
  bomSheet._set(row1, BOM_COL.COLOR, 'Red');

  const row2 = bomSheet.getLastRow() + 1;
  bomSheet._set(row2, BOM_COL.PRODUCT_ID, 'PRD-4');
  bomSheet._set(row2, BOM_COL.PRODUCT_NAME, 'Color Bike');
  bomSheet._set(row2, BOM_COL.ITEM_NAME, 'Paint Old');
  bomSheet._set(row2, BOM_COL.SIZE, '');
  bomSheet._set(row2, BOM_COL.RATE, 5);
  bomSheet._set(row2, BOM_COL.QTY_PER_PRODUCT, 1);
  bomSheet._set(row2, BOM_COL.COLOR, 'Blue');

  backfillBOMItemRefs('Paint Old', '', 'Paint Std', '');

  const res = getBOMData(TOKEN);
  const prd4 = (res.data || []).find(p => p.productId === 'PRD-4');
  assert(!!prd4 && prd4.components.length === 2, `different-color rows stay separate, not merged (got ${prd4 && prd4.components.length})`);
}

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
