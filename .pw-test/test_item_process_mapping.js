/**
 * Standalone Node harness that mocks the Apps Script runtime well enough to
 * load and execute the REAL server-side files (config.js, utils.js,
 * module_process.js) and exercise getProcessesForItem — the item-side
 * (inverse) view of the Process Components sheet that powers Item Master's
 * "Used in Processes" section.
 *
 * The cases that matter are the two identity rules in getProcessesForItem's
 * docblock: SOURCE_TYPE 'POOL' rows belong to a different identity space and
 * must never surface as an item's recipe even on an exact name collision,
 * and COMMON rows (the process-wide recipe entry) must stay separated from
 * per-color override rows.
 *
 * Run: node .pw-test/test_item_process_mapping.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

// ─────────────────────────────────────────────────────────────────────────
// Minimal in-memory Sheet/Range/Spreadsheet mocks (mirrors test_process_color_groups.js)
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

// Deliberately NO CacheService in the sandbox: getCachedListResponse treats
// an unavailable CacheService as "skip caching" (see utils.js), which is
// what we want so each getProcessData call re-reads the sheet we just wrote.
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

['config.js', 'utils.js', 'module_process.js'].forEach(f => {
  const code = fs.readFileSync(path.join(ROOT, f), 'utf8');
  vm.runInContext(code, ctx, { filename: f });
});

vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.PROCESS_COMPONENTS_COL = PROCESS_COMPONENTS_COL;
  global.COMPONENT_COLOR_GROUP_COMMON = COMPONENT_COLOR_GROUP_COMMON;
  global.PRODUCTION_COL = PRODUCTION_COL;
  global.ITEMS_COL = ITEMS_COL;
`, ctx, { filename: 'expose.js' });

const {
  APP_CONFIG, PROCESS_COMPONENTS_COL, ITEMS_COL, PRODUCTION_COL,
  saveProcess, getProcessesForItem, saveItemProcessMappings
} = ctx;

// _itemExistsInMaster reads the real Items Master sheet, so the items under
// test have to exist there. Seeded directly rather than via module_items.js
// (saveItem) to keep this harness to the one module it is testing.
{
  const items = ss.addSheet(APP_CONFIG.SHEETS.ITEMS);
  items.appendRow(['Item Name', 'Size']);
  [['Thinner', '1L'], ['Thinner', '5L'], ['Brush', ''], ['Primer', '2L']]
    .forEach(([n, s]) => items.appendRow([n, s]));
}

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
function byId(records, processId) {
  return (records || []).find(r => r.processId === processId) || null;
}

// ─────────────────────────────────────────────────────────────────────────
// Fixture: 4 processes exercising every shape getProcessesForItem must
// distinguish for the item "Thinner" / size "1L".
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Fixture: build 4 processes ===');
const ids = {};
{
  // P1 — plain COMMON row, blank unit (= item's own Base Unit).
  let res = saveProcess({
    processName: 'Frame Painting', sequence: 1, lotPrefix: 'ZFP',
    outputItemName: 'Painted Frame TestOut', isFinalStage: false, active: true, remarks: '',
    components: JSON.stringify([
      { itemName: 'Thinner', size: '1L', sourceType: 'ITEM', qtyPerUnit: 0.25, colorGroup: 'COMMON', remarks: 'thin the enamel' },
      { itemName: 'Brush', size: '', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'COMMON' }
    ])
  });
  assert(res.success, 'P1 saved: ' + res.message);
  ids.p1 = res.data && res.data.processId;

  // P2 — COMMON row carrying an explicit Unit, PLUS two color rows.
  res = saveProcess({
    processName: 'Fork Painting', sequence: 2, lotPrefix: 'ZFK',
    outputItemName: 'Painted Fork TestOut', isFinalStage: false, active: true, remarks: '',
    components: JSON.stringify([
      { itemName: 'Thinner', size: '1L', sourceType: 'ITEM', qtyPerUnit: 0.4, unit: 'Kg', colorGroup: 'COMMON' },
      { itemName: 'Thinner', size: '1L', sourceType: 'ITEM', qtyPerUnit: 0.6, colorGroup: 'Red', colorAxis: 'Fork Color' },
      { itemName: 'Thinner', size: '1L', sourceType: 'ITEM', qtyPerUnit: 0.5, colorGroup: 'Blue', colorAxis: 'Fork Color' }
    ])
  });
  assert(res.success, 'P2 saved: ' + res.message);
  ids.p2 = res.data && res.data.processId;

  // P3 — the trap. A POOL row whose itemName collides EXACTLY with the
  // Items Master item under test, plus a same-name/different-size ITEM row.
  // Neither may be reported as Thinner 1L's recipe.
  res = saveProcess({
    processName: 'Frame Fitting', sequence: 3, lotPrefix: 'ZFF',
    outputItemName: 'Fitted Frame TestOut', isFinalStage: false, active: true, remarks: '',
    components: JSON.stringify([
      { itemName: 'Thinner', size: '1L', sourceType: 'POOL', qtyPerUnit: 9.99, colorGroup: 'COMMON' },
      { itemName: 'Thinner', size: '5L', sourceType: 'ITEM', qtyPerUnit: 7.77, colorGroup: 'COMMON' }
    ])
  });
  assert(res.success, 'P3 saved: ' + res.message);
  ids.p3 = res.data && res.data.processId;

  // P4 — INACTIVE process holding a color-only recipe entry (no COMMON row).
  res = saveProcess({
    processName: 'Legacy Dipping', sequence: 4, lotPrefix: 'ZLD',
    outputItemName: 'Dipped Frame TestOut', isFinalStage: false, active: false, remarks: '',
    components: JSON.stringify([
      { itemName: 'Thinner', size: '1L', sourceType: 'ITEM', qtyPerUnit: 1.2, colorGroup: 'Green' }
    ])
  });
  assert(res.success, 'P4 saved: ' + res.message);
  ids.p4 = res.data && res.data.processId;
}

// ─────────────────────────────────────────────────────────────────────────
// Test 1: the happy path — COMMON rows reported with qty/unit/remarks
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 1: COMMON rows reported as inRecipe ===');
let all;
{
  const res = getProcessesForItem('Thinner', '1L');
  assert(res.success, 'getProcessesForItem succeeds: ' + res.message);
  all = res.data || [];
  assert(all.length >= 4, 'every process is returned, not just matching ones (got ' + all.length + ')');

  const p1 = byId(all, ids.p1);
  assert(p1 && p1.inRecipe === true, 'P1 reports inRecipe=true');
  assert(p1 && p1.qtyPerUnit === 0.25, 'P1 qtyPerUnit is 0.25 (got ' + (p1 && p1.qtyPerUnit) + ')');
  assert(p1 && p1.unit === '', 'P1 blank unit stays blank = item Base Unit (got "' + (p1 && p1.unit) + '")');
  assert(p1 && p1.remarks === 'thin the enamel', 'P1 remarks carried through (got "' + (p1 && p1.remarks) + '")');
  assert(p1 && p1.colorVariants.length === 0, 'P1 has no color variants');

  const p2 = byId(all, ids.p2);
  assert(p2 && p2.inRecipe === true, 'P2 reports inRecipe=true');
  assert(p2 && p2.qtyPerUnit === 0.4, 'P2 qtyPerUnit is 0.4 (got ' + (p2 && p2.qtyPerUnit) + ')');
  assert(p2 && p2.unit === 'Kg', 'P2 explicit unit Kg preserved (got "' + (p2 && p2.unit) + '")');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 2: color rows are reported separately, never as the COMMON qty
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 2: color rows kept out of the COMMON entry ===');
{
  const p2 = byId(all, ids.p2);
  assert(p2 && p2.colorVariants.length === 2, 'P2 surfaces both color rows (got ' + (p2 && p2.colorVariants.length) + ')');
  const names = (p2.colorVariants || []).map(v => v.colorGroup);
  assert(JSON.stringify(names) === JSON.stringify(['Blue', 'Red']), 'color variants sorted by name (got ' + JSON.stringify(names) + ')');
  const blue = p2.colorVariants.find(v => v.colorGroup === 'Blue');
  assert(blue && blue.qtyPerUnit === 0.5, 'Blue variant keeps its own qty 0.5 (got ' + (blue && blue.qtyPerUnit) + ')');
  assert(blue && blue.colorAxis === 'Fork Color', 'Blue variant keeps its color axis (got "' + (blue && blue.colorAxis) + '")');
  // The whole point: the COMMON qty must not have been overwritten by a
  // color row, which is exactly what a naive "last row for this item wins"
  // implementation would do.
  assert(p2.qtyPerUnit === 0.4, 'P2 COMMON qty unaffected by color rows (got ' + p2.qtyPerUnit + ')');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 3: POOL rows and other sizes never surface (identity rules)
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 3: POOL rows and size mismatches excluded ===');
{
  const p3 = byId(all, ids.p3);
  assert(p3 && p3.inRecipe === false, 'P3 reports inRecipe=false despite an exact-name POOL row');
  assert(p3 && p3.qtyPerUnit === null, 'P3 qty is null, not the POOL row 9.99 or the 5L row 7.77 (got ' + (p3 && p3.qtyPerUnit) + ')');
  assert(p3 && p3.colorVariants.length === 0, 'P3 has no color variants');

  // ...and the 5L row does belong to the 5L item.
  const res5 = getProcessesForItem('Thinner', '5L');
  assert(res5.success, 'getProcessesForItem for the 5L size succeeds');
  const p3For5L = byId(res5.data, ids.p3);
  assert(p3For5L && p3For5L.inRecipe === true, 'the 5L ITEM row IS reported for size 5L');
  assert(p3For5L && p3For5L.qtyPerUnit === 7.77, '5L qty is 7.77 (got ' + (p3For5L && p3For5L.qtyPerUnit) + ')');
  const p1For5L = byId(res5.data, ids.p1);
  assert(p1For5L && p1For5L.inRecipe === false, 'the 1L row is NOT reported for size 5L');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 4: color-only process — inRecipe false but variants present
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 4: color-only recipe entry, on an inactive process ===');
{
  const p4 = byId(all, ids.p4);
  assert(p4 && p4.inRecipe === false, 'P4 has no COMMON row so inRecipe=false');
  assert(p4 && p4.colorVariants.length === 1, 'P4 still surfaces its color-only row (got ' + (p4 && p4.colorVariants.length) + ')');
  assert(p4 && p4.colorVariants[0].colorGroup === 'Green', 'P4 color row is Green');
  assert(p4 && p4.active === false, 'P4 active flag reported so the UI can badge it Inactive');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 5: matching is case/whitespace-insensitive, ordering is by Sequence
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 5: case-insensitive match, sequence ordering ===');
{
  const res = getProcessesForItem('  tHiNnEr  ', ' 1l ');
  assert(res.success, 'padded/mixed-case lookup succeeds');
  const p1 = byId(res.data, ids.p1);
  assert(p1 && p1.inRecipe === true && p1.qtyPerUnit === 0.25, 'mixed-case + padded name/size still matches P1');

  const seqs = (all || []).map(r => r.sequence);
  const sorted = seqs.slice().sort((a, b) => a - b);
  assert(JSON.stringify(seqs) === JSON.stringify(sorted), 'records sorted by sequence ascending (got ' + JSON.stringify(seqs) + ')');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 6: an item in no recipe at all, and the blank-name guard
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 6: unused item and blank-name guard ===');
{
  const res = getProcessesForItem('Nonexistent Widget', '');
  assert(res.success, 'unknown item is a success, not an error');
  assert((res.data || []).length >= 4, 'still returns the full process list');
  assert((res.data || []).every(r => r.inRecipe === false && r.colorVariants.length === 0), 'nothing reported as in-recipe');

  const blank = getProcessesForItem('   ', '');
  assert(blank.success === false, 'blank item name is rejected');
}

// ═════════════════════════════════════════════════════════════════════════
// saveItemProcessMappings — the write half
// ═════════════════════════════════════════════════════════════════════════

const compSheet = () => ss.getSheetByName(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);

// Every Process Components row, as objects, for whole-sheet assertions.
function allComponentRows() {
  const sheet = compSheet();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, 10).getValues().map(r => ({
    processId: String(r[PROCESS_COMPONENTS_COL.PROCESS_ID - 1] || '').trim(),
    itemName: String(r[PROCESS_COMPONENTS_COL.ITEM_NAME - 1] || '').trim(),
    size: String(r[PROCESS_COMPONENTS_COL.SIZE - 1] || '').trim(),
    narration: String(r[PROCESS_COMPONENTS_COL.NARRATION - 1] || '').trim(),
    qtyPerUnit: r[PROCESS_COMPONENTS_COL.QTY_PER_UNIT - 1],
    remarks: String(r[PROCESS_COMPONENTS_COL.REMARKS - 1] || '').trim(),
    sourceType: String(r[PROCESS_COMPONENTS_COL.SOURCE_TYPE - 1] || '').trim(),
    colorGroup: String(r[PROCESS_COMPONENTS_COL.COLOR_GROUP - 1] || '').trim(),
    colorAxis: String(r[PROCESS_COMPONENTS_COL.COLOR_AXIS - 1] || '').trim(),
    unit: String(r[PROCESS_COMPONENTS_COL.UNIT - 1] || '').trim()
  }));
}
function findRow(processId, itemName, size, colorGroup) {
  return allComponentRows().find(r =>
    r.processId === processId && r.itemName === itemName &&
    r.size === (size || '') && r.colorGroup === colorGroup) || null;
}

// ─────────────────────────────────────────────────────────────────────────
// Test 7: add to a new process, update an existing one, in one call
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 7: add + update in one bulk call ===');
{
  const before = allComponentRows().length;
  const res = saveItemProcessMappings('Thinner', '1L', [
    // P1 already has a COMMON row at 0.25 -> update, and set a unit.
    { processId: ids.p1, inRecipe: true, qtyPerUnit: 0.9, unit: 'Kg', remarks: 'bumped' },
    // P4 has only a Green color row -> this ADDS a COMMON row alongside it.
    { processId: ids.p4, inRecipe: true, qtyPerUnit: 3, unit: '', remarks: '' }
  ]);
  assert(res.success, 'save succeeds: ' + res.message);
  assert(res.data.updated === 1, 'reports 1 updated (got ' + res.data.updated + ')');
  assert(res.data.added === 1, 'reports 1 added (got ' + res.data.added + ')');
  assert(res.data.removed === 0, 'reports 0 removed (got ' + res.data.removed + ')');
  assert(allComponentRows().length === before + 1, 'exactly one row added to the sheet');

  const p1Row = findRow(ids.p1, 'Thinner', '1L', 'COMMON');
  assert(p1Row && p1Row.qtyPerUnit === 0.9, 'P1 qty patched to 0.9 (got ' + (p1Row && p1Row.qtyPerUnit) + ')');
  assert(p1Row && p1Row.unit === 'Kg', 'P1 unit patched to Kg (got "' + (p1Row && p1Row.unit) + '")');
  assert(p1Row && p1Row.remarks === 'bumped', 'P1 remarks patched (got "' + (p1Row && p1Row.remarks) + '")');

  const p4Row = findRow(ids.p4, 'Thinner', '1L', 'COMMON');
  assert(p4Row && p4Row.sourceType === 'ITEM', 'added row is SOURCE_TYPE ITEM (got "' + (p4Row && p4Row.sourceType) + '")');
  assert(p4Row && p4Row.colorGroup === 'COMMON', 'added row is COLOR_GROUP COMMON (got "' + (p4Row && p4Row.colorGroup) + '")');
  assert(p4Row && p4Row.colorAxis === '', 'added row carries no color axis');
  assert(p4Row && p4Row.qtyPerUnit === 3, 'added row qty is 3 (got ' + (p4Row && p4Row.qtyPerUnit) + ')');

  // The Green color row it sits beside must be untouched.
  const green = findRow(ids.p4, 'Thinner', '1L', 'Green');
  assert(green && green.qtyPerUnit === 1.2, 'P4 Green color row survives untouched (got ' + (green && green.qtyPerUnit) + ')');

  // ...and the item's OTHER process rows, plus other items, are untouched.
  const brush = findRow(ids.p1, 'Brush', '', 'COMMON');
  assert(brush && brush.qtyPerUnit === 1, 'another item in the same process untouched');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 8: untick removes ONLY the COMMON ITEM row
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 8: untick removes only the COMMON row ===');
{
  const res = saveItemProcessMappings('Thinner', '1L', [
    { processId: ids.p2, inRecipe: false }
  ]);
  assert(res.success, 'removal succeeds: ' + res.message);
  assert(res.data.removed === 1, 'reports 1 removed (got ' + res.data.removed + ')');

  assert(findRow(ids.p2, 'Thinner', '1L', 'COMMON') === null, 'P2 COMMON row gone');
  const red = findRow(ids.p2, 'Thinner', '1L', 'Red');
  const blue = findRow(ids.p2, 'Thinner', '1L', 'Blue');
  assert(red && red.qtyPerUnit === 0.6, 'P2 Red color row survives (got ' + (red && red.qtyPerUnit) + ')');
  assert(blue && blue.qtyPerUnit === 0.5, 'P2 Blue color row survives (got ' + (blue && blue.qtyPerUnit) + ')');

  const back = byId(res.data.processes, ids.p2);
  assert(back && back.inRecipe === false, 'returned fresh state shows P2 no longer in recipe');
  assert(back && back.colorVariants.length === 2, 'returned fresh state still lists both color variants');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 9: multi-row removal deletes the right rows (bottom-up ordering)
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 9: removing several processes at once ===');
{
  // Put Primer 2L into three processes, then remove two of them in one call.
  let res = saveItemProcessMappings('Primer', '2L', [
    { processId: ids.p1, inRecipe: true, qtyPerUnit: 1 },
    { processId: ids.p2, inRecipe: true, qtyPerUnit: 2 },
    { processId: ids.p3, inRecipe: true, qtyPerUnit: 3 }
  ]);
  assert(res.success && res.data.added === 3, 'Primer added to 3 processes (got ' + res.data.added + ')');

  const otherItemsBefore = allComponentRows().filter(r => r.itemName !== 'Primer').length;

  res = saveItemProcessMappings('Primer', '2L', [
    { processId: ids.p1, inRecipe: false },
    { processId: ids.p3, inRecipe: false }
  ]);
  assert(res.success && res.data.removed === 2, 'two Primer rows removed (got ' + res.data.removed + ')');

  assert(findRow(ids.p1, 'Primer', '2L', 'COMMON') === null, 'P1 Primer row gone');
  assert(findRow(ids.p3, 'Primer', '2L', 'COMMON') === null, 'P3 Primer row gone');
  const kept = findRow(ids.p2, 'Primer', '2L', 'COMMON');
  assert(kept && kept.qtyPerUnit === 2, 'P2 Primer row (not submitted for removal) kept its qty 2 (got ' + (kept && kept.qtyPerUnit) + ')');

  const otherItemsAfter = allComponentRows().filter(r => r.itemName !== 'Primer').length;
  assert(otherItemsBefore === otherItemsAfter,
    'no unrelated row was shifted away by the deletes (' + otherItemsBefore + ' vs ' + otherItemsAfter + ')');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 10: a POOL row of the same name blocks adding an ITEM row
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 10: POOL collision is refused, not silently duplicated ===');
{
  // P3 holds a COMMON POOL row for Thinner 1L (see fixture). Adding an ITEM
  // row beside it would build a process _findDuplicateComponent can no
  // longer save, since its key ignores sourceType.
  const before = allComponentRows().length;
  const res = saveItemProcessMappings('Thinner', '1L', [
    { processId: ids.p3, inRecipe: true, qtyPerUnit: 5 }
  ]);
  assert(res.success === false, 'add is refused');
  assert(/Warehouse Pool component/i.test(res.message || ''), 'message explains the pool collision (got "' + res.message + '")');
  assert(allComponentRows().length === before, 'no row was written');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 11: validation rejects the whole batch, writing nothing
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 11: validation is all-or-nothing ===');
{
  const before = JSON.stringify(allComponentRows());

  let res = saveItemProcessMappings('Thinner', '1L', [
    { processId: ids.p1, inRecipe: true, qtyPerUnit: 7 },      // would succeed
    { processId: ids.p4, inRecipe: true, qtyPerUnit: 0 }        // invalid
  ]);
  assert(res.success === false, 'qty 0 rejects the batch');
  assert(/greater than 0/i.test(res.message || ''), 'message tells the user to untick instead (got "' + res.message + '")');
  assert(JSON.stringify(allComponentRows()) === before, 'the valid row in the same batch was NOT written');

  res = saveItemProcessMappings('Thinner', '1L', [
    { processId: 'PRC-NOPE', inRecipe: true, qtyPerUnit: 1 }
  ]);
  assert(res.success === false && /no longer exists/i.test(res.message || ''), 'unknown process id rejected');

  res = saveItemProcessMappings('Thinner', '1L', [
    { processId: ids.p1, inRecipe: true, qtyPerUnit: 1 },
    { processId: ids.p1, inRecipe: false }
  ]);
  assert(res.success === false && /twice/i.test(res.message || ''), 'the same process submitted twice is rejected');

  res = saveItemProcessMappings('Ghost Item', '', [
    { processId: ids.p1, inRecipe: true, qtyPerUnit: 1 }
  ]);
  assert(res.success === false && /Items Master/i.test(res.message || ''), 'item missing from Items Master is rejected');

  assert(JSON.stringify(allComponentRows()) === before, 'sheet unchanged after every rejected call');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 12: removing from a process that has lots warns but still applies
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 12: production-lot warning on removal ===');
{
  // Give P1 a Production lot.
  let prod = ss.getSheetByName(APP_CONFIG.SHEETS.PRODUCTION);
  if (!prod) prod = ss.addSheet(APP_CONFIG.SHEETS.PRODUCTION);
  if (prod.getLastRow() < 1) {
    const header = new Array(PRODUCTION_COL.LOT_NUMBER).fill('');
    header[PRODUCTION_COL.PROCESS_ID - 1] = 'Process ID';
    prod.appendRow(header);
  }
  const lot = new Array(PRODUCTION_COL.LOT_NUMBER).fill('');
  lot[PRODUCTION_COL.PROCESS_ID - 1] = ids.p1;
  lot[PRODUCTION_COL.QTY - 1] = 10;
  prod.appendRow(lot);

  const res = saveItemProcessMappings('Thinner', '1L', [
    { processId: ids.p1, inRecipe: false }
  ]);
  assert(res.success, 'removal from a process with lots still succeeds: ' + res.message);
  assert(res.data.removed === 1, 'the row was actually removed (got ' + res.data.removed + ')');
  assert(findRow(ids.p1, 'Thinner', '1L', 'COMMON') === null, 'P1 COMMON row gone');
  assert((res.data.warnings || []).length === 1, 'one warning returned (got ' + (res.data.warnings || []).length + ')');
  assert(/production lots/i.test(res.data.warnings[0] || ''), 'warning mentions production lots (got "' + res.data.warnings[0] + '")');
  assert(/Existing lots keep/i.test(res.data.warnings[0] || ''), 'warning reassures that past lots are unchanged');

  // A process with no lots must NOT warn.
  const res2 = saveItemProcessMappings('Primer', '2L', [{ processId: ids.p2, inRecipe: false }]);
  assert(res2.success && (res2.data.warnings || []).length === 0, 'no warning for a process with no lots');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 13: a no-op submission is reported as such
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 13: no-op submissions ===');
{
  const res = saveItemProcessMappings('Thinner', '1L', [
    { processId: ids.p1, inRecipe: false }   // already not in recipe
  ]);
  assert(res.success, 'no-op succeeds');
  assert(res.data.added === 0 && res.data.updated === 0 && res.data.removed === 0, 'nothing counted');
  assert(/No changes/i.test(res.message || ''), 'message says no changes (got "' + res.message + '")');

  const empty = saveItemProcessMappings('Thinner', '1L', []);
  assert(empty.success === false, 'an empty mapping list is rejected rather than silently doing nothing');
}

// ─────────────────────────────────────────────────────────────────────────
console.log(failures === 0 ? '\nAll assertions passed.' : `\n${failures} assertion(s) FAILED.`);
process.exit(failures === 0 ? 0 : 1);
