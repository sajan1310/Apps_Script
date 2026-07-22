/**
 * Standalone Node harness (same pattern as test_process_color_groups.js) that
 * loads the REAL server-side files and verifies the 2026-07-22 change:
 * computeColorGroupsForProcess (module_process.js) must widen a
 * color-enabled process's "Colors to Produce" list to the FULL Color Master
 * list, not just the colors its own recipe/pool history has actually
 * touched — while a process with NO color dimension at all must still get
 * an empty list (never suddenly grow a checklist just because Color Master
 * is non-empty). See module_process.js's computeColorGroupsForProcess and
 * _getColorMasterNames doc comments.
 *
 * test_process_color_groups.js deliberately does NOT load module_tags.js,
 * so its own "Green" color is correctly rejected there (Color Master widening
 * never kicks in without getColors() available) — that test's assertions
 * must stay untouched. This is a separate file specifically so Color Master
 * IS loaded and the widening path is actually exercised end-to-end.
 *
 * Run: node .pw-test/test_color_master_widening.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

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
    this.rows = [];
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

// Minimal in-memory CacheService — getColors() (module_tags.js) reads/writes
// through this, so it must actually work (not just no-op) for the cache-hit
// path getAllProcessColorGroups relies on to stay cheap in a loop.
class FakeScriptCache {
  constructor() { this.store = {}; }
  get(key) { return Object.prototype.hasOwnProperty.call(this.store, key) ? this.store[key] : null; }
  put(key, value) { this.store[key] = value; }
  removeAll() { this.store = {}; }
}

const ss = new FakeSpreadsheet();
const scriptCache = new FakeScriptCache();

const sandbox = {
  SpreadsheetApp: {
    getActiveSpreadsheet: () => ss,
    flush: () => {}
  },
  CacheService: {
    getScriptCache: () => scriptCache
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
  'module_tags.js',
  'module_process.js',
  'module_production.js',
  'module_warehouse.js'
];

files.forEach(f => {
  const code = fs.readFileSync(path.join(ROOT, f), 'utf8');
  vm.runInContext(code, ctx, { filename: f });
});

vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.PROCESS_COMPONENTS_COL = PROCESS_COMPONENTS_COL;
  global.PRODUCTION_COL = PRODUCTION_COL;
  global.COMPONENT_COLOR_GROUP_COMMON = COMPONENT_COLOR_GROUP_COMMON;
`, ctx, { filename: 'expose.js' });

const {
  APP_CONFIG, saveProcess, getProcessColorGroups, getAllProcessColorGroups,
  saveProduction, saveColor, excludeWarehousePoolColors, includeWarehousePoolColor
} = ctx;

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
console.log('\n=== Setup: seed a 5-name Color Master ===');
['Red', 'Blue', 'Green', 'Orange', 'Yellow'].forEach(name => {
  const res = saveColor({ name });
  assert(res.success, `Color Master "${name}" saved: ` + res.message);
});

// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 1: a color-enabled process widens to the full Color Master ===');
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
      { itemName: 'Red Paint', sourceType: 'ITEM', qtyPerUnit: 2, colorGroup: 'Red' },
      { itemName: 'Blue Paint', sourceType: 'ITEM', qtyPerUnit: 2, colorGroup: 'Blue' }
    ])
  });
  assert(res.success, 'saveProcess succeeds: ' + res.message);
  framePaintingId = res.data && res.data.processId;

  const colorsRes = getProcessColorGroups(framePaintingId);
  assert(colorsRes.success, 'getProcessColorGroups succeeds');
  const expected = ['Blue', 'Green', 'Orange', 'Red', 'Yellow'];
  assert(
    JSON.stringify(colorsRes.data) === JSON.stringify(expected),
    'widened to the full Color Master, not just the 2 recipe-tagged colors (got ' + JSON.stringify(colorsRes.data) + ')'
  );
}

// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 2: a process with no color dimension stays empty (no accidental opt-in) ===');
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
  assert(
    colorsRes.success && colorsRes.data.length === 0,
    'still reports zero color groups despite a non-empty Color Master (got ' + JSON.stringify(colorsRes.data) + ')'
  );
}

// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 3: saveProduction now accepts a Color Master color the recipe never tagged ===');
{
  // Completed (not Pending) — recalculateWarehousePool only credits a real
  // Warehouse Pool bucket for a Completed lot, and Test 5 below needs
  // "Green" to have real bucket history to prove excludeWarehousePoolColors
  // protects it.
  const goodColor = saveProduction({
    processId: framePaintingId,
    qty: 5,
    assignedTo: 'Test Contractor',
    status: 'Completed',
    colorBreakdown: JSON.stringify([{ color: 'Green', qty: 5 }]),
    componentsConsumed: JSON.stringify([
      { itemName: 'Brush', sourceType: 'ITEM', qty: 5 }
    ])
  });
  assert(goodColor.success, '"Green" (Color Master, not recipe-tagged) is now accepted: ' + goodColor.message);

  const stillBad = saveProduction({
    processId: framePaintingId,
    qty: 5,
    assignedTo: 'Test Contractor',
    status: 'Pending',
    colorBreakdown: JSON.stringify([{ color: 'Purple', qty: 5 }]),
    componentsConsumed: JSON.stringify([
      { itemName: 'Brush', sourceType: 'ITEM', qty: 5 }
    ])
  });
  assert(stillBad.success === false, 'a color absent from BOTH the recipe AND Color Master ("Purple") is still rejected');
}

// ─────────────────────────────────────────────────────────────────────────
// 2026-07-22 (later same day): getAllProcessColorGroups no longer widens to
// the full Color Master the way getProcessColorGroups (singular, the
// Production checklist's own call path) deliberately still does — the
// Warehouse Pool breakdown dialog's placeholder rows are fed by this bulk
// variant, and unioning in every Color Master entry there just produced one
// zero-qty placeholder row per unused color per process (real clutter, not
// real flexibility — see feature_axis_color_pairings_and_composite_rename_fix
// / the Phase 3 Warehouse Pool combinations follow-up). `colors` here is now
// recipe/pool-detected colors UNION colors this process has actually logged
// producing (see _getProductionLoggedColorsByProcess - "Green" below is a
// real example: not recipe-tagged, but Test 3 above logged a real Completed
// lot with it) UNION anything explicitly INCLUDEd via includeWarehousePoolColor
// (the "+ Add Combination" escape hatch, still fully independent of this
// narrowing — see Test 6). `removable` stays keyed off baseColors only (not
// logged history) — a logged-but-not-recipe/pool color like "Green" still
// LOOKS removable here, but excludeWarehousePoolColors' own separate
// real-bucket-history guard rejects the actual removal attempt regardless
// (see Test 5 immediately below), so nothing real is ever actually at risk.
console.log('\n=== Test 4: bulk getAllProcessColorGroups is recipe/pool/logged-history-scoped, UNLIKE the single-process (checklist) endpoint ===');
{
  const bulkRes = getAllProcessColorGroups();
  assert(bulkRes.success, 'getAllProcessColorGroups succeeds');
  const expectedColors = ['Blue', 'Green', 'Red']; // 2 recipe-tagged + "Green" from Test 3's logged Completed lot - no Color Master union
  assert(
    JSON.stringify(bulkRes.data[framePaintingId].colors.slice().sort()) === JSON.stringify(expectedColors),
    'bulk variant is recipe-tagged + actually-logged colors only, no blanket Color Master union (got ' + JSON.stringify(bulkRes.data[framePaintingId]) + ')'
  );
  assert(
    JSON.stringify(bulkRes.data[framePaintingId].removable) === JSON.stringify(['Green']),
    '"Green" (logged but not recipe/pool-configured) is the only removable-looking entry (got ' + JSON.stringify(bulkRes.data[framePaintingId].removable) + ')'
  );
  assert(
    Array.isArray(bulkRes.data[packingId].colors) && bulkRes.data[packingId].colors.length === 0,
    'bulk variant still reports zero for the no-color-dimension, never-produced process (got ' + JSON.stringify(bulkRes.data[packingId]) + ')'
  );

  // The single-process (checklist) endpoint is a completely separate call
  // path and must be COMPLETELY UNAFFECTED by the above - still the full
  // widened 5-color Color Master union, exactly as Test 1 already proved.
  const singleRes = getProcessColorGroups(framePaintingId);
  const expectedWidened = ['Blue', 'Green', 'Orange', 'Red', 'Yellow'];
  assert(
    JSON.stringify(singleRes.data) === JSON.stringify(expectedWidened),
    'getProcessColorGroups (singular) still widens to the full Color Master, unaffected by the bulk variant\'s narrowing (got ' + JSON.stringify(singleRes.data) + ')'
  );
}

// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 5: excludeWarehousePoolColors protects recipe-configured and real-history colors ===');
{
  const blockConfigured = excludeWarehousePoolColors(framePaintingId, ['Red']);
  assert(blockConfigured.success === false, '"Red" (recipe-configured) is rejected (got: ' + blockConfigured.message + ')');

  const blockHistory = excludeWarehousePoolColors(framePaintingId, ['Green']);
  assert(blockHistory.success === false, '"Green" (has real production history from Test 3) is rejected (got: ' + blockHistory.message + ')');

  const okRemove = excludeWarehousePoolColors(framePaintingId, ['Orange', 'Yellow']);
  assert(okRemove.success, 'removing pure Color-Master noise ("Orange", "Yellow") succeeds: ' + okRemove.message);

  const afterRemove = getProcessColorGroups(framePaintingId);
  assert(
    !afterRemove.data.includes('Orange') && !afterRemove.data.includes('Yellow'),
    'checklist no longer offers the removed colors (got ' + JSON.stringify(afterRemove.data) + ')'
  );
  assert(
    afterRemove.data.includes('Red') && afterRemove.data.includes('Green') && afterRemove.data.includes('Blue'),
    'every other color is untouched (got ' + JSON.stringify(afterRemove.data) + ')'
  );
}

// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 6: includeWarehousePoolColor adds a custom combo and undoes a prior exclusion ===');
{
  const addCustom = includeWarehousePoolColor(packingId, 'Midnight Purple');
  assert(addCustom.success, 'force-adding a custom color to a plain (no-color-dimension) process succeeds: ' + addCustom.message);

  const packingColors = getProcessColorGroups(packingId);
  assert(
    packingColors.data.length === 1 && packingColors.data[0] === 'Midnight Purple',
    'the plain process now offers exactly its one force-included color (got ' + JSON.stringify(packingColors.data) + ')'
  );

  const undoExclude = includeWarehousePoolColor(framePaintingId, 'Orange');
  assert(undoExclude.success, 're-including a previously-excluded color succeeds: ' + undoExclude.message);
  const afterUndo = getProcessColorGroups(framePaintingId);
  assert(afterUndo.data.includes('Orange'), '"Orange" is back on the checklist after re-including it (got ' + JSON.stringify(afterUndo.data) + ')');

  // The narrower getAllProcessColorGroups (Test 4) must still pick up both
  // manual includes above ("Midnight Purple" on the plain packing process,
  // "Orange" re-included on Frame Painting) - the narrowing only drops the
  // AUTOMATIC Color Master union, never the explicit manual escape hatch.
  const bulkAfterIncludes = getAllProcessColorGroups();
  assert(
    (bulkAfterIncludes.data[packingId].colors || []).includes('Midnight Purple'),
    'bulk variant still reflects a manually-INCLUDEd color on an otherwise-plain process (got ' + JSON.stringify(bulkAfterIncludes.data[packingId]) + ')'
  );
  assert(
    (bulkAfterIncludes.data[framePaintingId].colors || []).includes('Orange'),
    'bulk variant still reflects a manually re-INCLUDEd color (got ' + JSON.stringify(bulkAfterIncludes.data[framePaintingId]) + ')'
  );
}

console.log('\n' + (failures === 0 ? 'ALL TESTS PASSED' : failures + ' TEST(S) FAILED'));
process.exit(failures === 0 ? 0 : 1);
