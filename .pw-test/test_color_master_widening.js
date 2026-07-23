/**
 * Standalone Node harness (same pattern as test_process_color_groups.js)
 * that loads the REAL server-side files. Originally verified a 2026-07-22
 * change where a color-enabled process's "Colors to Produce" list widened
 * to the FULL Color Master list. That widening was deliberately REVERSED
 * later the same day, once it became clear it let a color only ever used
 * by one process (e.g. Painted Mudguard) "reflect" onto a completely
 * unrelated process's checklist (e.g. Fitted Frame) just because both
 * pull from the same global Color Master — see
 * _computeKnownColorsForProcess (module_process.js), now the single shared
 * definition of "known colors" both getProcessColorGroups (singular) and
 * getAllProcessColorGroups (bulk, Warehouse Pool dialog) use: scoped to
 * THAT process's own recipe-tagged + pool-detected + actually-logged
 * Production history + manually-INCLUDEd colors, generic across every
 * process, and process identity is the only scope boundary (not Model,
 * not Process Type). A process with NO color dimension at all must still
 * get an empty list either way (never suddenly grow a checklist just
 * because Color Master is non-empty).
 *
 * test_process_color_groups.js deliberately does NOT load module_tags.js,
 * so its own "Green" color is correctly rejected there regardless of any
 * widening — that test's assertions are untouched by any of this.
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
  // Missing before: invalidateListCache (utils.js) calls cache.remove(key)
  // on every process/color/etc. write path - without this, that call threw
  // (swallowed by invalidateListCache's own try/catch), silently leaving
  // getProcessData's cached process list stale for the rest of the test
  // run the moment a SECOND getAllProcessColorGroups()/getProcessData()
  // call happened after a new process was created post-cache-population.
  // The real GAS CacheService has .remove() - this mock just never did.
  remove(key) { delete this.store[key]; }
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
console.log('\n=== Test 1: a color-enabled process stays scoped to its OWN recipe-tagged colors, no Color Master union ===');
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
  const expected = ['Blue', 'Red']; // just the 2 recipe-tagged colors - "Green"/"Orange"/"Yellow" exist in Color Master but were never touched by this process
  assert(
    JSON.stringify(colorsRes.data) === JSON.stringify(expected),
    'scoped to the 2 recipe-tagged colors only, no blanket Color Master union (got ' + JSON.stringify(colorsRes.data) + ')'
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
// 2026-07-22: getAllProcessColorGroups (bulk, Warehouse Pool dialog) and
// getProcessColorGroups (singular, Production checklist) now share the
// exact same per-process-scoped definition of "known colors" —
// _computeKnownColorsForProcess (module_process.js): recipe/pool-detected
// colors UNION colors this process has actually logged producing (see
// _getProductionLoggedColorsByProcess - "Green" below is a real example:
// not recipe-tagged, but Test 3 above logged a real Completed lot with it)
// UNION anything explicitly INCLUDEd via includeWarehousePoolColor (the
// "+ Add Combination" escape hatch — see Test 6), MINUS EXCLUDEs. No
// blanket Color Master union anywhere anymore, on either endpoint —
// unioning in every Color Master entry regardless of relevance was both
// real clutter (Warehouse Pool placeholder rows) AND real cross-process
// bleed-through (a color only ever used by one process showing up on an
// unrelated process's checklist purely because Color Master is global —
// see Test 7 below for a direct reproduction of that exact scenario).
// `removable` stays keyed off baseColors only (not logged history) — a
// logged-but-not-recipe/pool color like "Green" still LOOKS removable
// here, but excludeWarehousePoolColors' own separate real-bucket-history
// guard rejects the actual removal attempt regardless (see Test 5
// immediately below), so nothing real is ever actually at risk.
console.log('\n=== Test 4: bulk getAllProcessColorGroups and the single-process (checklist) endpoint now agree - both recipe/pool/logged-history-scoped ===');
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

  // The single-process (checklist) endpoint now computes the exact same
  // scoped set for this process - no more "widened" vs "narrow" split
  // between the two call paths.
  const singleRes = getProcessColorGroups(framePaintingId);
  assert(
    JSON.stringify(singleRes.data) === JSON.stringify(expectedColors),
    'getProcessColorGroups (singular) now matches the bulk variant exactly - same scoped colors, no Color Master union (got ' + JSON.stringify(singleRes.data) + ')'
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

// ─────────────────────────────────────────────────────────────────────────
// Direct reproduction of the reported scenario: "Fitting Frame Crysta" and
// "Painted Mudguard Crysta" are two separate processed items (same Model,
// different processes/recipes) - Painted Mudguard's own logged colors must
// never reflect onto Fitting Frame's checklist, and vice versa, purely
// because both draw from the same global Color Master. This is exactly
// what _computeKnownColorsForProcess's per-processId scoping guarantees,
// generic for every process pair, not special-cased to this one example.
console.log('\n=== Test 7: two unrelated processes (same Model, different recipes) never see each other\'s logged colors ===');
{
  const frameRes = saveProcess({
    processName: 'Fitting Frame Crysta',
    sequence: 5,
    lotPrefix: 'FFC',
    outputItemName: 'Fitted Frame Crysta Output',
    isFinalStage: false,
    active: true,
    remarks: '',
    components: JSON.stringify([{ itemName: 'Frame Bolt', sourceType: 'ITEM', qtyPerUnit: 4, colorGroup: 'COMMON' }])
  });
  const mudguardRes = saveProcess({
    processName: 'Painted Mudguard Crysta',
    sequence: 6,
    lotPrefix: 'PMC',
    outputItemName: 'Painted Mudguard Crysta Output',
    isFinalStage: false,
    active: true,
    remarks: '',
    components: JSON.stringify([{ itemName: 'Mudguard Bracket', sourceType: 'ITEM', qtyPerUnit: 2, colorGroup: 'COMMON' }])
  });
  const frameId = frameRes.data && frameRes.data.processId;
  const mudguardId = mudguardRes.data && mudguardRes.data.processId;
  assert(!!frameId && !!mudguardId, 'both processes created (got frameId=' + frameId + ', mudguardId=' + mudguardId + ')');

  const frameLot = saveProduction({
    processId: frameId,
    assignedTo: 'Test Contractor',
    status: 'Completed',
    colorBreakdown: JSON.stringify([{ color: 'Steel Silver', qty: 4, isCustom: true }]),
    componentsConsumed: JSON.stringify([{ itemName: 'Frame Bolt', sourceType: 'ITEM', qty: 16, colorGroup: 'COMMON' }])
  });
  const mudguardLot = saveProduction({
    processId: mudguardId,
    assignedTo: 'Test Contractor',
    status: 'Completed',
    colorBreakdown: JSON.stringify([{ color: 'Forest Green', qty: 3, isCustom: true }]),
    componentsConsumed: JSON.stringify([{ itemName: 'Mudguard Bracket', sourceType: 'ITEM', qty: 6, colorGroup: 'COMMON' }])
  });
  assert(frameLot.success, 'Fitting Frame Crysta lot ("Steel Silver") saves: ' + frameLot.message);
  assert(mudguardLot.success, 'Painted Mudguard Crysta lot ("Forest Green") saves: ' + mudguardLot.message);

  const frameChecklist = getProcessColorGroups(frameId);
  const mudguardChecklist = getProcessColorGroups(mudguardId);
  assert(frameChecklist.data.includes('Steel Silver'), 'Fitting Frame\'s own checklist includes its own logged color "Steel Silver" (got ' + JSON.stringify(frameChecklist.data) + ')');
  assert(!frameChecklist.data.includes('Forest Green'), 'Fitting Frame\'s checklist does NOT include Painted Mudguard\'s "Forest Green" (got ' + JSON.stringify(frameChecklist.data) + ')');
  assert(mudguardChecklist.data.includes('Forest Green'), 'Painted Mudguard\'s own checklist includes its own logged color "Forest Green" (got ' + JSON.stringify(mudguardChecklist.data) + ')');
  assert(!mudguardChecklist.data.includes('Steel Silver'), 'Painted Mudguard\'s checklist does NOT include Fitting Frame\'s "Steel Silver" (got ' + JSON.stringify(mudguardChecklist.data) + ')');

  // Same isolation must hold on the bulk (Warehouse Pool dialog) endpoint too.
  const bulk = getAllProcessColorGroups();
  assert(!(bulk.data[frameId].colors || []).includes('Forest Green'), 'bulk variant: Fitting Frame does not see Painted Mudguard\'s color either');
  assert(!(bulk.data[mudguardId].colors || []).includes('Steel Silver'), 'bulk variant: Painted Mudguard does not see Fitting Frame\'s color either');
}

console.log('\n' + (failures === 0 ? 'ALL TESTS PASSED' : failures + ' TEST(S) FAILED'));
process.exit(failures === 0 ? 0 : 1);
