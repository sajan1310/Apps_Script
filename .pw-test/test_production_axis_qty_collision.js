/**
 * Standalone Node harness (mirrors test_process_color_links.js) exercising
 * saveProduction's Lot Qty derivation for a multi-Color-Axis lot.
 *
 * Reproduces the reported bug: a lot with two independent Color Axes (e.g.
 * a "Rim Color" axis and a separate "Frame Color" axis) whose color sets
 * happen to share one identical color name (both independently having their
 * own "Purple" — each axis's colors come from its own pool item's own
 * independent color history, with no cross-axis uniqueness guarantee, see
 * computeColorAxesForProcess) got the wrong Lot Qty: the non-primary axis's
 * "Purple" row was wrongly counted a second time on top of the real primary
 * total, because the qty derivation matched a checked color against the
 * primary axis purely by color NAME, not by which axis it was actually
 * checked in. Screenshot case: Rim axis (Blue:6, Pink:12, Purple:6, Red:12
 * -> 36) + Frame axis's colliding "Purple:6" row = 42 shown, instead of 36.
 *
 * Run: node .pw-test/test_production_axis_qty_collision.js
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

const ss = new FakeSpreadsheet();

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

const files = [
  'config.js',
  'utils.js',
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
  global.COMPONENT_SOURCE_TYPES = COMPONENT_SOURCE_TYPES;
  global.COMPONENT_COLOR_GROUP_COMMON = COMPONENT_COLOR_GROUP_COMMON;
`, ctx, { filename: 'expose.js' });

const {
  APP_CONFIG, COMPONENT_SOURCE_TYPES, COMPONENT_COLOR_GROUP_COMMON,
  saveProcess, saveProduction, getProductionData
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

function createProcess(name, prefix, outputItemName, components) {
  const res = saveProcess({
    processName: name,
    sequence: 1,
    lotPrefix: prefix,
    outputItemName: outputItemName,
    isFinalStage: false,
    active: true,
    remarks: '',
    components: JSON.stringify(components || [])
  });
  assert(res.success, `saveProcess(${name}) succeeds: ${res.message}`);
  return res.data && res.data.processId;
}

function seedPoolColors(outputItemName, processId, colors) {
  const poolSheet = ss.getSheetByName(APP_CONFIG.SHEETS.WAREHOUSE_POOL) || (function () {
    const s = ss.insertSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
    s.appendRow(['Output Item Name', 'Process ID', 'Product Tag', 'Produced Qty', 'Consumed Qty', 'Available Qty', 'Color']);
    return s;
  })();
  colors.forEach(color => {
    poolSheet.appendRow([outputItemName, processId, '', 100, 0, 100, color]);
  });
}

console.log('\n=== Setup: two pool items feeding two independent Color Axes, sharing the color name "Purple" ===');
const rimId = createProcess('Rim Component', 'ZR', 'Rim Component', []);
const frameId = createProcess('Frame Component', 'ZF', 'Frame Component', []);
const fittingId = createProcess('Fitting Frame', 'FZ', 'Fitted Frame', [
  { itemName: 'Rim Component', sourceType: COMPONENT_SOURCE_TYPES.POOL, qtyPerUnit: 1, colorGroup: COMPONENT_COLOR_GROUP_COMMON },
  { itemName: 'Frame Component', sourceType: COMPONENT_SOURCE_TYPES.POOL, qtyPerUnit: 1, colorGroup: COMPONENT_COLOR_GROUP_COMMON }
]);

// Rim axis: Blue, Pink, Purple, Red. Frame axis: Blue-White, Pink-SeaGreen,
// Purple, Red-Yellow — "Purple" exists in BOTH axes' own independent pool
// color history, exactly the collision that triggered the bug.
seedPoolColors('Rim Component', rimId, ['Blue', 'Pink', 'Purple', 'Red']);
seedPoolColors('Frame Component', frameId, ['Blue-White', 'Pink-SeaGreen', 'Purple', 'Red-Yellow']);

// Mirrors exactly what Script.html's getCheckedColorQtys() sends: the Rim
// axis is Primary (countsTowardTotal true, axisKey = its own axis key), the
// Frame axis is non-primary and auto-synced to mirror the matching Rim
// row's qty (countsTowardTotal false, its own DIFFERENT axisKey) — see
// Script.html's _syncMatchingNonPrimaryRows/handleColorCheckToggle.
const colorBreakdown = [
  { color: 'Blue', qty: 6, isCustom: false, countsTowardTotal: true, axisKey: 'pool:rim component' },
  { color: 'Pink', qty: 12, isCustom: false, countsTowardTotal: true, axisKey: 'pool:rim component' },
  { color: 'Purple', qty: 6, isCustom: false, countsTowardTotal: true, axisKey: 'pool:rim component' },
  { color: 'Red', qty: 12, isCustom: false, countsTowardTotal: true, axisKey: 'pool:rim component' },
  { color: 'Blue-White', qty: 6, isCustom: false, countsTowardTotal: false, axisKey: 'pool:frame component' },
  { color: 'Pink-SeaGreen', qty: 12, isCustom: false, countsTowardTotal: false, axisKey: 'pool:frame component' },
  { color: 'Purple', qty: 6, isCustom: false, countsTowardTotal: false, axisKey: 'pool:frame component' },
  { color: 'Red-Yellow', qty: 12, isCustom: false, countsTowardTotal: false, axisKey: 'pool:frame component' }
];

console.log('\n=== Test: Lot Qty must be the Rim (primary) axis total only (36), not inflated by the colliding Frame-axis "Purple" row (42) or the full sum of both axes (72) ===');
{
  const res = saveProduction({
    processId: fittingId,
    assignedTo: 'Test Contractor',
    primaryColorAxis: 'Rim Component',
    colorBreakdown: JSON.stringify(colorBreakdown),
    componentsConsumed: JSON.stringify([
      { itemName: 'Rim Component', size: '', color: '', sourceType: 'POOL', qty: 36, colorGroup: 'COMMON' },
      { itemName: 'Frame Component', size: '', color: '', sourceType: 'POOL', qty: 36, colorGroup: 'COMMON' }
    ])
  });
  assert(res.success, 'saveProduction succeeds: ' + res.message);

  const lots = getProductionData().data || [];
  const lot = lots.find(l => l.processId === fittingId);
  assert(!!lot, 'saved lot is readable back');
  assert(lot && lot.qty === 36, `Lot Qty is 36 (Rim axis only: 6+12+6+12), not inflated by the colliding Frame-axis "Purple" (got ${lot && lot.qty})`);
}

console.log('\n' + (failures === 0 ? 'ALL TESTS PASSED' : failures + ' TEST(S) FAILED'));
process.exit(failures === 0 ? 0 : 1);
