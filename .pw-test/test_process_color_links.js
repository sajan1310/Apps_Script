/**
 * Standalone Node harness (mirrors test_process_color_groups.js) exercising
 * the new Process Color Links feature: linking two processes' colors so a
 * downstream process's checklist pairs them instead of cross-multiplying.
 *
 * Run: node .pw-test/test_process_color_links.js
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
  global.WAREHOUSE_POOL_COL = WAREHOUSE_POOL_COL;
  global.COMPONENT_SOURCE_TYPES = COMPONENT_SOURCE_TYPES;
  global.COMPONENT_COLOR_GROUP_COMMON = COMPONENT_COLOR_GROUP_COMMON;
`, ctx, { filename: 'expose.js' });

const {
  APP_CONFIG, WAREHOUSE_POOL_COL, COMPONENT_SOURCE_TYPES, COMPONENT_COLOR_GROUP_COMMON,
  saveProcess, getProcessColorGroups, getProcessColorLinksData, deleteProcess
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
// Seed 4 processes: A (2 colors), B (6 "X/White" colors), D (6 plain
// colors), C (downstream, consumes A+B+D as POOL components).
// ─────────────────────────────────────────────────────────────────────────
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
    poolSheet.appendRow([outputItemName, processId, '', 10, 0, 10, color]);
  });
}

console.log('\n=== Setup: create processes A, B, D, C ===');
const aId = createProcess('Process A', 'ZA', 'Item A', []);
const bId = createProcess('Process B', 'ZB', 'Item B', []);
const dId = createProcess('Process D', 'ZD', 'Item D', []);
const cId = createProcess('Process C', 'ZC', 'Item C', [
  { itemName: 'Item A', sourceType: COMPONENT_SOURCE_TYPES.POOL, qtyPerUnit: 1, colorGroup: COMPONENT_COLOR_GROUP_COMMON },
  { itemName: 'Item B', sourceType: COMPONENT_SOURCE_TYPES.POOL, qtyPerUnit: 1, colorGroup: COMPONENT_COLOR_GROUP_COMMON },
  { itemName: 'Item D', sourceType: COMPONENT_SOURCE_TYPES.POOL, qtyPerUnit: 1, colorGroup: COMPONENT_COLOR_GROUP_COMMON }
]);

seedPoolColors('Item A', aId, ['Black', 'BCP']);
seedPoolColors('Item B', bId, ['Orange/White', 'Blue/White', 'Red/White', 'Pink/White', 'Seagreen/White', 'Purple/White']);
seedPoolColors('Item D', dId, ['Red', 'Pink', 'Blue', 'Purple', 'Seagreen', 'Orange']);

// ─────────────────────────────────────────────────────────────────────────
// Process C has 3 independent pool axes (A, B, D) and no primaryColorAxis
// configured. getProcessColorGroups must return the flat UNION of each
// axis's own colors here, NOT a cross product — Script.html's
// renderGroupedColorChecklist (the actual "Colors to Produce" checklist the
// operator fills in) already renders one independent checkbox group per
// axis the moment 2+ axes are detected, regardless of whether a Primary
// Axis has been saved yet. A cross-multiplied validation list here would
// reject every real single-axis color the operator can actually check —
// this was a real production bug (see computeColorGroupsForProcess).
console.log('\n=== Test 1: no links -> 3 independent axes, flat union (2+6+6=14), NOT a cross product ===');
{
  const res = getProcessColorGroups(cId);
  assert(res.success, 'getProcessColorGroups(C) succeeds: ' + res.message);
  assert(res.data.length === 14, `14 individual colors with no links (got ${res.data.length}): ${JSON.stringify(res.data)}`);
  assert(res.data.includes('Black') && res.data.includes('BCP'), 'Process A colors present as their own entries');
  assert(res.data.includes('Orange/White') && res.data.includes('Orange'), 'Process B and D colors present as independent entries, not cross-multiplied');
  assert(!res.data.some(c => c.includes(' / ')), 'no cross-multiplied composite strings anywhere (unlinked axes stay independent)');
}

// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 2: link B<->D on all 6 colors -> B and D merge into ONE paired axis, still independent of A (2+6=8) ===');
{
  const pairs = [
    ['Orange/White', 'Orange'], ['Blue/White', 'Blue'], ['Red/White', 'Red'],
    ['Pink/White', 'Pink'], ['Seagreen/White', 'Seagreen'], ['Purple/White', 'Purple']
  ];
  const res = saveProcess({
    processId: bId,
    processName: 'Process B',
    sequence: 1,
    lotPrefix: 'ZB',
    outputItemName: 'Item B',
    isFinalStage: false,
    active: true,
    remarks: '',
    components: JSON.stringify([]),
    colorLinks: JSON.stringify(pairs.map(([myColor, theirColor]) => ({ otherProcessId: dId, myColor, theirColor })))
  });
  assert(res.success, 'saveProcess(B) with colorLinks succeeds: ' + res.message);

  const checklist = getProcessColorGroups(cId);
  assert(checklist.success, 'getProcessColorGroups(C) succeeds after linking');
  assert(checklist.data.length === 8, `8 entries after B<->D link: A's own 2 + the merged pair-axis's own 6 (got ${checklist.data.length}): ${JSON.stringify(checklist.data)}`);
  assert(checklist.data.includes('Black') && checklist.data.includes('BCP'), 'Process A colors still present as their own independent entries (not cross-multiplied with the merged pair axis)');
  assert(checklist.data.includes('Orange/White / Orange'), 'merged pair "Orange/White / Orange" present as its own single axis entry');
  assert(checklist.data.includes('Blue/White / Blue'), 'merged pair "Blue/White / Blue" present as its own single axis entry');
  assert(!checklist.data.some(c => c.includes('Orange/White / Blue')), 'a mismatched pairing like "Orange/White / Blue" must NOT appear');
  assert(!checklist.data.some(c => c.split(' / ').length > 2), 'no entry combines A with the merged pair axis (that would be a 3-way cross product again)');
}

// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 3: getProcessColorLinksData symmetric from both sides ===');
{
  const fromB = getProcessColorLinksData(bId);
  assert(fromB.success && fromB.data.length === 6, `Process B reports 6 links (got ${fromB.success ? fromB.data.length : 'FAILED'})`);
  assert(fromB.data.every(l => l.otherProcessId === dId), 'every link from B points to D');

  const fromD = getProcessColorLinksData(dId);
  assert(fromD.success && fromD.data.length === 6, `Process D reports 6 links (got ${fromD.success ? fromD.data.length : 'FAILED'})`);
  assert(fromD.data.every(l => l.otherProcessId === bId), 'every link from D points to B');
  const orangeFromD = fromD.data.find(l => l.myColor === 'Orange');
  assert(!!orangeFromD && orangeFromD.theirColor === 'Orange/White', 'D->B orientation resolves Orange -> Orange/White (got ' + JSON.stringify(orangeFromD) + ')');
}

// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 4: deleting Process D cascades link cleanup ===');
{
  const delRes = deleteProcess(dId);
  assert(delRes.success, 'deleteProcess(D) succeeds: ' + delRes.message);

  const fromB = getProcessColorLinksData(bId);
  assert(fromB.success && fromB.data.length === 0, `Process B has no dangling links after D deleted (got ${fromB.success ? fromB.data.length : 'FAILED'})`);
}

// ─────────────────────────────────────────────────────────────────────────
console.log('\n' + (failures === 0 ? 'ALL TESTS PASSED' : failures + ' TEST(S) FAILED'));
process.exit(failures === 0 ? 0 : 1);
