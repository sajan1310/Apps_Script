/**
 * Standalone Node harness (same mock-SpreadsheetApp pattern as
 * test_process_color_groups.js) exercising the Color Axes feature:
 * computeColorAxesForProcess / getProcessColorAxes, the Primary Color Axis
 * persisted on Process Master, and saveProduction's fix for the quantity
 * double-count bug — checking a "Rim Color" row (the real output batch) AND
 * a "Mudguard Color" row (a same-batch component tag, not an extra unit)
 * for one lot must total the PRIMARY axis's quantity only, not the sum of
 * both, exactly reproducing the "Red-White: 10 + Red: 10 -> 20" bug report.
 *
 * Run: node .pw-test/test_color_axes.js
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
  insertColumnsAfter(afterPosition, howMany) {
    this.rows.forEach(row => { const blanks = new Array(howMany).fill(''); row.splice(afterPosition, 0, ...blanks); });
  }
}

class FakeSpreadsheet {
  constructor() { this.sheets = {}; }
  getSheetByName(name) { return this.sheets[name] || null; }
  addSheet(name) { const s = new FakeSheet(name); this.sheets[name] = s; return s; }
  insertSheet(name) { return this.addSheet(name); }
}

const ss = new FakeSpreadsheet();

const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush: () => {} },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  console,
  Logger: { log: () => {} },
  Utilities: { getUuid: () => 'uuid-' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) }
};
sandbox.global = sandbox;
const ctx = vm.createContext(sandbox);

['config.js', 'utils.js', 'module_process.js', 'module_production.js', 'module_warehouse.js'].forEach(f => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), ctx, { filename: f });
});

vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.PROCESS_COL = PROCESS_COL;
`, ctx, { filename: 'expose.js' });

const {
  APP_CONFIG, PROCESS_COL,
  saveProcess, getProcessData, getProcessColorAxes, getProcessColorGroups,
  saveProduction, getProductionData
} = ctx;

let failures = 0;
function assert(cond, msg) {
  if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); }
}

// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Setup: process with 2 independent tag-based Color Axes, Rim Color primary ===');
let processId;
{
  const res = saveProcess({
    processName: 'Fitted Frame Assembly',
    sequence: 1,
    lotPrefix: 'FFA',
    outputItemName: 'Fitted Frame Test Output',
    isFinalStage: false,
    active: true,
    remarks: '',
    primaryColorAxis: 'Rim Color',
    components: JSON.stringify([
      { itemName: 'Assembly Screws', sourceType: 'ITEM', qtyPerUnit: 4, colorGroup: 'COMMON' },
      { itemName: 'Painted Rim - Blue-White', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Blue-White', colorAxis: 'Rim Color' },
      { itemName: 'Painted Rim - Red-White', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Red-White', colorAxis: 'Rim Color' },
      { itemName: 'Mudguard - Blue', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Blue', colorAxis: 'Mudguard Color' },
      { itemName: 'Mudguard - Red', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Red', colorAxis: 'Mudguard Color' }
    ])
  });
  assert(res.success, 'saveProcess succeeds: ' + res.message);
  processId = res.data && res.data.processId;
  assert(!!processId, 'processId returned (got "' + processId + '")');

  const procSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PROCESS_MASTER);
  // Grew to 11 when Dispatch Differentiator was added (PROCESS_COL.DISPATCH_DIFFERENTIATOR)
  // — pinned to the constant so the next column addition updates this by itself.
  assert(procSheet.getLastColumn() === PROCESS_COL.DISPATCH_DIFFERENTIATOR,
    'Process Master sheet spans every defined column (got ' + procSheet.getLastColumn() + ')');
  assert(procSheet.getRange(1, 10, 1, 1).getValue() === 'Primary Color Axis', 'column 10 header is "Primary Color Axis"');

  // initProcessMasterSheet seeds 4 default rows on first creation, so this
  // process isn't necessarily row 2 — look its row up by ID instead of
  // assuming a fixed position.
  const allProcs = getProcessData(false);
  const ownRow = (allProcs.data || []).find(p => p.processId === processId);
  assert(!!ownRow, 'own process row found via getProcessData');
  assert(ownRow && ownRow.primaryColorAxis === 'Rim Color', 'saved row carries primaryColorAxis "Rim Color" (got "' + (ownRow && ownRow.primaryColorAxis) + '")');

  // rowIdx is no longer part of getProcessData()'s response (trimmed as an
  // unused field — every real consumer keys off processId instead), so
  // find the raw sheet row directly for this sheet-level assertion.
  const idCol = procSheet.getRange(2, PROCESS_COL.PROCESS_ID, procSheet.getLastRow() - 1, 1).getValues();
  const ownRawRow = idCol.findIndex(r => r[0] === processId) + 2;
  assert(procSheet.getRange(ownRawRow, PROCESS_COL.PRIMARY_COLOR_AXIS, 1, 1).getValue() === 'Rim Color', 'raw sheet cell also carries "Rim Color"');
}

console.log('\n=== Test 1: getProcessColorAxes returns 2 independent groups, not cross-multiplied ===');
{
  const res = getProcessColorAxes(processId);
  assert(res.success, 'getProcessColorAxes succeeds: ' + res.message);
  const axes = (res.data && res.data.axes) || [];
  assert(axes.length === 2, 'exactly 2 axes reported (got ' + axes.length + ')');

  const rim = axes.find(a => a.label === 'Rim Color');
  const mud = axes.find(a => a.label === 'Mudguard Color');
  assert(!!rim, 'Rim Color axis present');
  assert(!!mud, 'Mudguard Color axis present');
  assert(rim && JSON.stringify(rim.colors) === JSON.stringify(['Blue-White', 'Red-White']), 'Rim Color axis colors (got ' + JSON.stringify(rim && rim.colors) + ')');
  assert(mud && JSON.stringify(mud.colors) === JSON.stringify(['Blue', 'Red']), 'Mudguard Color axis colors (got ' + JSON.stringify(mud && mud.colors) + ')');
  assert(!axes.some(a => a.label.includes('/')), 'no composite "X / Y" combo strings anywhere (got labels ' + JSON.stringify(axes.map(a => a.label)) + ')');

  assert(res.data.primaryColorAxis === 'Rim Color', 'primaryColorAxis echoed back (got "' + res.data.primaryColorAxis + '")');
  assert(!!res.data.primaryAxisKey && res.data.primaryAxisKey === (rim && rim.key), 'primaryAxisKey resolves to the Rim Color axis (got "' + res.data.primaryAxisKey + '" vs rim.key "' + (rim && rim.key) + '")');
}

console.log('\n=== Test 2: getProcessColorGroups (legacy flat contract) still a flat union, unbroken for existing consumers ===');
{
  const res = getProcessColorGroups(processId);
  assert(res.success, 'getProcessColorGroups succeeds');
  assert(JSON.stringify(res.data) === JSON.stringify(['Blue', 'Blue-White', 'Red', 'Red-White']), 'flat union of all 4 axis colors, sorted (got ' + JSON.stringify(res.data) + ')');
}

console.log('\n=== Test 3: saveProduction sums ONLY the primary axis - the reported double-count bug ===');
{
  // Exactly the reported scenario: Rim Color "Red-White" (the real 10-unit
  // batch) checked alongside Mudguard Color "Red" (the same batch, described
  // from the mudguard's own component naming) - both entered as qty 10.
  const res = saveProduction({
    processId: processId,
    assignedTo: 'Test Contractor',
    status: 'Pending',
    colorBreakdown: JSON.stringify([
      { color: 'Red-White', qty: 10 },
      { color: 'Red', qty: 10 }
    ]),
    componentsConsumed: JSON.stringify([
      { itemName: 'Assembly Screws', sourceType: 'ITEM', qty: 40, colorGroup: 'COMMON' },
      { itemName: 'Painted Rim - Red-White', sourceType: 'ITEM', qty: 10, colorGroup: 'Red-White' },
      { itemName: 'Mudguard - Red', sourceType: 'ITEM', qty: 10, colorGroup: 'Red' }
    ])
  });
  assert(res.success, 'saveProduction succeeds: ' + res.message);
  const lotNumber = res.data && res.data.lotNumber;
  assert(!!lotNumber, 'lot number generated (got "' + lotNumber + '")');

  const allLots = getProductionData();
  const lot = (allLots.data || []).find(l => l.lotNumber === lotNumber);
  assert(!!lot, 'saved lot found by lot number');
  assert(lot.qty === 10, 'lot quantity is 10 (the primary Rim Color total), NOT 20 (got ' + (lot && lot.qty) + ')');
  assert(Array.isArray(lot.colorBreakdown) && lot.colorBreakdown.length === 2, 'both breakdown rows still recorded for consumption/history (got ' + JSON.stringify(lot && lot.colorBreakdown) + ')');
}

console.log('\n=== Test 4: a lot checking ONLY the non-primary axis is rejected, not silently zeroed ===');
{
  const res = saveProduction({
    processId: processId,
    assignedTo: 'Test Contractor',
    status: 'Pending',
    colorBreakdown: JSON.stringify([{ color: 'Red', qty: 10 }]),
    componentsConsumed: JSON.stringify([
      { itemName: 'Mudguard - Red', sourceType: 'ITEM', qty: 10, colorGroup: 'Red' }
    ])
  });
  assert(res.success === false, 'lot with only a non-primary (Mudguard Color) row is rejected');
  assert(/Rim Color/i.test(res.message), 'rejection message names the required primary axis (got "' + res.message + '")');
}

console.log('\n=== Test 5: legacy behavior (no primaryColorAxis) still sums every row - unaffected by this feature ===');
{
  const legacyRes = saveProcess({
    processName: 'Legacy Multi-Color Process',
    sequence: 2,
    lotPrefix: 'LMC',
    outputItemName: 'Legacy Output',
    isFinalStage: false,
    active: true,
    remarks: '',
    // No primaryColorAxis at all - the pre-existing behavior for any process
    // that hasn't opted into Color Axes.
    components: JSON.stringify([
      { itemName: 'Red Paint', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Red' },
      { itemName: 'Blue Paint', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Blue' }
    ])
  });
  const legacyProcessId = legacyRes.data && legacyRes.data.processId;

  const res = saveProduction({
    processId: legacyProcessId,
    assignedTo: 'Test Contractor',
    status: 'Pending',
    colorBreakdown: JSON.stringify([{ color: 'Red', qty: 6 }, { color: 'Blue', qty: 4 }]),
    componentsConsumed: JSON.stringify([
      { itemName: 'Red Paint', sourceType: 'ITEM', qty: 6, colorGroup: 'Red' },
      { itemName: 'Blue Paint', sourceType: 'ITEM', qty: 4, colorGroup: 'Blue' }
    ])
  });
  assert(res.success, 'legacy saveProduction succeeds: ' + res.message);
  const lot = (getProductionData().data || []).find(l => l.lotNumber === (res.data && res.data.lotNumber));
  assert(!!lot && lot.qty === 10, 'legacy process still sums every checked color (6+4=10) unchanged (got ' + (lot && lot.qty) + ')');
}

console.log('\n=== Test 6: operator picks Primary Axis on the Production form itself (no Process-editor setup yet) ===');
{
  // Same 2-axis recipe as the main Setup process, but saved WITHOUT a
  // primaryColorAxis - simulating a process nobody has been into the
  // Process editor to configure yet (see Script.html's
  // _buildColorAxisGroupHeader / setPrimaryColorAxisChoice - the operator
  // can pick Primary right on the Production Lot form instead).
  const res = saveProcess({
    processName: 'Unconfigured Multi-Axis Process',
    sequence: 3,
    lotPrefix: 'UMA',
    outputItemName: 'Unconfigured Output',
    isFinalStage: false,
    active: true,
    remarks: '',
    components: JSON.stringify([
      { itemName: 'Painted Rim - Blue-White', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Blue-White', colorAxis: 'Rim Color' },
      { itemName: 'Painted Rim - Red-White', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Red-White', colorAxis: 'Rim Color' },
      { itemName: 'Mudguard - Blue', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Blue', colorAxis: 'Mudguard Color' },
      { itemName: 'Mudguard - Red', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Red', colorAxis: 'Mudguard Color' }
    ])
  });
  const umaId = res.data && res.data.processId;
  assert(!!umaId, 'Unconfigured Multi-Axis Process created (got "' + umaId + '")');

  const beforeRow = (getProcessData(false).data || []).find(p => p.processId === umaId);
  assert(beforeRow && beforeRow.primaryColorAxis === '', 'process has no primaryColorAxis yet (got "' + (beforeRow && beforeRow.primaryColorAxis) + '")');

  const lot1 = saveProduction({
    processId: umaId,
    assignedTo: 'Test Contractor',
    status: 'Pending',
    primaryColorAxis: 'Rim Color', // <- picked on the Production form, not configured on the process
    colorBreakdown: JSON.stringify([
      { color: 'Red-White', qty: 8 },
      { color: 'Red', qty: 8 }
    ]),
    componentsConsumed: JSON.stringify([
      { itemName: 'Painted Rim - Red-White', sourceType: 'ITEM', qty: 8, colorGroup: 'Red-White' },
      { itemName: 'Mudguard - Red', sourceType: 'ITEM', qty: 8, colorGroup: 'Red' }
    ])
  });
  assert(lot1.success, 'lot with an in-form Primary Axis pick saves: ' + lot1.message);
  const savedLot1 = (getProductionData().data || []).find(l => l.lotNumber === (lot1.data && lot1.data.lotNumber));
  assert(!!savedLot1 && savedLot1.qty === 8, 'this lot\'s quantity is 8 (Rim Color total), not 16 (got ' + (savedLot1 && savedLot1.qty) + ')');

  const afterRow = (getProcessData(false).data || []).find(p => p.processId === umaId);
  assert(afterRow && afterRow.primaryColorAxis === 'Rim Color', 'the in-form pick was persisted onto the process for future lots (got "' + (afterRow && afterRow.primaryColorAxis) + '")');

  // A SECOND lot, saved with no explicit primaryColorAxis in formData at
  // all, must now benefit from the persisted default automatically.
  const lot2 = saveProduction({
    processId: umaId,
    assignedTo: 'Test Contractor',
    status: 'Pending',
    colorBreakdown: JSON.stringify([
      { color: 'Blue-White', qty: 5 },
      { color: 'Blue', qty: 5 }
    ]),
    componentsConsumed: JSON.stringify([
      { itemName: 'Painted Rim - Blue-White', sourceType: 'ITEM', qty: 5, colorGroup: 'Blue-White' },
      { itemName: 'Mudguard - Blue', sourceType: 'ITEM', qty: 5, colorGroup: 'Blue' }
    ])
  });
  assert(lot2.success, 'second lot (no explicit pick) saves: ' + lot2.message);
  const savedLot2 = (getProductionData().data || []).find(l => l.lotNumber === (lot2.data && lot2.data.lotNumber));
  assert(!!savedLot2 && savedLot2.qty === 5, 'second lot inherits the now-persisted Rim Color default automatically (got ' + (savedLot2 && savedLot2.qty) + ')');
}

console.log('\n=== Test 7: countsTowardTotal excludes a legacy "Other"-bucket duplicate row (no formal Primary Axis at all) ===');
{
  // Reproduces the live report: a process with NO 2+ Color Axes configured
  // (so getProcessColorAxes never activates axis mode) still has a checklist
  // "Other" bucket for colors with no pool-item signature match (see
  // Script.html's renderGroupedColorChecklist). The client now
  // segment-matches an "Other" row like "Pink" against a checked
  // pool-signature-group row like "Pink / BCP / Pink-White" and tags it
  // countsTowardTotal: false, since it's the same physical batch tagged a
  // second way, not an additional one — saveProduction must honor that flag
  // even with zero primaryColorAxis config anywhere.
  const res = saveProcess({
    processName: 'Legacy Other-Bucket Process',
    sequence: 4,
    lotPrefix: 'LOB',
    outputItemName: 'Legacy Other-Bucket Output',
    isFinalStage: false,
    active: true,
    remarks: '',
    components: JSON.stringify([
      { itemName: 'Fitted Frame - Pink Combo', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Pink / BCP / Pink-White' },
      { itemName: 'Loose Tag - Pink', sourceType: 'ITEM', qtyPerUnit: 1, colorGroup: 'Pink' }
    ])
  });
  const processId = res.data && res.data.processId;
  assert(!!processId, 'Legacy Other-Bucket Process created (got "' + processId + '")');

  const lot = saveProduction({
    processId,
    assignedTo: 'Test Contractor',
    status: 'Pending',
    // No primaryColorAxis anywhere - this process was never put through the
    // Color Axes feature. The client marks the auto-matched "Pink" row
    // countsTowardTotal:false; the pool-signature row it matched stays true.
    colorBreakdown: JSON.stringify([
      { color: 'Pink / BCP / Pink-White', qty: 12, countsTowardTotal: true },
      { color: 'Pink', qty: 12, countsTowardTotal: false }
    ]),
    componentsConsumed: JSON.stringify([
      { itemName: 'Fitted Frame - Pink Combo', sourceType: 'ITEM', qty: 12, colorGroup: 'Pink / BCP / Pink-White' },
      { itemName: 'Loose Tag - Pink', sourceType: 'ITEM', qty: 12, colorGroup: 'Pink' }
    ])
  });
  assert(lot.success, 'lot with a countsTowardTotal:false duplicate row saves: ' + lot.message);
  const savedLot = (getProductionData().data || []).find(l => l.lotNumber === (lot.data && lot.data.lotNumber));
  assert(!!savedLot && savedLot.qty === 12, 'lot quantity is 12, not 24 - the "Pink" duplicate did not double-count (got ' + (savedLot && savedLot.qty) + ')');
  assert(!!savedLot && savedLot.colorBreakdown && savedLot.colorBreakdown.length === 2,
    'both breakdown rows still recorded for consumption/history (got ' + JSON.stringify(savedLot && savedLot.colorBreakdown) + ')');

  // A payload with NO countsTowardTotal field at all (an older client, or a
  // row the grouping logic never flagged) must default to counting - the
  // exact pre-existing "sum everything" legacy behavior, unaffected.
  const backCompatLot = saveProduction({
    processId,
    assignedTo: 'Test Contractor',
    status: 'Pending',
    colorBreakdown: JSON.stringify([
      { color: 'Pink / BCP / Pink-White', qty: 5 },
      { color: 'Pink', qty: 3 }
    ]),
    componentsConsumed: JSON.stringify([
      { itemName: 'Fitted Frame - Pink Combo', sourceType: 'ITEM', qty: 5, colorGroup: 'Pink / BCP / Pink-White' },
      { itemName: 'Loose Tag - Pink', sourceType: 'ITEM', qty: 3, colorGroup: 'Pink' }
    ])
  });
  assert(backCompatLot.success, 'lot with no countsTowardTotal field at all saves: ' + backCompatLot.message);
  const savedBackCompatLot = (getProductionData().data || []).find(l => l.lotNumber === (backCompatLot.data && backCompatLot.data.lotNumber));
  assert(!!savedBackCompatLot && savedBackCompatLot.qty === 8,
    'omitted countsTowardTotal defaults to counting - legacy sum-everything unchanged (5+3=8, got ' + (savedBackCompatLot && savedBackCompatLot.qty) + ')');
}

console.log('\n' + (failures === 0 ? 'ALL TESTS PASSED' : failures + ' TEST(S) FAILED'));
process.exit(failures === 0 ? 0 : 1);
