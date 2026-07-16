/**
 * Standalone Node harness that mocks the Apps Script runtime (SpreadsheetApp,
 * LockService, etc.) well enough to load and execute the REAL server-side
 * files (config.js, utils.js, module_items.js, module_stock.js,
 * module_bill.js, module_po.js, module_bom.js) and run the merge-on-duplicate
 * + identity-backfill feature against in-memory fake sheets.
 *
 * Run: node .pw-test/test_merge_and_backfill.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

// ─────────────────────────────────────────────────────────────────────────
// Minimal in-memory Sheet/Range/Spreadsheet mocks
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
    // last non-fully-empty row
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
    const blank = [];
    for (let i = 0; i < n; i++) this.rows.splice(r - 1, 0, []);
  }
}

class FakeSpreadsheet {
  constructor() {
    this.sheets = {};
  }
  getSheetByName(name) {
    return this.sheets[name] || null;
  }
  addSheet(name, headerRows = 1) {
    const s = new FakeSheet(name);
    this.sheets[name] = s;
    return s;
  }
}

// ─────────────────────────────────────────────────────────────────────────
// Build sandbox, load real files into it
// ─────────────────────────────────────────────────────────────────────────

const ss = new FakeSpreadsheet();
const logs = [];

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
  Logger: { log: (...args) => { logs.push(args); } },
  Utilities: { getUuid: () => 'uuid-' + Math.random().toString(36).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => 'test@example.com' }) },
  recalculateStock: undefined
};
sandbox.global = sandbox;

const ctx = vm.createContext(sandbox);

const files = [
  'config.js',
  'utils.js',
  'module_items.js',
  'module_stock.js',
  'module_bill.js',
  'module_po.js',
  'module_bom.js',
  'module_process.js',
  'module_return.js',
  'module_wastage.js',
  'module_issue.js',
  'module_production.js'
];

files.forEach(f => {
  const code = fs.readFileSync(path.join(ROOT, f), 'utf8');
  vm.runInContext(code, ctx, { filename: f });
});

// `const`/`let` top-level declarations in vm scripts live in a separate
// lexical environment from the sandbox object, so they don't show up as
// `ctx.X` automatically (unlike `var`/function declarations). Re-expose the
// bindings we need onto `global` (which we pointed at the sandbox itself).
vm.runInContext(`
  global.APP_CONFIG = APP_CONFIG;
  global.ITEMS_COL = ITEMS_COL;
  global.STOCK_COL = STOCK_COL;
  global.BILL_COL = BILL_COL;
  global.PO_COL = PO_COL;
  global.BOM_COL = BOM_COL;
  global.RETURN_COL = RETURN_COL;
  global.WASTAGE_COL = WASTAGE_COL;
  global.ISSUE_COL = ISSUE_COL;
  global.PRODUCTION_COL = PRODUCTION_COL;
  global.COMPONENT_SOURCE_TYPES = COMPONENT_SOURCE_TYPES;
`, ctx, { filename: 'expose.js' });

const {
  APP_CONFIG, ITEMS_COL, STOCK_COL, BILL_COL, PO_COL, BOM_COL,
  RETURN_COL, WASTAGE_COL, ISSUE_COL, PRODUCTION_COL, COMPONENT_SOURCE_TYPES,
  saveItem, mergeItemEdit, syncStockForItem,
  backfillBillItemRefs, backfillPOItemRefs, backfillBOMItemRefs,
  backfillReturnItemRefs, backfillWastageItemRefs, backfillIssueItemRefs,
  backfillProductionConsumedItemRefs, getItemIdentityDriftReport,
  fixItemIdentityDriftReference
} = ctx;

// ─────────────────────────────────────────────────────────────────────────
// Test setup: seed sheets
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

function seedItemsSheet() {
  const sheet = ss.addSheet(APP_CONFIG.SHEETS.ITEMS);
  // header
  sheet._set(1, 1, 'Item Name'); sheet._set(1, 2, 'Size'); sheet._set(1, 3, 'Remarks');
  sheet._set(1, 4, 'Narration'); sheet._set(1, 5, 'Spec');
  // row 2: 24 inch item (source)
  sheet._set(2, ITEMS_COL.ITEM_NAME, 'BACK-REST-D-26---BLACK');
  sheet._set(2, ITEMS_COL.SIZE, '24 inch');
  sheet._set(2, ITEMS_COL.REMARKS, '');
  sheet._set(2, ITEMS_COL.NARRATION, '');
  sheet._set(2, ITEMS_COL.SPECIFICATION, '');
  sheet._set(2, ITEMS_COL.VENDOR_DATA, 'PJ Cycle Industries');
  sheet._set(2, ITEMS_COL.VENDOR_DATA + 1, 53);
  // row 3: GENERAL item (target)
  sheet._set(3, ITEMS_COL.ITEM_NAME, 'BACK-REST-D-26---BLACK');
  sheet._set(3, ITEMS_COL.SIZE, 'GENERAL');
  sheet._set(3, ITEMS_COL.REMARKS, '');
  sheet._set(3, ITEMS_COL.NARRATION, '');
  sheet._set(3, ITEMS_COL.SPECIFICATION, '');
  sheet._set(3, ITEMS_COL.VENDOR_DATA, 'Prayag Industries');
  sheet._set(3, ITEMS_COL.VENDOR_DATA + 1, 51);
  sheet._set(3, ITEMS_COL.VENDOR_DATA + 2, 'PJ Cycle Industries');
  sheet._set(3, ITEMS_COL.VENDOR_DATA + 3, 53);
  return sheet;
}

function seedStockSheet() {
  const sheet = ss.addSheet(APP_CONFIG.SHEETS.STOCK);
  sheet._set(1, 1, 'Item Name'); sheet._set(1, 2, 'Size'); sheet._set(1, 3, 'Initial');
  sheet._set(1, 4, 'Current'); sheet._set(1, 5, 'Threshold');
  // source: 24 inch, stock 10
  sheet._set(2, STOCK_COL.ITEM_NAME, 'BACK-REST-D-26---BLACK');
  sheet._set(2, STOCK_COL.SIZE, '24 inch');
  sheet._set(2, STOCK_COL.INITIAL_STOCK, 10);
  sheet._set(2, STOCK_COL.CURRENT_STOCK, 10);
  sheet._set(2, STOCK_COL.THRESHOLD, 0);
  // target: GENERAL, stock 632
  sheet._set(3, STOCK_COL.ITEM_NAME, 'BACK-REST-D-26---BLACK');
  sheet._set(3, STOCK_COL.SIZE, 'GENERAL');
  sheet._set(3, STOCK_COL.INITIAL_STOCK, 632);
  sheet._set(3, STOCK_COL.CURRENT_STOCK, 632);
  sheet._set(3, STOCK_COL.THRESHOLD, 0);
  return sheet;
}

function seedBillSheet() {
  const sheet = ss.addSheet(APP_CONFIG.SHEETS.BILL);
  // row 2: a bill line item referencing the 24-inch item
  sheet._set(2, BILL_COL.PO_NUMBER, 'DIRECT');
  sheet._set(2, BILL_COL.BILL_NUMBER, 'B-1');
  sheet._set(2, BILL_COL.BILL_DATE, '01/01/2026');
  sheet._set(2, BILL_COL.VENDOR, 'PJ Cycle Industries');
  sheet._set(2, BILL_COL.CONTACT, '');
  sheet._set(2, BILL_COL.ITEM_NAME, 'BACK-REST-D-26---BLACK');
  sheet._set(2, BILL_COL.SIZE, '24 inch');
  sheet._set(2, BILL_COL.NARRATION, '');
  sheet._set(2, BILL_COL.QTY, 5);
  sheet._set(2, BILL_COL.UNIT, 'Pcs');
  sheet._set(2, BILL_COL.PRICE, 53);
  sheet._set(2, BILL_COL.GST, 18);
  sheet._set(2, BILL_COL.TOTAL, 312.7);
  sheet._set(2, BILL_COL.REMARKS, '');
  return sheet;
}

function seedPOSheet() {
  const sheet = ss.addSheet(APP_CONFIG.SHEETS.PO);
  // rows 2 = header2 per PO_SETTINGS.DATA_START_ROW=3, put data at row 3
  sheet._set(3, PO_COL.PO_NUMBER, 'PO-1');
  sheet._set(3, PO_COL.PO_DATE, '01/01/2026');
  sheet._set(3, PO_COL.VENDOR, 'PJ Cycle Industries');
  sheet._set(3, PO_COL.CONTACT, '');
  sheet._set(3, PO_COL.ITEM_NAME, 'BACK-REST-D-26---BLACK');
  sheet._set(3, PO_COL.NARRATION, 'keep-me');
  sheet._set(3, PO_COL.SIZE, '24 inch');
  sheet._set(3, PO_COL.QTY, 5);
  sheet._set(3, PO_COL.UNIT, 'Pcs');
  return sheet;
}

function seedBomSheet() {
  const sheet = ss.addSheet(APP_CONFIG.SHEETS.BOM);
  sheet._set(2, BOM_COL.PRODUCT_ID, 'P-1');
  sheet._set(2, BOM_COL.PRODUCT_NAME, 'Bike A');
  sheet._set(2, BOM_COL.ITEM_NAME, 'BACK-REST-D-26---BLACK');
  sheet._set(2, BOM_COL.SIZE, '24 inch');
  sheet._set(2, BOM_COL.NARRATION, '');
  sheet._set(2, BOM_COL.RATE, 53);
  sheet._set(2, BOM_COL.VENDOR, 'PJ Cycle Industries');
  sheet._set(2, BOM_COL.QTY_PER_PRODUCT, 1);
  return sheet;
}

function seedReturnSheet() {
  const sheet = ss.addSheet(APP_CONFIG.SHEETS.RETURN);
  sheet._set(2, RETURN_COL.RETURN_NUMBER, 'RET-1');
  sheet._set(2, RETURN_COL.RETURN_DATE, '01/01/2026');
  sheet._set(2, RETURN_COL.VENDOR, 'PJ Cycle Industries');
  sheet._set(2, RETURN_COL.ITEM_NAME, 'BACK-REST-D-26---BLACK');
  sheet._set(2, RETURN_COL.SIZE, '24 inch');
  sheet._set(2, RETURN_COL.QTY, 1);
  sheet._set(2, RETURN_COL.BASE_QTY, 1);
  return sheet;
}

function seedWastageSheet() {
  const sheet = ss.addSheet(APP_CONFIG.SHEETS.WASTAGE);
  sheet._set(2, WASTAGE_COL.WASTAGE_ID, 'WST-1');
  sheet._set(2, WASTAGE_COL.DATE, '01/01/2026');
  sheet._set(2, WASTAGE_COL.ITEM_NAME, 'BACK-REST-D-26---BLACK');
  sheet._set(2, WASTAGE_COL.SIZE, '24 inch');
  sheet._set(2, WASTAGE_COL.QTY, 1);
  sheet._set(2, WASTAGE_COL.BASE_QTY, 1);
  return sheet;
}

function seedIssueSheet() {
  const sheet = ss.addSheet(APP_CONFIG.SHEETS.ISSUE);
  sheet._set(2, ISSUE_COL.ISSUE_ID, 'ISS-1');
  sheet._set(2, ISSUE_COL.DATE, '01/01/2026');
  sheet._set(2, ISSUE_COL.ISSUED_TO, 'Contractor A');
  sheet._set(2, ISSUE_COL.ITEM_NAME, 'BACK-REST-D-26---BLACK');
  sheet._set(2, ISSUE_COL.SIZE, '24 inch');
  sheet._set(2, ISSUE_COL.QTY, 1);
  sheet._set(2, ISSUE_COL.BASE_QTY, 1);
  return sheet;
}

function seedProductionSheet() {
  const sheet = ss.addSheet(APP_CONFIG.SHEETS.PRODUCTION);
  sheet._set(2, PRODUCTION_COL.DATE, '01/01/2026');
  sheet._set(2, PRODUCTION_COL.QTY, 10);
  sheet._set(2, PRODUCTION_COL.STATUS, 'Completed');
  sheet._set(2, PRODUCTION_COL.PROCESS_ID, 'PRC-1');
  sheet._set(2, PRODUCTION_COL.LOT_NUMBER, 'LOT-1');
  sheet._set(2, PRODUCTION_COL.COMPONENTS_CONSUMED, JSON.stringify([
    { itemName: 'BACK-REST-D-26---BLACK', size: '24 inch', sourceType: 'ITEM', qty: 10 },
    { itemName: 'Painted Frame', size: '', sourceType: 'POOL', qty: 10 }
  ]));
  sheet._set(2, PRODUCTION_COL.CUSTOM_COMPONENTS, JSON.stringify([
    { itemName: 'BACK-REST-D-26---BLACK', size: '24 inch', narration: '', color: '', requiredQty: 10 }
  ]));
  return sheet;
}

// ─────────────────────────────────────────────────────────────────────────
// Test 1: rename backfill (no merge)
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 1: rename backfill (no collision) ===');
{
  const itemsSheet = seedItemsSheet();
  seedStockSheet();
  const billSheet = seedBillSheet();
  const poSheet = seedPOSheet();
  const bomSheet = seedBomSheet();
  const returnSheet = seedReturnSheet();
  const wastageSheet = seedWastageSheet();
  const issueSheet = seedIssueSheet();
  const productionSheet = seedProductionSheet();

  // Remove the GENERAL row so the rename has no collision target.
  itemsSheet.deleteRow(3);

  const formData = {
    itemName: 'BACK-REST-D-26---BLACK',
    itemSize: '26 inch', // renamed, no collision
    itemRemarks: '',
    itemNarration: '',
    itemSpec: '',
    vendors: JSON.stringify([{ vendor: 'PJ Cycle Industries', rate: 53 }]),
    originalName: 'BACK-REST-D-26---BLACK',
    originalSize: '24 inch',
    originalNarration: ''
  };

  const res = saveItem(formData);
  assert(res.success, 'rename saveItem succeeds: ' + res.message);

  const stockSheet = ss.getSheetByName(APP_CONFIG.SHEETS.STOCK);
  const stockRow2 = stockSheet.getRange(2, 1, 1, 5).getValues()[0];
  assert(stockRow2[1] === '26 inch', 'Stock row renamed to 26 inch (got "' + stockRow2[1] + '")');
  assert(stockRow2[3] === 10, 'Stock quantity preserved at 10 (got ' + stockRow2[3] + ')');

  const billRow = billSheet.getRange(2, BILL_COL.ITEM_NAME, 1, 2).getValues()[0];
  assert(billRow[1] === '26 inch', 'Bill Ledger backfilled to 26 inch (got "' + billRow[1] + '")');

  const poRow = poSheet.getRange(3, PO_COL.ITEM_NAME, 1, 3).getValues()[0];
  assert(poRow[2] === '26 inch', 'PO Tracker backfilled to 26 inch (got "' + poRow[2] + '")');
  assert(poRow[1] === 'keep-me', 'PO Tracker narration preserved untouched (got "' + poRow[1] + '")');

  const bomRow = bomSheet.getRange(2, BOM_COL.ITEM_NAME, 1, 2).getValues()[0];
  assert(bomRow[1] === '26 inch', 'BOM backfilled to 26 inch (got "' + bomRow[1] + '")');

  const returnRow = returnSheet.getRange(2, RETURN_COL.ITEM_NAME, 1, 2).getValues()[0];
  assert(returnRow[1] === '26 inch', 'Return Ledger backfilled to 26 inch (got "' + returnRow[1] + '")');

  const wastageRow = wastageSheet.getRange(2, WASTAGE_COL.ITEM_NAME, 1, 2).getValues()[0];
  assert(wastageRow[1] === '26 inch', 'Wastage Log backfilled to 26 inch (got "' + wastageRow[1] + '")');

  const issueRow = issueSheet.getRange(2, ISSUE_COL.ITEM_NAME, 1, 2).getValues()[0];
  assert(issueRow[1] === '26 inch', 'Issued Stock Log backfilled to 26 inch (got "' + issueRow[1] + '")');

  const consumed = JSON.parse(productionSheet.getRange(2, PRODUCTION_COL.COMPONENTS_CONSUMED, 1, 1).getValue());
  assert(consumed[0].itemName === 'BACK-REST-D-26---BLACK' && consumed[0].size === '26 inch',
    'Production Components Consumed (ITEM-sourced) backfilled to 26 inch (got ' + JSON.stringify(consumed[0]) + ')');
  assert(consumed[1].itemName === 'Painted Frame' && consumed[1].sourceType === 'POOL',
    'Production Components Consumed POOL-sourced entry left untouched (got ' + JSON.stringify(consumed[1]) + ')');

  const custom = JSON.parse(productionSheet.getRange(2, PRODUCTION_COL.CUSTOM_COMPONENTS, 1, 1).getValue());
  assert(custom[0].size === '26 inch', 'Production legacy Custom Components backfilled to 26 inch (got ' + JSON.stringify(custom[0]) + ')');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 2: duplicate-on-edit returns mergeable info (no merge yet)
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 2: duplicate edit returns mergeable info ===');
{
  ss.sheets = {};
  seedItemsSheet();
  seedStockSheet();
  seedBillSheet();
  seedPOSheet();
  seedBomSheet();

  const formData = {
    itemName: 'BACK-REST-D-26---BLACK',
    itemSize: 'GENERAL', // collides with row 3
    itemRemarks: '',
    itemNarration: '',
    itemSpec: '',
    vendors: JSON.stringify([{ vendor: 'PJ Cycle Industries', rate: 53 }]),
    originalName: 'BACK-REST-D-26---BLACK',
    originalSize: '24 inch',
    originalNarration: ''
  };

  const res = saveItem(formData);
  assert(res.success === false, 'duplicate edit fails as expected');
  assert(res.data && res.data.mergeable === true, 'response.data.mergeable === true (data=' + JSON.stringify(res.data) + ')');
  assert(res.data && res.data.targetStock === 632, 'targetStock reported as 632 (got ' + (res.data && res.data.targetStock) + ')');
  assert(res.data && res.data.targetVendorCount === 2, 'targetVendorCount reported as 2 (got ' + (res.data && res.data.targetVendorCount) + ')');

  // Confirm nothing was written — no-op on failure.
  const itemsSheet = ss.getSheetByName(APP_CONFIG.SHEETS.ITEMS);
  assert(itemsSheet.getRange(2, ITEMS_COL.SIZE, 1, 1).getValue() === '24 inch', 'source row untouched after blocked duplicate save');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 3: full merge flow
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 3: mergeItemEdit full flow ===');
{
  ss.sheets = {};
  seedItemsSheet();
  seedStockSheet();
  const billSheet = seedBillSheet();
  const poSheet = seedPOSheet();
  const bomSheet = seedBomSheet();
  const returnSheet = seedReturnSheet();
  const wastageSheet = seedWastageSheet();
  const issueSheet = seedIssueSheet();
  const productionSheet = seedProductionSheet();

  const formData = {
    itemName: 'BACK-REST-D-26---BLACK',
    itemSize: 'GENERAL',
    itemRemarks: '',
    itemNarration: '',
    itemSpec: '',
    // Source's edited vendor list: PJ Cycle Industries at a DIFFERENT rate (collision),
    // plus a brand new vendor not on the target.
    vendors: JSON.stringify([
      { vendor: 'PJ Cycle Industries', rate: 60 },
      { vendor: 'New Vendor Co', rate: 40 }
    ]),
    originalName: 'BACK-REST-D-26---BLACK',
    originalSize: '24 inch',
    originalNarration: ''
  };

  const res = mergeItemEdit(formData);
  assert(res.success, 'mergeItemEdit succeeds: ' + res.message);

  const itemsSheet = ss.getSheetByName(APP_CONFIG.SHEETS.ITEMS);
  assert(itemsSheet.getLastRow() === 2, 'source row deleted, only target row remains (lastRow=' + itemsSheet.getLastRow() + ')');

  const targetRowVals = itemsSheet.getRange(2, 1, 1, itemsSheet.getLastColumn()).getValues()[0];
  assert(targetRowVals[ITEMS_COL.SIZE - 1] === 'GENERAL', 'target row still says GENERAL');

  // Vendor union check: Prayag (51, target-only), PJ Cycle (53, target wins over source's 60), New Vendor Co (40, source-only)
  const vendorPairs = [];
  for (let c = ITEMS_COL.VENDOR_DATA - 1; c < targetRowVals.length; c += 2) {
    if (targetRowVals[c]) vendorPairs.push([targetRowVals[c], targetRowVals[c + 1]]);
  }
  const vendorMap = Object.fromEntries(vendorPairs);
  assert(vendorMap['Prayag Industries'] === 51, 'Prayag Industries kept at 51 (got ' + vendorMap['Prayag Industries'] + ')');
  assert(vendorMap['PJ Cycle Industries'] === 53, 'PJ Cycle Industries target rate (53) wins over source rate (60) (got ' + vendorMap['PJ Cycle Industries'] + ')');
  assert(vendorMap['New Vendor Co'] === 40, 'New Vendor Co (source-only) appended at 40 (got ' + vendorMap['New Vendor Co'] + ')');
  assert(Object.keys(vendorMap).length === 3, 'exactly 3 vendors after union (got ' + Object.keys(vendorMap).length + ')');

  const stockSheet = ss.getSheetByName(APP_CONFIG.SHEETS.STOCK);
  assert(stockSheet.getLastRow() === 2, 'source Stock row deleted, only target remains (lastRow=' + stockSheet.getLastRow() + ')');
  const stockVals = stockSheet.getRange(2, 1, 1, 5).getValues()[0];
  assert(stockVals[1] === 'GENERAL', 'remaining Stock row is GENERAL');
  assert(stockVals[3] === 642, 'Stock combined to 10 + 632 = 642 (got ' + stockVals[3] + ')');

  const billRow = billSheet.getRange(2, BILL_COL.ITEM_NAME, 1, 2).getValues()[0];
  assert(billRow[1] === 'GENERAL', 'Bill Ledger backfilled from 24 inch to GENERAL (got "' + billRow[1] + '")');

  const poRow = poSheet.getRange(3, PO_COL.ITEM_NAME, 1, 3).getValues()[0];
  assert(poRow[2] === 'GENERAL', 'PO Tracker backfilled to GENERAL (got "' + poRow[2] + '")');
  assert(poRow[1] === 'keep-me', 'PO Tracker narration still preserved after merge backfill');

  const bomRow = bomSheet.getRange(2, BOM_COL.ITEM_NAME, 1, 2).getValues()[0];
  assert(bomRow[1] === 'GENERAL', 'BOM backfilled to GENERAL (got "' + bomRow[1] + '")');

  const returnRow = returnSheet.getRange(2, RETURN_COL.ITEM_NAME, 1, 2).getValues()[0];
  assert(returnRow[1] === 'GENERAL', 'Return Ledger backfilled to GENERAL (got "' + returnRow[1] + '")');

  const wastageRow = wastageSheet.getRange(2, WASTAGE_COL.ITEM_NAME, 1, 2).getValues()[0];
  assert(wastageRow[1] === 'GENERAL', 'Wastage Log backfilled to GENERAL (got "' + wastageRow[1] + '")');

  const issueRow = issueSheet.getRange(2, ISSUE_COL.ITEM_NAME, 1, 2).getValues()[0];
  assert(issueRow[1] === 'GENERAL', 'Issued Stock Log backfilled to GENERAL (got "' + issueRow[1] + '")');

  const consumed = JSON.parse(productionSheet.getRange(2, PRODUCTION_COL.COMPONENTS_CONSUMED, 1, 1).getValue());
  assert(consumed[0].size === 'GENERAL', 'Production Components Consumed backfilled to GENERAL on merge (got ' + JSON.stringify(consumed[0]) + ')');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 4: new-item creation collision stays a hard block (no merge offered)
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 4: new item collision is still a hard block ===');
{
  ss.sheets = {};
  seedItemsSheet();
  seedStockSheet();

  const formData = {
    itemName: 'BACK-REST-D-26---BLACK',
    itemSize: 'GENERAL', // collides with existing row, but this is a NEW item (no originalName)
    itemRemarks: '',
    itemNarration: '',
    itemSpec: '',
    vendors: JSON.stringify([]),
    originalName: '',
    originalSize: '',
    originalNarration: ''
  };

  const res = saveItem(formData);
  assert(res.success === false, 'new-item collision fails as expected');
  assert(res.data === null, 'new-item collision does NOT offer mergeable (data=' + JSON.stringify(res.data) + ')');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 5: re-running merge after it already happened fails cleanly
// (simulates a second concurrent request landing after the first completed)
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 5: repeat merge request fails cleanly, no double-merge ===');
{
  ss.sheets = {};
  seedItemsSheet();
  seedStockSheet();

  const formData = {
    itemName: 'BACK-REST-D-26---BLACK',
    itemSize: 'GENERAL',
    itemRemarks: '',
    itemNarration: '',
    itemSpec: '',
    vendors: JSON.stringify([{ vendor: 'PJ Cycle Industries', rate: 53 }]),
    originalName: 'BACK-REST-D-26---BLACK',
    originalSize: '24 inch',
    originalNarration: ''
  };

  const first = mergeItemEdit(formData);
  assert(first.success, 'first merge succeeds: ' + first.message);

  const second = mergeItemEdit(formData);
  assert(second.success === false, 'second (duplicate) merge request fails cleanly');
  assert(/not found/i.test(second.message), 'failure message explains source is gone (got "' + second.message + '")');

  const stockSheet = ss.getSheetByName(APP_CONFIG.SHEETS.STOCK);
  const stockVals = stockSheet.getRange(2, 1, 1, 5).getValues()[0];
  assert(stockVals[3] === 642, 'stock was combined exactly once, not twice (got ' + stockVals[3] + ')');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 6: getItemIdentityDriftReport() diagnostic
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 6: getItemIdentityDriftReport diagnostic ===');
{
  ss.sheets = {};
  seedItemsSheet();
  seedStockSheet();
  seedBillSheet();
  seedPOSheet();
  seedBomSheet();
  seedReturnSheet();
  seedWastageSheet();
  seedIssueSheet();
  seedProductionSheet();

  const clean = getItemIdentityDriftReport();
  assert(clean.success, 'clean report call succeeds: ' + clean.message);
  assert(clean.data.length === 0, 'no drift found when every reference matches a current Items Master row (got ' + JSON.stringify(clean.data) + ')');

  // Simulate leftover damage from before this fix existed: a Wastage row
  // and a Production Components-Consumed entry still naming a size Items
  // Master no longer has (as if a rename backfill had been skipped).
  const wastageSheet = ss.getSheetByName(APP_CONFIG.SHEETS.WASTAGE);
  wastageSheet._set(2, WASTAGE_COL.SIZE, '99 inch (renamed away)');

  const prodSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PRODUCTION);
  const staleConsumed = JSON.parse(prodSheet.getRange(2, PRODUCTION_COL.COMPONENTS_CONSUMED, 1, 1).getValue());
  staleConsumed[0].size = '99 inch (renamed away)';
  prodSheet.getRange(2, PRODUCTION_COL.COMPONENTS_CONSUMED, 1, 1).setValue(JSON.stringify(staleConsumed));

  const dirty = getItemIdentityDriftReport();
  assert(dirty.success, 'dirty report call succeeds: ' + dirty.message);
  assert(dirty.data.length === 2, 'exactly 2 stale references found (got ' + JSON.stringify(dirty.data) + ')');

  const bySheet = Object.fromEntries(dirty.data.map(f => [f.sheet, f]));
  assert(!!bySheet['Wastage Log'], 'stale Wastage Log reference flagged');
  assert(!!bySheet['Production (Components Consumed)'], 'stale Production Components-Consumed reference flagged');
  assert(bySheet['Wastage Log'].size === '99 inch (renamed away)', 'flagged finding reports the actual stale size (got "' + bySheet['Wastage Log'].size + '")');

  // The POOL-sourced entry in Production's snapshot must never be flagged —
  // its itemName is an upstream process's Output Item Name, not Items Master.
  assert(!dirty.data.some(f => f.itemName === 'Painted Frame'), 'POOL-sourced Production entry is never flagged as drift');
}

// ─────────────────────────────────────────────────────────────────────────
// Test 7: fixItemIdentityDriftReference() repairs a stale reference
// ─────────────────────────────────────────────────────────────────────────
console.log('\n=== Test 7: fixItemIdentityDriftReference repairs stale references ===');
{
  ss.sheets = {};
  seedItemsSheet();
  seedStockSheet();
  seedBillSheet();
  seedPOSheet();
  seedBomSheet();
  seedWastageSheet();
  seedIssueSheet();
  const productionSheet = seedProductionSheet();

  // Simulate leftover pre-fix damage: Wastage and Production's Components
  // Consumed still name a size ("99 inch") Items Master no longer has.
  const wastageSheet = ss.getSheetByName(APP_CONFIG.SHEETS.WASTAGE);
  wastageSheet._set(2, WASTAGE_COL.SIZE, '99 inch');
  const staleConsumed = JSON.parse(productionSheet.getRange(2, PRODUCTION_COL.COMPONENTS_CONSUMED, 1, 1).getValue());
  staleConsumed[0].size = '99 inch';
  productionSheet.getRange(2, PRODUCTION_COL.COMPONENTS_CONSUMED, 1, 1).setValue(JSON.stringify(staleConsumed));

  // Reject: target doesn't exist in Items Master.
  const badFix = fixItemIdentityDriftReference('BACK-REST-D-26---BLACK', '99 inch', 'BACK-REST-D-26---BLACK', 'NOT-A-REAL-SIZE');
  assert(badFix.success === false, 'fix rejects a target that is not in Items Master');

  // Real fix: repoint to the existing GENERAL row.
  const goodFix = fixItemIdentityDriftReference('BACK-REST-D-26---BLACK', '99 inch', 'BACK-REST-D-26---BLACK', 'GENERAL');
  assert(goodFix.success, 'fix succeeds when target exists: ' + goodFix.message);

  const fixedWastageRow = wastageSheet.getRange(2, WASTAGE_COL.ITEM_NAME, 1, 2).getValues()[0];
  assert(fixedWastageRow[1] === 'GENERAL', 'Wastage Log repointed to GENERAL (got "' + fixedWastageRow[1] + '")');

  const fixedConsumed = JSON.parse(productionSheet.getRange(2, PRODUCTION_COL.COMPONENTS_CONSUMED, 1, 1).getValue());
  assert(fixedConsumed[0].size === 'GENERAL', 'Production Components Consumed repointed to GENERAL (got ' + JSON.stringify(fixedConsumed[0]) + ')');

  // Confirm the drift report is now clean.
  const afterFix = getItemIdentityDriftReport();
  assert(afterFix.data.length === 0, 'no drift remains after fixItemIdentityDriftReference (got ' + JSON.stringify(afterFix.data) + ')');
}

// ─────────────────────────────────────────────────────────────────────────
console.log('\n' + (failures === 0 ? 'ALL TESTS PASSED' : failures + ' TEST(S) FAILED'));
process.exit(failures === 0 ? 0 : 1);
