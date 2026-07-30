/**
 * Regression: saveProcess's edit path wrote a hardcoded 10-column range while
 * handing it an 11-value row (PROCESS_COL.DISPATCH_DIFFERENTIATOR = 11), so
 * every "Save" on an existing process failed with:
 *
 *   "The number of columns in the data does not match the number of columns
 *    in the range. The data has 11 but the range has 10."
 *
 * The existing tests never caught it because their FakeRange.setValues just
 * writes cell-by-cell and ignores the range's declared width. This harness's
 * setValues/setValue enforce the range bounds exactly as Apps Script does, so
 * any future column added to PROCESS_COL without updating the write width
 * fails here instead of in production.
 *
 * Run: node .pw-test/verify_process_save_row_width.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = 'c:\\Users\\erkar\\my-app-script-project';

class FakeRange {
  constructor(s, r, c, nr, nc) { this.sheet = s; this.row = r; this.col = c; this.numRows = nr; this.numCols = nc; }
  getValues() { const o = []; for (let r = 0; r < this.numRows; r++) { const a = []; for (let c = 0; c < this.numCols; c++) a.push(this.sheet._get(this.row + r, this.col + c)); o.push(a); } return o; }
  getValue() { return this.sheet._get(this.row, this.col); }
  // Strict, matching Apps Script: the data's shape must equal the range's shape.
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
  appendRow(a) { const r = this.getLastRow() + 1; a.forEach((v, i) => this._set(r, i + 1, v)); }
  deleteRow(r) { this.rows.splice(r - 1, 1); } deleteRows(r, n) { this.rows.splice(r - 1, n); }
  insertRows(r, n) { for (let i = 0; i < n; i++) this.rows.splice(r - 1, 0, []); }
  insertColumnsAfter(a, h) { this.rows.forEach(w => { w.splice(a, 0, ...new Array(h).fill('')); }); }
}
class FakeSpreadsheet { constructor() { this.sheets = {}; } getSheetByName(n) { return this.sheets[n] || null; } addSheet(n) { const s = new FakeSheet(n); this.sheets[n] = s; return s; } insertSheet(n) { return this.addSheet(n); } }
const ss = new FakeSpreadsheet();
const sandbox = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush: () => {} },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
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
// `const` declarations inside the vm context aren't reachable as properties of
// the sandbox object, so re-expose the two the assertions below read.
vm.runInContext('global.APP_CONFIG=APP_CONFIG; global.PROCESS_COL=PROCESS_COL;', ctx, { filename: 'expose.js' });
const C = ctx;
const { PROCESS_COL } = C;

let failures = 0;
function assert(cond, msg) { if (!cond) { failures++; console.error('FAIL:', msg); } else { console.log('PASS:', msg); } }

const base = {
  processName: 'Width Check', sequence: 4, lotPrefix: 'ZWC', outputItemName: 'Width Checked Frame',
  processType: 'Painting', isFinalStage: false, active: true, remarks: 'first',
  primaryColorAxis: '', dispatchDifferentiator: '', components: JSON.stringify([])
};

console.log('=== Test 1: create ===');
let res = C.saveProcess(base);
assert(res.success, 'new process saved: ' + (res.message || ''));
const pid = res.data && res.data.processId;

console.log('\n=== Test 2: edit writes a full-width row (the reported bug) ===');
res = C.saveProcess(Object.assign({}, base, {
  processId: pid, remarks: 'edited', primaryColorAxis: 'Frame', dispatchDifferentiator: 'Width Checked Frame'
}));
assert(res.success, 'existing process re-saved without a column-count error: ' + (res.message || ''));

console.log('\n=== Test 3: every column round-trips, none shifted ===');
const rec = (C.getProcessData(false).data || []).find(p => p.processId === pid);
assert(!!rec, 'edited process still readable');
assert(rec && rec.remarks === 'edited', `remarks updated (got "${rec && rec.remarks}")`);
assert(rec && rec.processType === 'Painting', `processType intact (got "${rec && rec.processType}")`);
assert(rec && rec.primaryColorAxis === 'Frame', `primaryColorAxis intact (got "${rec && rec.primaryColorAxis}")`);
assert(rec && rec.dispatchDifferentiator === 'Width Checked Frame',
  `dispatchDifferentiator intact (got "${rec && rec.dispatchDifferentiator}")`);

console.log('\n=== Test 4: the written row is exactly PROCESS_COL wide ===');
const procSheet = ss.getSheetByName(C.APP_CONFIG.SHEETS.PROCESS_MASTER);
const width = Math.max(...Object.keys(PROCESS_COL).map(k => PROCESS_COL[k]));
assert(width === PROCESS_COL.DISPATCH_DIFFERENTIATOR,
  `DISPATCH_DIFFERENTIATOR is the last column (width ${width})`);
assert(procSheet.getLastColumn() === width,
  `sheet holds all ${width} columns after the edit (got ${procSheet.getLastColumn()})`);

console.log(`\n${failures === 0 ? 'ALL PASS' : failures + ' FAILURE(S)'}`);
process.exit(failures === 0 ? 0 : 1);
