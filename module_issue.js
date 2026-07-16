/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_issue.gs — ISSUED STOCK LOG MODULE
 *
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Records ad-hoc issuance of Stock items — components a contractor needs
 * beyond what a Process's own recipe (BOM) calls for. Deliberately kept
 * separate from Production's Components Consumed list: issuing an item here
 * never touches a lot's BOM/costing, it only debits Stock directly (same
 * mechanism as Wastage — see module_wastage.js).
 * Reference (e.g. a Production Lot #) is optional and purely informational.
 * Each entry can have multiple item rows.
 *
 * Dependencies (shared globals, same Apps Script project namespace):
 * ───────────────────────────────────────────────────────────────────────────
 * - config.js (APP_CONFIG, ISSUE_COL)
 * - utils.js (getSheet, buildResponse, sanitizeString, toSafeDateString)
 * - module_bill.js (_toValidNumber — shared numeric validator)
 * - module_units.js / module_items.js (unit conversion helpers)
 * ═══════════════════════════════════════════════════════════════════════════
 */

const ISSUE_LOCK_TIMEOUT_MS = 15000;

// ─────────────────────────────────────────────────────────────────────────
// SHEET INITIALIZATION / SELF-HEALING
// ─────────────────────────────────────────────────────────────────────────

function initIssueSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.ISSUE);
    if (!sheet) {
      sheet = ss.insertSheet(APP_CONFIG.SHEETS.ISSUE);
    }

    const headers = [
      'Issue ID', 'Date', 'Issued To', 'Reference', 'Item Name', 'Size',
      'Qty', 'Unit', 'Remarks', 'Base Qty', 'Rate', 'Vendor'
    ];

    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Issued Stock Log sheet initialized successfully.');
  } catch (error) {
    Log.error('[initIssueSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize Issued Stock Log sheet: ' + error.message);
  }
}

function _getIssueSheet() {
  try {
    return getSheet(APP_CONFIG.SHEETS.ISSUE);
  } catch (e) {
    initIssueSheet();
    return getSheet(APP_CONFIG.SHEETS.ISSUE);
  }
}

/**
 * Self-healing migration: ensures the Issued Stock Log sheet has the
 * Rate / Vendor headers, writing them if an existing sheet predates
 * these optional columns. Mirrors _ensureRateHistoryColumns (module_vendors.js).
 * @param {Sheet} sheet - The Issued Stock Log sheet
 * @private
 */
function _ensureIssueColumns(sheet) {
  const cache = CacheService.getScriptCache();
  if (cache.get('issue_columns_ok') === '1') return;

  const lastCol = sheet.getLastColumn();
  const headers = lastCol >= 1 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const hasRate = headers.some(h => String(h).trim().toLowerCase() === 'rate');
  const hasVendor = headers.some(h => String(h).trim().toLowerCase() === 'vendor');

  if (!hasRate) sheet.getRange(1, ISSUE_COL.RATE).setValue('Rate');
  if (!hasVendor) sheet.getRange(1, ISSUE_COL.VENDOR).setValue('Vendor');
  if (!hasRate || !hasVendor) SpreadsheetApp.flush();

  cache.put('issue_columns_ok', '1', 3600);
}

// ─────────────────────────────────────────────────────────────────────────
// ID GENERATION
// ─────────────────────────────────────────────────────────────────────────

function _generateIssueId() {
  const now = new Date();
  const datePart = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
  const timePart = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HHmmss');
  return 'ISS-' + datePart + '-' + timePart;
}

// ─────────────────────────────────────────────────────────────────────────
// ISSUE DATA RETRIEVAL
// ─────────────────────────────────────────────────────────────────────────

/**
 * getIssueData()
 *
 * Retrieves all issued-stock records, grouped by Issue ID.
 * Returns newest first.
 *
 * @returns {Object} API response — data: Array of issue groups:
 *   { issueId, date, dateRaw, issuedTo, reference, vendor, remarks,
 *     items: [{name, size, qty, unit, baseQty, rate, value}], totalQty, totalValue }
 */
function getIssueData() {
  try {
    const sheet = _getIssueSheet();
    _ensureIssueColumns(sheet);

    const startRow = APP_CONFIG.ISSUE_SETTINGS.DATA_START_ROW;
    const lastRow = sheet.getLastRow();

    if (lastRow < startRow) {
      return buildResponse(true, []);
    }

    const numCols = ISSUE_COL.VENDOR;
    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, numCols).getValues();

    const issueMap = {};

    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const issueId = String(r[ISSUE_COL.ISSUE_ID - 1] || '').trim();
      if (!issueId) continue;

      if (!issueMap[issueId]) {
        const rawDate = r[ISSUE_COL.DATE - 1];
        const parsedDate = toSafeDateString(rawDate);

        const isoDateStr = parsedDate
          ? parsedDate.split('/').reverse().join('-')
          : '';

        issueMap[issueId] = {
          issueId: issueId,
          date: parsedDate || (rawDate ? String(rawDate) : 'N/A'),
          dateRaw: isoDateStr,
          issuedTo: String(r[ISSUE_COL.ISSUED_TO - 1] || ''),
          reference: String(r[ISSUE_COL.REFERENCE - 1] || ''),
          vendor: String(r[ISSUE_COL.VENDOR - 1] || ''),
          remarks: String(r[ISSUE_COL.REMARKS - 1] || ''),
          items: [],
          totalQty: 0,
          totalValue: 0
        };
      }

      const qty = _toValidNumber(r[ISSUE_COL.QTY - 1], 'Qty', false);
      const rate = _toValidNumber(r[ISSUE_COL.RATE - 1] || 0, 'Rate', true);

      issueMap[issueId].items.push({
        name: String(r[ISSUE_COL.ITEM_NAME - 1] || ''),
        size: String(r[ISSUE_COL.SIZE - 1] || ''),
        qty: qty,
        unit: String(r[ISSUE_COL.UNIT - 1] || 'Pcs'),
        baseQty: _toValidNumber(r[ISSUE_COL.BASE_QTY - 1], 'Base Qty', true) || qty,
        rate: rate,
        value: qty * rate
      });

      issueMap[issueId].totalQty += qty;
      issueMap[issueId].totalValue += qty * rate;
    }

    const sorted = Object.values(issueMap).sort(function(a, b) {
      const timeA = a.dateRaw ? new Date(a.dateRaw).getTime() : 0;
      const timeB = b.dateRaw ? new Date(b.dateRaw).getTime() : 0;
      return timeB - timeA;
    });

    return buildResponse(true, sorted);
  } catch (error) {
    Log.error('[getIssueData] Error:', error.message);
    logAction('ERROR', 'getIssueData', 'module_issue', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to load issued stock records: ' + error.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────
// ISSUE CREATION
// ─────────────────────────────────────────────────────────────────────────

/**
 * Builds a { lowercased-name: canonical-name } map from the Vendor Master,
 * read once per call — mirrors _getItemUnitInfoMap/_getUnitsMap's shape.
 * Used by saveIssueStock to detect when "Issued To" matches a known vendor.
 * @private
 */
function _getVendorNameMap() {
  const map = {};
  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.VENDORS);
    const lastRow = sheet ? sheet.getLastRow() : 0;
    if (!sheet || lastRow < 2) return map;

    const data = sheet.getRange(2, VENDORS_COL.VENDOR_NAME, lastRow - 1, 1).getValues();
    data.forEach(function(row) {
      const name = String(row[0] || '').trim();
      if (!name) return;
      map[name.toLowerCase()] = name;
    });
  } catch (e) {
    // Vendors sheet not found/accessible — treat as "no known vendors".
  }
  return map;
}

/**
 * saveIssueStock(formData)
 *
 * Records ad-hoc item issuance. Issued To is required; Reference is optional.
 * Each item row requires: name, qty, unit.
 * BASE_QTY is computed via unit conversion and debits stock.
 * If "Issued To" matches a Vendor Master entry (case-insensitive), the issue
 * is also linked to that vendor for Vendor Ledger purposes — there is no
 * separate Vendor field; typing/selecting a known vendor name is enough.
 *
 * @param {Object} formData
 *   @param {string} formData.date - DD/MM/YYYY or YYYY-MM-DD
 *   @param {string} formData.issuedTo - Who/what this issuance is for (matched against Vendor Master)
 *   @param {string} [formData.reference] - Optional free-text reference (e.g. Lot #)
 *   @param {string} [formData.remarks] - Optional header-level remarks
 *   @param {Array|string} formData.items - Each: {name, size, qty, unit, rate}
 * @returns {Object} API response
 */
function saveIssueStock(formData) {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(ISSUE_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again in a moment.');
  }

  try {
    const sheet = _getIssueSheet();
    _ensureIssueColumns(sheet);

    let items;
    try {
      items = typeof formData.items === 'string'
        ? JSON.parse(formData.items)
        : (formData.items || []);
    } catch (error) {
      return buildResponse(false, null, 'Invalid items data: could not parse JSON.');
    }

    if (!Array.isArray(items) || items.length === 0) {
      return buildResponse(false, null, 'Cannot issue stock with zero items. Add at least one item.');
    }

    const issuedTo = sanitizeString(formData.issuedTo || '', 'issuedTo');
    if (!issuedTo) {
      return buildResponse(false, null, 'Issued To is required.');
    }

    const issueDateNative = toSafeDateObject(formData.date);
    if (!issueDateNative) {
      return buildResponse(false, null, 'Invalid issue date. Accepted formats: DD/MM/YYYY or YYYY-MM-DD.');
    }

    const issueId = _generateIssueId();
    const reference = sanitizeString(formData.reference || '', 'reference');
    const remarks = sanitizeString(formData.remarks || '', 'remarks');

    const vendorNameMap = _getVendorNameMap();
    const vendor = vendorNameMap[issuedTo.toLowerCase()] || '';

    const itemUnitMap = _getItemUnitInfoMap();
    const unitsMap = _getUnitsMap();

    const newRows = items.map(function(item) {
      const qty = _toValidNumber(item.qty, 'Qty', false);
      const unit = sanitizeString(item.unit || 'Pcs', 'item.unit');

      const unitInfo = _lookupItemUnitInfo(itemUnitMap, item.name, item.size || '');
      let baseQty = qty;
      try {
        baseQty = convertQtyToBaseUnit(qty, unit, unitInfo, unitsMap);
      } catch (e) {
        baseQty = qty;
      }

      const rate = _toValidNumber(item.rate || 0, 'Rate', true);

      return [
        issueId,                                                // 1: ISSUE_ID
        issueDateNative,                                         // 2: DATE
        issuedTo,                                                // 3: ISSUED_TO
        reference,                                               // 4: REFERENCE
        sanitizeString(item.name, 'item.name'),                  // 5: ITEM_NAME
        sanitizeString(item.size || '', 'item.size'),            // 6: SIZE
        qty,                                                     // 7: QTY
        unit,                                                    // 8: UNIT
        remarks,                                                 // 9: REMARKS
        baseQty,                                                 // 10: BASE_QTY
        rate,                                                    // 11: RATE
        vendor                                                   // 12: VENDOR
      ];
    });

    const appendRow = sheet.getLastRow() + 1;
    sheet.getRange(appendRow, 1, newRows.length, newRows[0].length).setValues(newRows);

    if (typeof recalculateStock === 'function') {
      recalculateStock();
    }

    SpreadsheetApp.flush();

    logAction('CREATE', APP_CONFIG.SHEETS.ISSUE, issueId, `Items: ${items.length}`, 'SUCCESS');
    return buildResponse(true, { issueId }, `Stock issue ${issueId} logged successfully.`);
  } catch (error) {
    Log.error('[saveIssueStock] Error:', error.message);
    logAction('ERROR', 'saveIssueStock', '', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to save stock issue: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────────────────
// ISSUE DELETION
// ─────────────────────────────────────────────────────────────────────────

function deleteIssueBulk(issueIds) {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(ISSUE_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const targetSet = new Set(
      (issueIds || []).map(id => String(id).trim()).filter(Boolean)
    );

    if (targetSet.size === 0) {
      return buildResponse(true, null, 'No issued stock records selected.');
    }

    const sheet = _getIssueSheet();

    const { rowsDeleted, deletedIds } = _rewriteWithoutMatchingRowsBulk(
      sheet,
      APP_CONFIG.ISSUE_SETTINGS.DATA_START_ROW,
      ISSUE_COL.ISSUE_ID,
      targetSet
    );

    if (typeof recalculateStock === 'function') {
      recalculateStock();
    }

    SpreadsheetApp.flush();

    const msg = `Deleted ${targetSet.size} issued stock record(s) (${rowsDeleted} row(s) removed).`;
    logAction('BULK_DELETE', APP_CONFIG.SHEETS.ISSUE, 'multiple', msg, 'SUCCESS');

    return buildResponse(true, { deletedIds }, msg);
  } catch (error) {
    Log.error('[deleteIssueBulk] Error:', error.message);
    logAction('ERROR', 'deleteIssueBulk', 'multiple', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete issued stock records: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Repoints historical Issued Stock Log rows from an item's old name+size to
 * its new name+size. Called from module_items.js's
 * _propagateItemIdentityChange after a rename (saveItem) or a merge
 * (mergeItemEdit) — mirrors backfillBillItemRefs (module_bill.js). Runs
 * under the caller's existing document lock — does not acquire its own.
 * @param {string} oldName
 * @param {string} oldSize
 * @param {string} newName
 * @param {string} newSize
 */
function backfillIssueItemRefs(oldName, oldSize, newName, newSize) {
  const sheet = getSheet(APP_CONFIG.SHEETS.ISSUE);
  if (!sheet) return;

  const startRow = APP_CONFIG.ISSUE_SETTINGS.DATA_START_ROW;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  const numRows = lastRow - startRow + 1;
  const range = sheet.getRange(startRow, ISSUE_COL.ITEM_NAME, numRows, 2);
  const values = range.getValues();

  const tOldName = String(oldName || '').trim().toLowerCase();
  const tOldSize = String(oldSize || '').trim().toLowerCase();

  let changed = false;
  for (let i = 0; i < values.length; i++) {
    const rowName = String(values[i][0] || '').trim().toLowerCase();
    const rowSize = String(values[i][1] || '').trim().toLowerCase();
    if (rowName === tOldName && rowSize === tOldSize) {
      values[i][0] = newName;
      values[i][1] = newSize;
      changed = true;
    }
  }

  if (changed) {
    range.setValues(values);
  }
}
