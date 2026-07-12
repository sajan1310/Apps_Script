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
      'Qty', 'Unit', 'Remarks', 'Base Qty'
    ];

    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Issued Stock Log sheet initialized successfully.');
  } catch (error) {
    console.error('[initIssueSheet] Error:', error.message);
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
 *   { issueId, date, dateRaw, issuedTo, reference, remarks,
 *     items: [{name, size, qty, unit, baseQty}], totalQty }
 */
function getIssueData() {
  try {
    const sheet = _getIssueSheet();

    const startRow = APP_CONFIG.ISSUE_SETTINGS.DATA_START_ROW;
    const lastRow = sheet.getLastRow();

    if (lastRow < startRow) {
      return buildResponse(true, []);
    }

    const numCols = ISSUE_COL.BASE_QTY;
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
          remarks: String(r[ISSUE_COL.REMARKS - 1] || ''),
          items: [],
          totalQty: 0
        };
      }

      const qty = _toValidNumber(r[ISSUE_COL.QTY - 1], 'Qty', false);

      issueMap[issueId].items.push({
        name: String(r[ISSUE_COL.ITEM_NAME - 1] || ''),
        size: String(r[ISSUE_COL.SIZE - 1] || ''),
        qty: qty,
        unit: String(r[ISSUE_COL.UNIT - 1] || 'Pcs'),
        baseQty: _toValidNumber(r[ISSUE_COL.BASE_QTY - 1], 'Base Qty', true) || qty
      });

      issueMap[issueId].totalQty += qty;
    }

    const sorted = Object.values(issueMap).sort(function(a, b) {
      const timeA = a.dateRaw ? new Date(a.dateRaw).getTime() : 0;
      const timeB = b.dateRaw ? new Date(b.dateRaw).getTime() : 0;
      return timeB - timeA;
    });

    return buildResponse(true, sorted);
  } catch (error) {
    console.error('[getIssueData] Error:', error.message);
    logAction('ERROR', 'getIssueData', 'module_issue', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to load issued stock records: ' + error.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────
// ISSUE CREATION
// ─────────────────────────────────────────────────────────────────────────

/**
 * saveIssueStock(formData)
 *
 * Records ad-hoc item issuance. Issued To is required; Reference is optional.
 * Each item row requires: name, qty, unit.
 * BASE_QTY is computed via unit conversion and debits stock.
 *
 * @param {Object} formData
 *   @param {string} formData.date - DD/MM/YYYY or YYYY-MM-DD
 *   @param {string} formData.issuedTo - Who/what this issuance is for
 *   @param {string} [formData.reference] - Optional free-text reference (e.g. Lot #)
 *   @param {string} [formData.remarks] - Optional header-level remarks
 *   @param {Array|string} formData.items - Each: {name, size, qty, unit}
 * @returns {Object} API response
 */
function saveIssueStock(formData) {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(ISSUE_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again in a moment.');
  }

  try {
    const sheet = _getIssueSheet();

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
        baseQty                                                  // 10: BASE_QTY
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
    console.error('[saveIssueStock] Error:', error.message);
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
    console.error('[deleteIssueBulk] Error:', error.message);
    logAction('ERROR', 'deleteIssueBulk', 'multiple', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete issued stock records: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}
