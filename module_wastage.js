/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_wastage.gs — WASTAGE LOG MODULE
 *
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Records component-wise material wastage / losses.
 * Vendor is optional. Each entry can have multiple item rows.
 * BASE_QTY debits Stock the same way vendor Returns do (see module_stock.js).
 *
 * Dependencies (shared globals, same Apps Script project namespace):
 * ───────────────────────────────────────────────────────────────────────────
 * - config.js (APP_CONFIG, WASTAGE_COL)
 * - utils.js (getSheet, buildResponse, sanitizeString, toSafeDateString)
 * - module_bill.js (_toValidNumber — shared numeric validator)
 * - module_units.js / module_items.js (unit conversion helpers)
 * ═══════════════════════════════════════════════════════════════════════════
 */

const WASTAGE_LOCK_TIMEOUT_MS = 15000;

// ─────────────────────────────────────────────────────────────────────────
// SHEET INITIALIZATION / SELF-HEALING
// ─────────────────────────────────────────────────────────────────────────

function initWastageSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.WASTAGE);
    if (!sheet) {
      sheet = ss.insertSheet(APP_CONFIG.SHEETS.WASTAGE);
    }

    const headers = [
      'Wastage ID', 'Date', 'Vendor', 'Item Name', 'Size',
      'Qty', 'Unit', 'Reason', 'Remarks', 'Base Qty'
    ];

    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Wastage Log sheet initialized successfully.');
  } catch (error) {
    Log.error('[initWastageSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize Wastage Log sheet: ' + error.message);
  }
}

function _getWastageSheet() {
  try {
    return getSheet(APP_CONFIG.SHEETS.WASTAGE);
  } catch (e) {
    initWastageSheet();
    return getSheet(APP_CONFIG.SHEETS.WASTAGE);
  }
}

// ─────────────────────────────────────────────────────────────────────────
// ID GENERATION
// ─────────────────────────────────────────────────────────────────────────

function _generateWastageId() {
  const now = new Date();
  const datePart = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
  const timePart = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HHmmss');
  return 'WST-' + datePart + '-' + timePart;
}

// ─────────────────────────────────────────────────────────────────────────
// WASTAGE DATA RETRIEVAL
// ─────────────────────────────────────────────────────────────────────────

/**
 * getWastageData()
 *
 * Retrieves all wastage records, grouped by Wastage ID.
 * Returns newest first.
 *
 * @returns {Object} API response — data: Array of wastage groups:
 *   { wastageId, date, dateRaw, vendor, remarks,
 *     items: [{name, size, qty, unit, reason, baseQty}],
 *     totalQty }
 */
function getWastageData() {
  try {
    const sheet = _getWastageSheet();

    const startRow = APP_CONFIG.WASTAGE_SETTINGS.DATA_START_ROW;
    const lastRow = sheet.getLastRow();

    if (lastRow < startRow) {
      return buildResponse(true, []);
    }

    const numCols = WASTAGE_COL.BASE_QTY;
    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, numCols).getValues();

    const wastageMap = {};

    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const wastageId = String(r[WASTAGE_COL.WASTAGE_ID - 1] || '').trim();
      if (!wastageId) continue;

      if (!wastageMap[wastageId]) {
        const rawDate = r[WASTAGE_COL.DATE - 1];
        const parsedDate = toSafeDateString(rawDate);

        const isoDateStr = parsedDate
          ? parsedDate.split('/').reverse().join('-')
          : '';

        wastageMap[wastageId] = {
          wastageId: wastageId,
          date: parsedDate || (rawDate ? String(rawDate) : 'N/A'),
          dateRaw: isoDateStr,
          vendor: String(r[WASTAGE_COL.VENDOR - 1] || ''),
          remarks: String(r[WASTAGE_COL.REMARKS - 1] || ''),
          items: [],
          totalQty: 0
        };
      }

      const qty = _toValidNumber(r[WASTAGE_COL.QTY - 1], 'Qty', false);

      wastageMap[wastageId].items.push({
        name: String(r[WASTAGE_COL.ITEM_NAME - 1] || ''),
        size: String(r[WASTAGE_COL.SIZE - 1] || ''),
        qty: qty,
        unit: String(r[WASTAGE_COL.UNIT - 1] || 'Pcs'),
        reason: String(r[WASTAGE_COL.REASON - 1] || ''),
        remarks: String(r[WASTAGE_COL.REMARKS - 1] || ''),
        baseQty: _toValidNumber(r[WASTAGE_COL.BASE_QTY - 1], 'Base Qty', true) || qty
      });

      wastageMap[wastageId].totalQty += qty;
    }

    const sorted = Object.values(wastageMap).sort(function(a, b) {
      const timeA = a.dateRaw ? new Date(a.dateRaw).getTime() : 0;
      const timeB = b.dateRaw ? new Date(b.dateRaw).getTime() : 0;
      return timeB - timeA;
    });

    return buildResponse(true, sorted);
  } catch (error) {
    Log.error('[getWastageData] Error:', error.message);
    logAction('ERROR', 'getWastageData', 'module_wastage', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to load wastage records: ' + error.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────
// WASTAGE CREATION
// ─────────────────────────────────────────────────────────────────────────

/**
 * saveWastage(formData)
 *
 * Records component-wise wastage. Vendor is optional.
 * Each item row requires: name, qty, unit, reason.
 * BASE_QTY is computed via unit conversion and debits stock.
 *
 * @param {Object} formData
 *   @param {string} formData.date - DD/MM/YYYY or YYYY-MM-DD
 *   @param {string} [formData.vendor] - Optional vendor name
 *   @param {string} [formData.remarks] - Optional header-level remarks
 *   @param {Array|string} formData.items - Each: {name, size, qty, unit, reason}
 * @returns {Object} API response
 */
function saveWastage(formData) {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(WASTAGE_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again in a moment.');
  }

  try {
    const sheet = _getWastageSheet();

    let items;
    try {
      items = typeof formData.items === 'string'
        ? JSON.parse(formData.items)
        : (formData.items || []);
    } catch (error) {
      return buildResponse(false, null, 'Invalid items data: could not parse JSON.');
    }

    if (!Array.isArray(items) || items.length === 0) {
      return buildResponse(false, null, 'Cannot save wastage with zero items. Add at least one item.');
    }

    const wastageDateNative = toSafeDateObject(formData.date);
    if (!wastageDateNative) {
      return buildResponse(false, null, 'Invalid wastage date. Accepted formats: DD/MM/YYYY or YYYY-MM-DD.');
    }

    const wastageId = _generateWastageId();
    const vendor = sanitizeString(formData.vendor || '', 'vendor');
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
        wastageId,                                              // 1: WASTAGE_ID
        wastageDateNative,                                      // 2: DATE
        vendor,                                                 // 3: VENDOR
        sanitizeString(item.name, 'item.name'),                 // 4: ITEM_NAME
        sanitizeString(item.size || '', 'item.size'),           // 5: SIZE
        qty,                                                    // 6: QTY
        unit,                                                   // 7: UNIT
        sanitizeString(item.reason || '', 'item.reason'),       // 8: REASON
        remarks,                                                // 9: REMARKS
        baseQty                                                 // 10: BASE_QTY
      ];
    });

    const appendRow = sheet.getLastRow() + 1;
    sheet.getRange(appendRow, 1, newRows.length, newRows[0].length).setValues(newRows);

    if (typeof recalculateStock === 'function') {
      recalculateStock();
    }

    SpreadsheetApp.flush();

    logAction('CREATE', APP_CONFIG.SHEETS.WASTAGE, wastageId, `Items: ${items.length}`, 'SUCCESS');
    return buildResponse(true, { wastageId }, `Wastage ${wastageId} logged successfully.`);
  } catch (error) {
    Log.error('[saveWastage] Error:', error.message);
    logAction('ERROR', 'saveWastage', '', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to save wastage: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────────────────
// WASTAGE DELETION
// ─────────────────────────────────────────────────────────────────────────

function deleteWastageBulk(wastageIds) {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(WASTAGE_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const targetSet = new Set(
      (wastageIds || []).map(id => String(id).trim()).filter(Boolean)
    );

    if (targetSet.size === 0) {
      return buildResponse(true, null, 'No wastage records selected.');
    }

    const sheet = _getWastageSheet();

    const { rowsDeleted, deletedIds } = _rewriteWithoutMatchingRowsBulk(
      sheet,
      APP_CONFIG.WASTAGE_SETTINGS.DATA_START_ROW,
      WASTAGE_COL.WASTAGE_ID,
      targetSet
    );

    if (typeof recalculateStock === 'function') {
      recalculateStock();
    }

    SpreadsheetApp.flush();

    const msg = `Deleted ${targetSet.size} wastage record(s) (${rowsDeleted} row(s) removed).`;
    logAction('BULK_DELETE', APP_CONFIG.SHEETS.WASTAGE, 'multiple', msg, 'SUCCESS');

    return buildResponse(true, { deletedIds }, msg);
  } catch (error) {
    Log.error('[deleteWastageBulk] Error:', error.message);
    logAction('ERROR', 'deleteWastageBulk', 'multiple', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete wastage records: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Repoints historical Wastage Log rows from an item's old name+size to its
 * new name+size. Called from module_items.js's _propagateItemIdentityChange
 * after a rename (saveItem) or a merge (mergeItemEdit) — mirrors
 * backfillBillItemRefs (module_bill.js). Runs under the caller's existing
 * document lock — does not acquire its own.
 * @param {string} oldName
 * @param {string} oldSize
 * @param {string} newName
 * @param {string} newSize
 */
function backfillWastageItemRefs(oldName, oldSize, newName, newSize) {
  const sheet = getSheet(APP_CONFIG.SHEETS.WASTAGE);
  if (!sheet) return;

  const startRow = APP_CONFIG.WASTAGE_SETTINGS.DATA_START_ROW;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  const numRows = lastRow - startRow + 1;
  const range = sheet.getRange(startRow, WASTAGE_COL.ITEM_NAME, numRows, 2);
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
