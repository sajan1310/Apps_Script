/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_return.gs — VENDOR RETURN LEDGER MODULE
 *
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Records goods sent back to a vendor (defective, excess, wrong item, etc.).
 * Mirrors module_bill.js but in reverse: QTY/BASE_QTY here are debited from
 * Stock (see module_stock.js's _getBilledAndConsumedQtyMaps) instead of
 * credited, and TOTAL reduces the amount owed to the vendor.
 *
 * Dependencies (shared globals, same Apps Script project namespace):
 * ───────────────────────────────────────────────────────────────────────────
 * - config.js (APP_CONFIG, RETURN_COL)
 * - utils.js (getSheet, buildResponse, sanitizeString, toSafeDateString,
 *   deleteRowsById)
 * - module_po.js (_rewriteWithoutMatchingRowsBulk)
 * - module_bill.js (_toValidNumber — shared numeric validator)
 * - module_units.js / module_items.js (unit conversion helpers)
 * ═══════════════════════════════════════════════════════════════════════════
 */

const RETURN_LOCK_TIMEOUT_MS = 15000;

// ─────────────────────────────────────────────────────────────────────────
// VALIDATION HELPERS
// ─────────────────────────────────────────────────────────────────────────

function _validateReturnNumber(returnNumber) {
  const s = String(returnNumber || '').trim();
  if (!s) {
    throw new Error('Return number must not be empty.');
  }
  sanitizeString(s, 'returnNumber');
  return s;
}

// ─────────────────────────────────────────────────────────────────────────
// SHEET INITIALIZATION / SELF-HEALING
// ─────────────────────────────────────────────────────────────────────────

function initReturnSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.RETURN);
    if (!sheet) {
      sheet = ss.insertSheet(APP_CONFIG.SHEETS.RETURN);
    }

    const headers = [
      'Return Number', 'Return Date', 'Vendor', 'Contact', 'Bill Number',
      'Item Name', 'Size', 'Narration', 'Qty', 'Unit', 'Price', 'Total',
      'Reason', 'Remarks', 'Base Qty', 'Base Rate'
    ];

    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Return Ledger sheet initialized successfully.');
  } catch (error) {
    Log.error('[initReturnSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize Return Ledger sheet: ' + error.message);
  }
}

function _getReturnSheet() {
  try {
    return getSheet(APP_CONFIG.SHEETS.RETURN);
  } catch (e) {
    initReturnSheet();
    return getSheet(APP_CONFIG.SHEETS.RETURN);
  }
}

// ─────────────────────────────────────────────────────────────────────────
// RETURN DATA RETRIEVAL
// ─────────────────────────────────────────────────────────────────────────

/**
 * getReturnData()
 *
 * Retrieves all vendor returns from the sheet and aggregates by return number,
 * mirroring getBillData()'s shape and conventions.
 *
 * @returns {Object} API response — data: Array of return objects, each:
 *   { returnNumber, returnDate, returnDateRaw, vendor, contact, billNumber,
 *     remarks, items: [{name, size, narration, unit, qty, price, lineTotal,
 *     reason, baseQty, baseRate}], totalQty, totalAmount }
 */
function getReturnData() {
  try {
    const sheet = _getReturnSheet();

    const startRow = APP_CONFIG.RETURN_SETTINGS.DATA_START_ROW;
    const lastRow = sheet.getLastRow();

    if (lastRow < startRow) {
      return buildResponse(true, []);
    }

    const numCols = RETURN_COL.BASE_RATE;
    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, numCols).getValues();

    const returnMap = {};

    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const returnNum = String(r[RETURN_COL.RETURN_NUMBER - 1] || '').trim();
      if (!returnNum) continue;

      if (!returnMap[returnNum]) {
        const rawDate = r[RETURN_COL.RETURN_DATE - 1];
        const parsedDate = toSafeDateString(rawDate);

        if (!parsedDate && rawDate) {
          Log.warn(`[getReturnData] Unparseable date in return #${returnNum}: "${rawDate}"`);
        }

        const isoDateStr = parsedDate
          ? parsedDate.split('/').reverse().join('-')
          : '';

        returnMap[returnNum] = {
          returnNumber: returnNum,
          returnDate: parsedDate || (rawDate ? String(rawDate) : 'N/A'),
          returnDateRaw: isoDateStr,
          vendor: String(r[RETURN_COL.VENDOR - 1] || ''),
          contact: String(r[RETURN_COL.CONTACT - 1] || ''),
          billNumber: String(r[RETURN_COL.BILL_NUMBER - 1] || ''),
          remarks: String(r[RETURN_COL.REMARKS - 1] || ''),
          items: [],
          totalQty: 0,
          totalAmount: 0
        };
      }

      const qty = _toValidNumber(r[RETURN_COL.QTY - 1], 'Qty');
      const price = _toValidNumber(r[RETURN_COL.PRICE - 1], 'Price');
      const lineTotal = parseFloat((qty * price).toFixed(2));

      returnMap[returnNum].items.push({
        name: String(r[RETURN_COL.ITEM_NAME - 1] || ''),
        size: String(r[RETURN_COL.SIZE - 1] || ''),
        narration: String(r[RETURN_COL.NARRATION - 1] || ''),
        unit: String(r[RETURN_COL.UNIT - 1] || 'Pcs'),
        qty: qty,
        price: price,
        lineTotal: lineTotal,
        reason: String(r[RETURN_COL.REASON - 1] || ''),
        baseQty: _toValidNumber(r[RETURN_COL.BASE_QTY - 1], 'Base Qty', true) || qty,
        baseRate: _toValidNumber(r[RETURN_COL.BASE_RATE - 1], 'Base Rate', true) || price
      });

      returnMap[returnNum].totalQty += qty;
      returnMap[returnNum].totalAmount += lineTotal;
    }

    const sorted = Object.values(returnMap).sort(function(a, b) {
      const timeA = a.returnDateRaw ? new Date(a.returnDateRaw).getTime() : 0;
      const timeB = b.returnDateRaw ? new Date(b.returnDateRaw).getTime() : 0;
      return timeB - timeA;
    });

    return buildResponse(true, sorted);
  } catch (error) {
    Log.error('[getReturnData] Error:', error.message);
    logAction('ERROR', 'getReturnData', 'module_return', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to load returns: ' + error.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────
// RETURN CREATION
// ─────────────────────────────────────────────────────────────────────────

/**
 * saveReturn(formData)
 *
 * Records goods returned to a vendor. Mirrors saveBill()'s validation and
 * write strategy (including edit-in-place support), minus GST (a return is
 * a straight credit, not a taxable line).
 *
 * @param {Object} formData
 *   @param {string} formData.returnNumber - Unique return/debit-note identifier
 *   @param {string} formData.returnDate - Date (DD/MM/YYYY or YYYY-MM-DD)
 *   @param {string} formData.vendor
 *   @param {string} formData.contact
 *   @param {string} formData.billNumber - Optional reference to the original bill
 *   @param {string} formData.remarks
 *   @param {Array|string} formData.items - Each: {name, size, narration, unit, qty, price, reason}
 * @returns {Object} API response
 */
function saveReturn(formData) {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(RETURN_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again in a moment.');
  }

  try {
    const sheet = _getReturnSheet();

    let items;
    try {
      items = typeof formData.items === 'string'
        ? JSON.parse(formData.items)
        : (formData.items || []);
    } catch (error) {
      return buildResponse(false, null, 'Invalid items data: could not parse JSON.');
    }

    if (!Array.isArray(items) || items.length === 0) {
      return buildResponse(false, null, 'Cannot save a return with zero items. Add at least one item.');
    }

    const oldReturnNumber = String(formData.existingReturnNumber || '').trim();
    const isEdit = !!oldReturnNumber;

    let returnNumber = String(formData.returnNumber || '').trim();
    if (!returnNumber) {
      const now = new Date();
      const pad = n => String(n).padStart(2, '0');
      returnNumber = `RET-${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
    } else {
      returnNumber = _validateReturnNumber(returnNumber);
    }
    const startRow = APP_CONFIG.RETURN_SETTINGS.DATA_START_ROW;

    if (_returnNumberExists(sheet, startRow, returnNumber, isEdit ? oldReturnNumber : null)) {
      return buildResponse(
        false,
        null,
        isEdit
          ? `The return number "${returnNumber}" already exists in another record. Edit aborted.`
          : `Return number "${returnNumber}" already exists. Please use a unique return number.`
      );
    }

    let firstMatchRow = -1;
    if (isEdit) {
      firstMatchRow = _findFirstReturnRow(sheet, startRow, oldReturnNumber);
      if (firstMatchRow === -1) {
        throw new Error(
          `Original Return #${oldReturnNumber} not found in sheet. Edit aborted to prevent data corruption.`
        );
      }
    }

    const returnDate = toSafeDateString(formData.returnDate);
    if (!returnDate) {
      return buildResponse(false, null, 'Invalid return date. Accepted formats: DD/MM/YYYY or YYYY-MM-DD.');
    }

    const returnDateParts = returnDate.split('/');
    const returnDateNative = new Date(
      parseInt(returnDateParts[2], 10),
      parseInt(returnDateParts[1], 10) - 1,
      parseInt(returnDateParts[0], 10)
    );

    const vendor = sanitizeString(formData.vendor, 'vendor');
    const contact = sanitizeString(formData.contact, 'contact');
    const billNumber = sanitizeString(formData.billNumber || '', 'billNumber');
    const remarks = sanitizeString(formData.remarks, 'remarks');

    const itemUnitMap = _getItemUnitInfoMap();
    const unitsMap = _getUnitsMap();

    const newRows = items.map(function(item) {
      const qty = _toValidNumber(item.qty, 'Qty', false);
      const price = _toValidNumber(item.price, 'Price');
      const unit = sanitizeString(item.unit || 'Pcs', 'item.unit');
      const lineTotal = parseFloat((qty * price).toFixed(2));

      const unitInfo = _lookupItemUnitInfo(itemUnitMap, item.name, item.size || '');
      let baseQty = qty;
      let baseRate = price;
      try {
        baseQty = convertQtyToBaseUnit(qty, unit, unitInfo, unitsMap);
        baseRate = convertRateToBaseUnit(price, unit, unitInfo, unitsMap);
      } catch (e) {
        baseQty = qty;
        baseRate = price;
      }

      return [
        returnNumber,                                       // 1: RETURN_NUMBER
        returnDateNative,                                   // 2: RETURN_DATE
        vendor,                                             // 3: VENDOR
        contact,                                            // 4: CONTACT
        billNumber,                                         // 5: BILL_NUMBER
        sanitizeString(item.name, 'item.name'),             // 6: ITEM_NAME
        sanitizeString(item.size || '', 'item.size'),       // 7: SIZE
        sanitizeString(item.narration || '', 'item.narration'), // 8: NARRATION
        qty,                                                // 9: QTY
        unit,                                               // 10: UNIT
        price,                                              // 11: PRICE
        lineTotal,                                          // 12: TOTAL
        sanitizeString(item.reason || '', 'item.reason'),   // 13: REASON
        remarks,                                            // 14: REMARKS
        baseQty,                                            // 15: BASE_QTY
        baseRate                                            // 16: BASE_RATE
      ];
    });

    if (isEdit) {
      _rewriteReturnsWithoutMatchingRows(sheet, startRow, oldReturnNumber);
      _insertReturnRowsAt(sheet, firstMatchRow, newRows);
    } else {
      const appendRow = sheet.getLastRow() + 1;
      sheet.getRange(appendRow, 1, newRows.length, newRows[0].length).setValues(newRows);
    }

    if (typeof recalculateStock === 'function') {
      recalculateStock();
    }

    SpreadsheetApp.flush();

    const msg = isEdit
      ? `Return #${returnNumber} updated successfully.`
      : `Return #${returnNumber} recorded successfully.`;
    logAction(isEdit ? 'UPDATE' : 'CREATE', APP_CONFIG.SHEETS.RETURN, returnNumber, `Items: ${items.length}`, 'SUCCESS');

    return buildResponse(true, { returnNumber }, msg);
  } catch (error) {
    Log.error('[saveReturn] Error:', error.message);
    logAction('ERROR', 'saveReturn', formData.returnNumber, error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to save return: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────────────────
// RETURN DELETION
// ─────────────────────────────────────────────────────────────────────────

function deleteReturn(returnNumber) {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(RETURN_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const validReturnNumber = _validateReturnNumber(returnNumber);
    const sheet = _getReturnSheet();

    deleteRowsById(
      validReturnNumber,
      sheet,
      APP_CONFIG.RETURN_SETTINGS.DATA_START_ROW,
      RETURN_COL.RETURN_NUMBER
    );

    if (typeof recalculateStock === 'function') {
      recalculateStock();
    }

    SpreadsheetApp.flush();

    const msg = `Return #${validReturnNumber} deleted successfully.`;
    logAction('DELETE', APP_CONFIG.SHEETS.RETURN, validReturnNumber, msg, 'SUCCESS');

    return buildResponse(true, null, msg);
  } catch (error) {
    Log.error('[deleteReturn] Error:', error.message);
    logAction('ERROR', 'deleteReturn', returnNumber, error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete return: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

function deleteReturnsBulk(returnNumbers) {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(RETURN_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const targetSet = new Set(
      (returnNumbers || []).map(r => String(r).trim()).filter(Boolean)
    );

    if (targetSet.size === 0) {
      return buildResponse(true, null, 'No returns selected.');
    }

    const sheet = _getReturnSheet();

    const { rowsDeleted, deletedIds } = _rewriteWithoutMatchingRowsBulk(
      sheet,
      APP_CONFIG.RETURN_SETTINGS.DATA_START_ROW,
      RETURN_COL.RETURN_NUMBER,
      targetSet
    );

    if (typeof recalculateStock === 'function') {
      recalculateStock();
    }

    SpreadsheetApp.flush();

    const msg = `Deleted ${targetSet.size} return(s) (${rowsDeleted} row(s) removed).`;
    logAction('BULK_DELETE', APP_CONFIG.SHEETS.RETURN, 'multiple', msg, 'SUCCESS');

    return buildResponse(true, { deletedIds }, msg);
  } catch (error) {
    Log.error('[deleteReturnsBulk] Error:', error.message);
    logAction('ERROR', 'deleteReturnsBulk', 'multiple', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete returns: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────────────────
// PRIVATE HELPERS — duplicate-check, row lookup, rewrite, insert
// (mirror module_bill.js's equivalents, scoped to the Return sheet)
// ─────────────────────────────────────────────────────────────────────────

function _returnNumberExists(sheet, startRow, returnNumber, excludeReturnNumber = null) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return false;

  const values = sheet
    .getRange(startRow, RETURN_COL.RETURN_NUMBER, lastRow - startRow + 1, 1)
    .getValues();

  const target = String(returnNumber).trim().toLowerCase();
  const exclude = excludeReturnNumber ? String(excludeReturnNumber).trim().toLowerCase() : null;

  for (let i = 0; i < values.length; i++) {
    const val = String(values[i][0] || '').trim().toLowerCase();
    if (val === target) {
      if (exclude && val === exclude) continue;
      return true;
    }
  }
  return false;
}

function _findFirstReturnRow(sheet, startRow, returnNumber) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return -1;

  const values = sheet
    .getRange(startRow, RETURN_COL.RETURN_NUMBER, lastRow - startRow + 1, 1)
    .getValues();

  const target = String(returnNumber).trim().toLowerCase();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim().toLowerCase() === target) {
      return startRow + i;
    }
  }
  return -1;
}

function _rewriteReturnsWithoutMatchingRows(sheet, startRow, returnNumber) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < startRow) return;

  const numRows = lastRow - startRow + 1;
  const allRows = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
  const matchIndex = RETURN_COL.RETURN_NUMBER - 1;
  const targetStr = String(returnNumber).trim().toLowerCase();

  const kept = allRows.filter(function(row) {
    return String(row[matchIndex] || '').trim().toLowerCase() !== targetStr;
  });

  sheet.getRange(startRow, 1, numRows, lastCol).clearContent();

  if (kept.length > 0) {
    sheet.getRange(startRow, 1, kept.length, lastCol).setValues(kept);
  }

  const rowsToDelete = numRows - kept.length;
  if (rowsToDelete > 0) {
    sheet.deleteRows(startRow + kept.length, rowsToDelete);
  }
}

function _insertReturnRowsAt(sheet, rowIndex, rows) {
  if (!rows || rows.length === 0) return;

  const currentLastRow = sheet.getLastRow();
  if (rowIndex <= currentLastRow) {
    sheet.insertRows(rowIndex, rows.length);
  }

  sheet.getRange(rowIndex, 1, rows.length, rows[0].length).setValues(rows);
}
