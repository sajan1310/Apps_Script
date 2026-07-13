/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_dispatch.gs — DISPATCH MODULE
 *
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * - "Ready to Dispatch" view: finished-goods inventory computed on read from
 *   completed Production lots minus what has already been dispatched (mirrors
 *   how recalculateStock derives raw-material stock, but needs no sheet of
 *   its own).
 * - Dispatched Goods ledger CRUD (Dispatch sheet). Entries may optionally
 *   reference a PI / Estimate (Order Number) or be standalone "direct
 *   supply" dispatches.
 *
 * Sheet Layout (Dispatch):
 * ───────────────────────────────────────────────────────────────────────────
 * Col A (1):   Dispatch Number (e.g. DSP-1001)
 * Col B (2):   Dispatch Date (DD/MM/YYYY)
 * Col C (3):   Order Number (optional reference to Client Orders sheet)
 * Col D (4):   Client Name
 * Col E (5):   Product_ID (reference to BOM.Product ID)
 * Col F (6):   Product Name
 * Col G (7):   Quantity
 * Col H (8):   Transport / Vehicle Details
 * Col I (9):   Remarks
 * Col J (10):  Invoice Number (optional, can be added later)
 * Col K (11):  Private Mark (optional, can be added later)
 * Col L (12):  GR Number (optional, can be added later)
 * ═══════════════════════════════════════════════════════════════════════════
 */

const DISPATCH_LOCK_TIMEOUT_MS = 15000;

/**
 * Initializes the Dispatch sheet with correct headers.
 */
function initDispatchSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.DISPATCH);
    if (!sheet) {
      sheet = ss.insertSheet(APP_CONFIG.SHEETS.DISPATCH);
    }

    const headers = [
      'Dispatch Number',
      'Dispatch Date',
      'Order Number',
      'Client Name',
      'Product_ID',
      'Product Name',
      'Quantity',
      'Transport / Vehicle Details',
      'Remarks',
      'Invoice Number',
      'Private Mark',
      'GR Number',
      'Logistics Contractor',
      'Logistics Rate',
      'Logistics Cost'
    ];

    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Dispatch sheet initialized successfully.');
  } catch (error) {
    Log.error('[initDispatchSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize Dispatch sheet: ' + error.message);
  }
}

/**
 * Backfills the "Invoice Number", "Private Mark" and "GR Number" columns on
 * Dispatch sheets created before this feature existed, so legacy rows don't
 * throw when read/written.
 */
function ensureDispatchExtraColumns(sheet) {
  try {
    if (sheet.getLastColumn() < DISPATCH_COL.GR_NUMBER) {
      const startCol = sheet.getLastColumn() + 1;
      sheet.insertColumnsAfter(sheet.getLastColumn(), DISPATCH_COL.GR_NUMBER - sheet.getLastColumn());
      sheet.getRange(1, startCol, 1, DISPATCH_COL.GR_NUMBER - startCol + 1)
        .setValues([['Invoice Number', 'Private Mark', 'GR Number']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    Log.error('[ensureDispatchExtraColumns] Error:', error.message);
  }
}

/**
 * Backfills the "Logistics Contractor", "Logistics Rate" and "Logistics
 * Cost" columns on Dispatch sheets created before the contractor rate card
 * feature existed, so legacy rows don't throw when read/written.
 */
function ensureDispatchLogisticsColumns(sheet) {
  try {
    if (sheet.getLastColumn() < DISPATCH_COL.LOGISTICS_COST) {
      const startCol = sheet.getLastColumn() + 1;
      sheet.insertColumnsAfter(sheet.getLastColumn(), DISPATCH_COL.LOGISTICS_COST - sheet.getLastColumn());
      sheet.getRange(1, startCol, 1, DISPATCH_COL.LOGISTICS_COST - startCol + 1)
        .setValues([['Logistics Contractor', 'Logistics Rate', 'Logistics Cost']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    Log.error('[ensureDispatchLogisticsColumns] Error:', error.message);
  }
}

/**
 * Auto-generates the next sequential Dispatch Number.
 * Format: DSP-1001, DSP-1002, ...
 */
function getNextDispatchNumber() {
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.DISPATCH);
    } catch (e) {
      initDispatchSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.DISPATCH);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return 'DSP-1001';

    const ids = sheet.getRange(2, DISPATCH_COL.DISPATCH_NUMBER, lastRow - 1, 1).getValues();
    let maxNum = 1000;

    ids.forEach(row => {
      const idStr = String(row[0] || '').trim();
      const match = idStr.match(/^DSP-(\d+)$/i);
      if (match) {
        const num = parseInt(match[1], 10);
        if (num > maxNum) maxNum = num;
      }
    });

    return 'DSP-' + (maxNum + 1);
  } catch (error) {
    Log.error('[getNextDispatchNumber] Error:', error.message);
    return 'DSP-1001';
  }
}

/**
 * Builds a map of produced/dispatched quantities per Product ID, sourced
 * from the Warehouse Pool. A Warehouse Pool row is "fully packed" — and
 * therefore visible to Dispatch — whenever it was produced by a final-stage
 * process: tagged rows (Product Tag set) surface under that Product;
 * untagged final-stage rows still surface under their own Output Item Name,
 * so completing a final-stage lot is always enough to make it Ready to
 * Dispatch without requiring an explicit Product tag. Anything produced by a
 * non-final-stage process stays invisible — that's intermediate WIP.
 * { [key]: { productId, productName, producedQty, dispatchedQty } }
 * @private
 */
function _computeReadyToDispatchMap() {
  const map = {};

  const finalStageIds = new Set(
    (typeof _getAllProcessRecords === 'function' ? _getAllProcessRecords() : [])
      .filter(p => p.isFinalStage)
      .map(p => p.processId.toLowerCase())
  );

  const poolResp = getWarehousePoolData();
  const poolRows = (poolResp && poolResp.data) || [];

  if (poolRows.length === 0) return map;

  // Resolve Product Name for each tagged Product ID from the Products (BOM) sheet.
  const productNameById = {};
  try {
    const bomSheet = getSheet(APP_CONFIG.SHEETS.BOM);
    const lastRow = bomSheet.getLastRow();
    if (lastRow >= 2) {
      const data = bomSheet.getRange(2, 1, lastRow - 1, BOM_COL.PRODUCT_NAME).getValues();
      data.forEach(row => {
        const pid = String(row[BOM_COL.PRODUCT_ID - 1] || '').trim();
        if (pid) productNameById[pid.toLowerCase()] = String(row[BOM_COL.PRODUCT_NAME - 1] || '').trim();
      });
    }
  } catch (e) { /* Products sheet not initialized yet */ }

  poolRows.forEach(r => {
    // null = unknown stage (legacy row with no Process ID stored) — only
    // tagged rows get the benefit of the doubt there, matching prior behavior.
    const knownFinalStage = r.processId ? finalStageIds.has(r.processId.toLowerCase()) : null;
    const isTagged = !!r.productTag;

    if (isTagged) {
      if (r.processId && !knownFinalStage) return; // tagged but its process isn't final-stage
    } else {
      if (!knownFinalStage) return; // untagged: only surface when definitely final-stage output
    }

    // Tagged rows surface under their Product; untagged final-stage rows
    // still surface, under their own Output Item Name, so a completed
    // final-stage lot is always visible to Dispatch even with no tag.
    const key = isTagged ? r.productTag.toLowerCase() : ('__output__' + r.outputItemName.toLowerCase());
    if (!map[key]) {
      map[key] = {
        productId: isTagged ? r.productTag : r.outputItemName,
        productName: isTagged ? (productNameById[key] || r.productTag) : r.outputItemName,
        producedQty: 0,
        dispatchedQty: 0
      };
    }

    map[key].producedQty += r.producedQty;
    map[key].dispatchedQty += r.consumedQty;
  });

  return map;
}

/**
 * Retrieves the "Ready to Dispatch" view: one row per Product ID with a
 * fully-packed, Product-tagged Warehouse Pool credit, with produced/
 * dispatched/ready quantities. Pure read — no persistence.
 */
function getReadyToDispatchData() {
  try {
    const map = _computeReadyToDispatchMap();
    const records = Object.values(map).map(r => ({
      productId: r.productId,
      productName: r.productName,
      producedQty: r.producedQty,
      dispatchedQty: r.dispatchedQty,
      readyQty: r.producedQty - r.dispatchedQty
    }));

    records.sort((a, b) => b.productId.localeCompare(a.productId, undefined, { numeric: true }));

    return buildResponse(true, records);
  } catch (error) {
    Log.error('[getReadyToDispatchData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load Ready to Dispatch data: ' + error.message);
  }
}

/**
 * Retrieves the Dispatched Goods ledger.
 * Includes the physical sheet row index (rowIdx) for edit/delete targeting.
 */
function getDispatchData() {
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.DISPATCH);
    } catch (e) {
      initDispatchSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.DISPATCH);
    }

    ensureDispatchExtraColumns(sheet);
    ensureDispatchLogisticsColumns(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, []);

    const data = sheet.getRange(2, 1, lastRow - 1, DISPATCH_COL.LOGISTICS_COST).getValues();
    const records = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const dispatchNumber = String(row[DISPATCH_COL.DISPATCH_NUMBER - 1] || '').trim();
      if (!dispatchNumber) continue;

      const rawDate = row[DISPATCH_COL.DISPATCH_DATE - 1];
      const dateStr = rawDate instanceof Date ? toSafeDateString(rawDate) : String(rawDate || '');

      records.push({
        rowIdx: i + 2,
        dispatchNumber: dispatchNumber,
        dispatchDate: dateStr,
        dateRaw: rawDate instanceof Date ? rawDate.toISOString() : null,
        orderNumber: String(row[DISPATCH_COL.ORDER_NUMBER - 1] || '').trim(),
        clientName: String(row[DISPATCH_COL.CLIENT_NAME - 1] || '').trim(),
        productId: String(row[DISPATCH_COL.PRODUCT_ID - 1] || '').trim(),
        productName: String(row[DISPATCH_COL.PRODUCT_NAME - 1] || '').trim(),
        qty: Number(row[DISPATCH_COL.QTY - 1]) || 0,
        transport: String(row[DISPATCH_COL.TRANSPORT - 1] || '').trim(),
        remarks: String(row[DISPATCH_COL.REMARKS - 1] || '').trim(),
        invoiceNumber: String(row[DISPATCH_COL.INVOICE_NUMBER - 1] || '').trim(),
        privateMark: String(row[DISPATCH_COL.PRIVATE_MARK - 1] || '').trim(),
        grNumber: String(row[DISPATCH_COL.GR_NUMBER - 1] || '').trim(),
        logisticsContractor: String(row[DISPATCH_COL.LOGISTICS_CONTRACTOR - 1] || '').trim(),
        logisticsRate: Number(row[DISPATCH_COL.LOGISTICS_RATE - 1]) || 0,
        logisticsCost: Number(row[DISPATCH_COL.LOGISTICS_COST - 1]) || 0
      });
    }

    // Sort by date descending, then rowIdx descending (newest first).
    // Timestamp is precomputed once per record instead of re-parsed on every
    // comparison during the sort.
    records.forEach(r => { r._sortTs = r.dateRaw ? new Date(r.dateRaw).getTime() : 0; });
    records.sort((a, b) => {
      if (b._sortTs !== a._sortTs) return b._sortTs - a._sortTs;
      return b.rowIdx - a.rowIdx;
    });
    records.forEach(r => { delete r._sortTs; });

    return buildResponse(true, records);
  } catch (error) {
    Log.error('[getDispatchData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load dispatch data: ' + error.message);
  }
}

/**
 * Sums Qty Ordered across every Client Orders line matching (orderNumber,
 * productId) case-insensitively (normally just one line, but sums in case
 * the same product appears twice on one order). Returns null if no such
 * line exists at all (order number and/or product not found on it), which
 * callers use to distinguish "nothing to check against" from "0 remaining".
 * @private
 */
function _getClientOrderLineQty(orderNumber, productId) {
  let sheet;
  try {
    sheet = getSheet(APP_CONFIG.SHEETS.CLIENT_ORDERS);
  } catch (e) {
    return null;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, CLIENT_ORDERS_COL.PRODUCTION_PUSHED).getValues();
  const targetOrder = orderNumber.toLowerCase();
  const targetProduct = productId.toLowerCase();
  let total = 0;
  let found = false;

  data.forEach(row => {
    const rowOrder = String(row[CLIENT_ORDERS_COL.ORDER_NUMBER - 1] || '').trim().toLowerCase();
    if (rowOrder !== targetOrder) return;
    const rowProduct = String(row[CLIENT_ORDERS_COL.PRODUCT_ID - 1] || '').trim().toLowerCase();
    if (rowProduct !== targetProduct) return;
    total += Number(row[CLIENT_ORDERS_COL.QTY_ORDERED - 1]) || 0;
    found = true;
  });

  return found ? total : null;
}

/**
 * Sums Dispatch Qty across every existing dispatch row matching
 * (orderNumber, productId), optionally excluding one row (the record
 * currently being edited, so it isn't double-counted against itself).
 * @private
 */
function _getDispatchedQtyForOrder(dispatchSheet, orderNumber, productId, excludeRowIdx) {
  const lastRow = dispatchSheet.getLastRow();
  if (lastRow < 2) return 0;

  const data = dispatchSheet.getRange(2, 1, lastRow - 1, DISPATCH_COL.QTY).getValues();
  const targetOrder = orderNumber.toLowerCase();
  const targetProduct = productId.toLowerCase();
  let total = 0;

  data.forEach((row, idx) => {
    const rowNum = idx + 2;
    if (excludeRowIdx && rowNum === excludeRowIdx) return;
    const rowOrder = String(row[DISPATCH_COL.ORDER_NUMBER - 1] || '').trim().toLowerCase();
    if (rowOrder !== targetOrder) return;
    const rowProduct = String(row[DISPATCH_COL.PRODUCT_ID - 1] || '').trim().toLowerCase();
    if (rowProduct !== targetProduct) return;
    total += Number(row[DISPATCH_COL.QTY - 1]) || 0;
  });

  return total;
}

/**
 * Saves a dispatch record (creates a new entry or updates an existing row).
 * Validates that the dispatched quantity does not exceed the currently
 * available "Ready to Dispatch" quantity for the product.
 */
function saveDispatch(formData) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(DISPATCH_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.DISPATCH);
    } catch (e) {
      initDispatchSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.DISPATCH);
    }

    ensureDispatchExtraColumns(sheet);
    ensureDispatchLogisticsColumns(sheet);

    const productId = sanitizeString(formData.productId, 'productId');
    const productName = sanitizeString(formData.productName, 'productName');
    if (!productId || !productName) {
      return buildResponse(false, null, 'Valid Product ID and Product Name are required.');
    }

    const qty = validateNumber(formData.qty, 0.001, 10000000);
    if (qty <= 0) {
      return buildResponse(false, null, 'Dispatch Quantity must be greater than zero.');
    }

    let dateVal = formData.dispatchDate;
    if (!dateVal) dateVal = new Date();
    const dateStr = toSafeDateString(dateVal);

    const orderNumber = sanitizeString(formData.orderNumber || '', 'orderNumber');
    const clientName = sanitizeString(formData.clientName || '', 'clientName');
    const transport = sanitizeString(formData.transport || '', 'transport');
    const remarks = sanitizeString(formData.remarks || '', 'remarks');
    const invoiceNumber = sanitizeString(formData.invoiceNumber || '', 'invoiceNumber');
    const privateMark = sanitizeString(formData.privateMark || '', 'privateMark');
    const grNumber = sanitizeString(formData.grNumber || '', 'grNumber');
    const logisticsContractor = sanitizeString(formData.logisticsContractor || '', 'logisticsContractor');

    const isEdit = !!formData.rowIdx;
    let targetRow = isEdit ? parseInt(formData.rowIdx, 10) : -1;
    let originalQty = 0;
    // Populated below for an edit — Dispatch Number/Product ID/Qty all live
    // within columns 1-7, so one batched read covers all three lookups that
    // otherwise would have been separate getValue() round-trips.
    let existingRow = null;

    if (isEdit) {
      if (isNaN(targetRow) || targetRow < 2 || targetRow > sheet.getLastRow()) {
        return buildResponse(false, null, 'Invalid dispatch record selected for edit.');
      }

      existingRow = sheet.getRange(targetRow, 1, 1, DISPATCH_COL.QTY).getValues()[0];
      const currentProductId = String(existingRow[DISPATCH_COL.PRODUCT_ID - 1]).trim();
      if (currentProductId.toLowerCase() !== productId.toLowerCase()) {
        return buildResponse(false, null, 'Data mismatch: Product ID does not match original row record.');
      }

      originalQty = Number(existingRow[DISPATCH_COL.QTY - 1]) || 0;
    }

    // Validate against currently available Ready to Dispatch quantity
    const readyMap = _computeReadyToDispatchMap();
    const entry = readyMap[productId.toLowerCase()];
    const currentReadyQty = entry ? (entry.producedQty - entry.dispatchedQty) : 0;
    const availableQty = currentReadyQty + originalQty; // add back this record's own qty when editing

    if (qty > availableQty + 0.0001) {
      return buildResponse(false, null, `Only ${availableQty} unit(s) of "${productName}" are Ready to Dispatch.`);
    }

    // Validate against the SPECIFIC PI/Estimate line's own remaining qty, not
    // just aggregate product stock — without this, two different orders for
    // the same product share one pool with no guard, letting one dispatch
    // over-fulfill a small order's line using stock meant for a different
    // order of the same product. Only checked when an Order Number is
    // actually referenced; a "Direct" dispatch (no orderNumber) has no order
    // line to check against.
    let orderLineQty = null;
    if (orderNumber) {
      orderLineQty = _getClientOrderLineQty(orderNumber, productId);
      if (orderLineQty !== null) {
        const alreadyDispatchedForOrder = _getDispatchedQtyForOrder(sheet, orderNumber, productId, isEdit ? targetRow : null);
        const availableForOrder = orderLineQty - alreadyDispatchedForOrder;
        if (qty > availableForOrder + 0.0001) {
          return buildResponse(false, null,
            `Only ${availableForOrder} unit(s) of "${productName}" remain pending on PI/Estimate "${orderNumber}" (ordered ${orderLineQty}, already dispatched ${alreadyDispatchedForOrder} against it). Use a different order reference, or Direct, if this dispatch is really for other stock.`);
        }
      }
    }

    const dispatchNumber = isEdit
      ? String(existingRow[DISPATCH_COL.DISPATCH_NUMBER - 1]).trim()
      : getNextDispatchNumber();

    // Snapshot the logistics contractor's rate at save time, same pattern
    // as Production's contractor payable. Optional — 0 if no rate card entry.
    const logisticsRate = typeof _getContractorRate === 'function'
      ? _getContractorRate(logisticsContractor, LOGISTICS_PROCESS_NAME)
      : 0;
    const logisticsCost = logisticsRate * qty;

    const rowData = [
      dispatchNumber,
      dateStr,
      orderNumber,
      clientName,
      productId,
      productName,
      qty,
      transport,
      remarks,
      invoiceNumber,
      privateMark,
      grNumber,
      logisticsContractor,
      logisticsRate,
      logisticsCost
    ];

    if (isEdit) {
      sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    if (typeof recalculateWarehousePool === 'function') {
      recalculateWarehousePool();
    }

    SpreadsheetApp.flush();

    const logMsg = isEdit
      ? `Dispatch updated: ${dispatchNumber} - ${productName} (${productId}), Qty: ${qty}`
      : `Dispatch recorded: ${dispatchNumber} - ${productName} (${productId}), Qty: ${qty}`;

    logAction(isEdit ? 'UPDATE' : 'CREATE', APP_CONFIG.SHEETS.DISPATCH, dispatchNumber, logMsg, 'SUCCESS');

    let successMsg = isEdit ? 'Dispatch record updated successfully.' : 'Dispatch recorded successfully.';

    // Soft-validate the optional Order Number reference (warn, don't block) —
    // reuses the orderLineQty lookup above instead of re-reading the sheet.
    // orderLineQty === null means no line for THIS product was found under
    // that order number at all (order number itself may still exist with
    // other products on it, or may not exist/may have been removed).
    if (orderNumber && orderLineQty === null) {
      successMsg += ` Note: PI / Estimate "${orderNumber}" has no line for "${productName}" (it may have been edited or removed).`;
    }

    return buildResponse(true, { dispatchNumber: dispatchNumber }, successMsg);
  } catch (error) {
    Log.error('[saveDispatch] Error:', error.message);
    logAction('ERROR', 'saveDispatch', formData.productId || 'NEW', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to save dispatch: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes a dispatch log entry.
 */
function deleteDispatch(rowIdx, expectedDispatchNumber, expectedQty) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(DISPATCH_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.DISPATCH);
    if (!sheet) throw new Error('Dispatch sheet not found.');

    const targetRow = parseInt(rowIdx, 10);
    if (isNaN(targetRow) || targetRow < 2 || targetRow > sheet.getLastRow()) {
      return buildResponse(false, null, 'Invalid dispatch record selected for deletion.');
    }

    const targetRowVals = sheet.getRange(targetRow, 1, 1, DISPATCH_COL.QTY).getValues()[0];
    const dispatchNumber = String(targetRowVals[DISPATCH_COL.DISPATCH_NUMBER - 1]).trim();
    const qty = Number(targetRowVals[DISPATCH_COL.QTY - 1]) || 0;

    if (expectedDispatchNumber !== undefined && expectedQty !== undefined) {
      if (dispatchNumber.toLowerCase() !== String(expectedDispatchNumber || '').trim().toLowerCase() ||
          Math.abs(qty - Number(expectedQty)) > 0.0001) {
        return buildResponse(false, null, 'Data mismatch: The record has been modified or shifted. Please refresh.');
      }
    }

    sheet.deleteRow(targetRow);

    if (typeof recalculateWarehousePool === 'function') {
      recalculateWarehousePool();
    }

    SpreadsheetApp.flush();

    const msg = `Dispatch record "${dispatchNumber}" deleted (Qty was ${qty}).`;
    logAction('DELETE', APP_CONFIG.SHEETS.DISPATCH, dispatchNumber, msg, 'SUCCESS');

    return buildResponse(true, null, 'Dispatch record deleted successfully.');
  } catch (error) {
    Log.error('[deleteDispatch] Error:', error.message);
    logAction('ERROR', 'deleteDispatch', String(rowIdx), error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete dispatch: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes multiple dispatch log entries in a single batch.
 * @param {Array<number|string>} rowIdxs - 1-based sheet row indexes to delete
 * @param {Array<{rowIdx:number|string, expectedDispatchNumber:string, expectedQty:number}>} [expectedRows] -
 *   Optional per-row guard (same mismatch check as the single-row deleteDispatch)
 *   so a row that shifted/changed since the client loaded its list is skipped
 *   instead of silently deleting whatever now sits at that row number.
 */
function deleteDispatchBulk(rowIdxs, expectedRows) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(DISPATCH_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.DISPATCH);
    if (!sheet) throw new Error('Dispatch sheet not found.');

    const lastRow = sheet.getLastRow();
    let targetRows = (rowIdxs || [])
      .map(r => parseInt(r, 10))
      .filter(r => !isNaN(r) && r >= 2 && r <= lastRow);

    if (targetRows.length === 0) {
      return buildResponse(true, null, 'No dispatch records selected.');
    }

    const expectedByRow = {};
    (expectedRows || []).forEach(e => {
      if (e && e.rowIdx !== undefined) expectedByRow[String(e.rowIdx)] = e;
    });

    // Single batched read (Dispatch Number through Qty) covering every row
    // in range, instead of one getRange().getValue() round-trip per row.
    const idQtyRows = sheet.getRange(2, 1, lastRow - 1, DISPATCH_COL.QTY).getValues();

    let skippedMismatch = 0;
    targetRows = targetRows.filter(rowNum => {
      const expected = expectedByRow[String(rowNum)];
      if (!expected || expected.expectedDispatchNumber === undefined || expected.expectedQty === undefined) {
        return true;
      }
      const rowVals = idQtyRows[rowNum - 2];
      const dispatchNumber = String(rowVals[DISPATCH_COL.DISPATCH_NUMBER - 1]).trim();
      const qty = Number(rowVals[DISPATCH_COL.QTY - 1]) || 0;
      if (dispatchNumber.toLowerCase() !== String(expected.expectedDispatchNumber || '').trim().toLowerCase() ||
          Math.abs(qty - Number(expected.expectedQty)) > 0.0001) {
        skippedMismatch++;
        return false;
      }
      return true;
    });

    if (targetRows.length === 0) {
      return buildResponse(false, null, 'Data mismatch: the selected record(s) have been modified or shifted. Please refresh.');
    }

    const targetRowSet = new Set(targetRows);
    const { rowsDeleted } = rewriteSheetExcludingRows(sheet, 2, (_row, rowNum) => targetRowSet.has(rowNum));

    if (typeof recalculateWarehousePool === 'function') {
      recalculateWarehousePool();
    }

    SpreadsheetApp.flush();

    const msg = skippedMismatch > 0
      ? `Deleted ${rowsDeleted} dispatch record(s). Skipped ${skippedMismatch} that were modified or shifted — please refresh and retry those.`
      : `Deleted ${rowsDeleted} dispatch record(s).`;
    logAction('BULK_DELETE', APP_CONFIG.SHEETS.DISPATCH, 'multiple', msg, 'SUCCESS');

    return buildResponse(true, null, msg);
  } catch (error) {
    Log.error('[deleteDispatchBulk] Error:', error.message);
    logAction('ERROR', 'deleteDispatchBulk', 'multiple', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete dispatch records: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}
