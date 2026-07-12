/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_warehouse.gs — WAREHOUSE POOL MODULE
 *
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Tracks ready stock of intermediate and finished process outputs (e.g.
 * "Painted Frame", "Fitted Rim", "Fitted Frame", "Packed Bicycle") that sit
 * between production stages. Credited when a Production lot for the
 * originating process is marked Completed; debited when a downstream
 * Production lot consumes it as a POOL-sourced component, or when a
 * Product-tagged final-stage credit is Dispatched.
 *
 * Sheet Layout (Warehouse Pool):
 * ───────────────────────────────────────────────────────────────────────────
 * Col A (1):   Output Item Name
 * Col B (2):   Process ID (originating process)
 * Col C (3):   Product Tag (blank for intermediate WIP; Product ID for
 *              tagged final-stage finished goods)
 * Col D (4):   Produced Qty
 * Col E (5):   Consumed Qty
 * Col F (6):   Available Qty (= Produced Qty - Consumed Qty)
 * Col G (7):   Color (Color Master name this bucket was credited under;
 *              blank for color-agnostic/uncolored output — a multi-color
 *              lot's output is split into one bucket per color so a
 *              downstream process can draw a specific color out of the pool)
 * ═══════════════════════════════════════════════════════════════════════════
 */

const WAREHOUSE_LOCK_TIMEOUT_MS = 15000;

/**
 * Initializes the Warehouse Pool sheet with correct headers.
 */
function initWarehousePoolSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
    if (!sheet) {
      sheet = ss.insertSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
    }

    const headers = ['Output Item Name', 'Process ID', 'Product Tag', 'Produced Qty', 'Consumed Qty', 'Available Qty', 'Color'];

    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Warehouse Pool sheet initialized successfully.');
  } catch (error) {
    console.error('[initWarehousePoolSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize Warehouse Pool sheet: ' + error.message);
  }
}

/**
 * Backfills the "Color" column on Warehouse Pool sheets created before
 * per-color pool tracking existed, so legacy sheets don't throw when
 * read/written.
 */
function ensureWarehousePoolColorColumn(sheet) {
  try {
    if (sheet.getLastColumn() < WAREHOUSE_POOL_COL.COLOR) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), WAREHOUSE_POOL_COL.COLOR - sheet.getLastColumn());
      sheet.getRange(1, WAREHOUSE_POOL_COL.COLOR, 1, 1)
        .setValues([['Color']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    console.error('[ensureWarehousePoolColorColumn] Error:', error.message);
  }
}

/**
 * @private
 * Builds the key used to bucket a pool row: Output Item Name + Product Tag
 * (blank tag for intermediate WIP) + Color (blank for color-agnostic output).
 */
function _poolKey(outputItemName, productTag, color) {
  return String(outputItemName || '').trim().toLowerCase() + '||' + String(productTag || '').trim().toLowerCase() + '||' + String(color || '').trim().toLowerCase();
}

// ── Warehouse Pool Opening Balances ─────────────────────────────────────
// Manual seed entries for stock that already existed before this app went
// live (or a one-off correction). Unlike the Warehouse Pool sheet itself —
// fully wiped and rebuilt by recalculateWarehousePool() on every call —
// this sheet is its own persistent source of truth, read once per rebuild
// and folded into the matching bucket's Produced Qty. Shares
// WAREHOUSE_LOCK_TIMEOUT_MS with the pool sheet itself since both take the
// same document lock.

/**
 * Initializes the Warehouse Pool Opening sheet with correct headers.
 */
function initWarehousePoolOpeningSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
    if (!sheet) {
      sheet = ss.insertSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
    }

    const headers = ['Output Item Name', 'Process ID', 'Product Tag', 'Color', 'Qty', 'Date', 'Remarks'];

    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Warehouse Pool Opening sheet initialized successfully.');
  } catch (error) {
    console.error('[initWarehousePoolOpeningSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize Warehouse Pool Opening sheet: ' + error.message);
  }
}

/**
 * Lists every opening-balance entry, with the originating Process Name
 * resolved for display (the sheet itself only stores Process ID).
 */
function getWarehousePoolOpeningData() {
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
    } catch (e) {
      initWarehousePoolOpeningSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, []);

    const allProcesses = typeof _getAllProcessRecords === 'function' ? _getAllProcessRecords() : [];
    const processNameById = {};
    allProcesses.forEach(p => { processNameById[p.processId.toLowerCase()] = p.processName; });

    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const records = data
      .map((row, i) => {
        const processId = String(row[WAREHOUSE_POOL_OPENING_COL.PROCESS_ID - 1] || '').trim();
        const rawDate = row[WAREHOUSE_POOL_OPENING_COL.DATE - 1];
        return {
          rowIdx: i + 2,
          outputItemName: String(row[WAREHOUSE_POOL_OPENING_COL.OUTPUT_ITEM_NAME - 1] || '').trim(),
          processId: processId,
          processName: processNameById[processId.toLowerCase()] || processId,
          productTag: String(row[WAREHOUSE_POOL_OPENING_COL.PRODUCT_TAG - 1] || '').trim(),
          color: String(row[WAREHOUSE_POOL_OPENING_COL.COLOR - 1] || '').trim(),
          qty: Number(row[WAREHOUSE_POOL_OPENING_COL.QTY - 1]) || 0,
          date: rawDate instanceof Date ? toSafeDateString(rawDate) : String(rawDate || ''),
          dateRaw: rawDate instanceof Date ? rawDate.toISOString() : null,
          remarks: String(row[WAREHOUSE_POOL_OPENING_COL.REMARKS - 1] || '').trim()
        };
      })
      .filter(r => r.outputItemName);

    records.sort((a, b) => {
      const dateA = a.dateRaw ? new Date(a.dateRaw) : new Date(0);
      const dateB = b.dateRaw ? new Date(b.dateRaw) : new Date(0);
      if (dateB - dateA !== 0) return dateB - dateA;
      return b.rowIdx - a.rowIdx;
    });

    return buildResponse(true, records);
  } catch (error) {
    console.error('[getWarehousePoolOpeningData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load Warehouse Pool opening balances: ' + error.message);
  }
}

/**
 * Records a new opening-balance entry for a Warehouse Pool bucket. The
 * Output Item Name is always derived from the selected Process (never
 * free-typed), so the seeded bucket always lines up with a real process's
 * output. Product Tag and Color are optional, same as a Production lot.
 * Negative quantities are allowed (zero is rejected as a no-op) since this
 * form doubles as a one-off correction — e.g. clawing back a bucket that
 * went negative from over-issued production.
 */
function saveWarehousePoolOpening(formData) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(WAREHOUSE_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
    } catch (e) {
      initWarehousePoolOpeningSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
    }

    const processId = sanitizeString(formData.processId || '', 'processId');
    if (!processId) {
      return buildResponse(false, null, 'A Process is required.');
    }

    const allProcesses = typeof _getAllProcessRecords === 'function' ? _getAllProcessRecords() : [];
    const process = allProcesses.find(p => p.processId.toLowerCase() === processId.toLowerCase());
    if (!process) {
      return buildResponse(false, null, 'Selected Process was not found.');
    }
    if (!process.outputItemName) {
      return buildResponse(false, null, `Process "${process.processName}" has no Output Item Name configured.`);
    }

    const productTag = process.isFinalStage ? sanitizeString(formData.productTag || '', 'productTag') : '';
    const color = sanitizeString(formData.color || '', 'color');

    const qty = validateNumber(formData.qty, -10000000, 10000000);
    if (qty === 0) {
      return buildResponse(false, null, 'Opening Quantity cannot be zero.');
    }

    const dateObj = toSafeDateObject(formData.date) || new Date();

    const remarks = sanitizeString(formData.remarks || '', 'remarks');

    sheet.appendRow([process.outputItemName, process.processId, productTag, color, qty, dateObj, remarks]);

    recalculateWarehousePool();
    SpreadsheetApp.flush();

    const logMsg = `Opening stock added: ${qty} x "${process.outputItemName}"${color ? `, Color: ${color}` : ''}${productTag ? ` (tagged: ${productTag})` : ''}.`;
    logAction('CREATE', APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING, process.processId, logMsg, 'SUCCESS');

    return buildResponse(true, null, 'Opening stock recorded successfully.');
  } catch (error) {
    console.error('[saveWarehousePoolOpening] Error:', error.message);
    logAction('ERROR', 'saveWarehousePoolOpening', formData.processId || 'NEW', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to record opening stock: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * @private
 * Reads the live Warehouse Pool sheet for the exact bucket matching
 * outputItemName + productTag + color, returning its current Available Qty
 * (0 if the bucket doesn't exist yet — e.g. a brand-new opening entry).
 */
function _getWarehousePoolBucketAvailableQty(outputItemName, productTag, color) {
  let sheet;
  try {
    sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
  } catch (e) {
    return 0;
  }
  ensureWarehousePoolColorColumn(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const targetKey = _poolKey(outputItemName, productTag, color);
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowKey = _poolKey(
      row[WAREHOUSE_POOL_COL.OUTPUT_ITEM_NAME - 1],
      row[WAREHOUSE_POOL_COL.PRODUCT_TAG - 1],
      row[WAREHOUSE_POOL_COL.COLOR - 1]
    );
    if (rowKey === targetKey) {
      return Number(row[WAREHOUSE_POOL_COL.AVAILABLE_QTY - 1]) || 0;
    }
  }
  return 0;
}

/**
 * Manually corrects a Warehouse Pool bucket's Available Qty (e.g. to seed
 * opening stock for a bucket that already has Production history, or to fix
 * a physical-recount discrepancy), by appending a delta row to the Warehouse
 * Pool Opening sheet — the same durable source recalculateWarehousePool()
 * already folds into Produced Qty — and rebuilding. A reason is required and
 * the adjustment is recorded in the audit Logs sheet so corrections stay
 * traceable, mirroring adjustStockManually() for raw-material Stock. Negative
 * values are allowed — over-issued production can legitimately leave a
 * bucket negative until the user reviews and corrects it.
 * @param {string} outputItemName
 * @param {string} processId - Originating Process (required so the
 *   correction row lines up with a real process's output, same as the
 *   opening-stock form).
 * @param {string} productTag
 * @param {string} color
 * @param {number} newAvailableQty - The corrected Available Qty value.
 * @param {string} reason - Why the adjustment was made (required).
 */
function adjustWarehousePoolManually(outputItemName, processId, productTag, color, newAvailableQty, reason) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(WAREHOUSE_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const itemName = sanitizeString(outputItemName || '', 'outputItemName');
    if (!itemName) {
      return buildResponse(false, null, 'Output Item Name is required.');
    }

    const procId = sanitizeString(processId || '', 'processId');
    if (!procId) {
      return buildResponse(false, null, 'Originating Process is required.');
    }

    const allProcesses = typeof _getAllProcessRecords === 'function' ? _getAllProcessRecords() : [];
    if (!allProcesses.some(p => p.processId.toLowerCase() === procId.toLowerCase())) {
      return buildResponse(false, null, 'Originating Process was not found.');
    }

    const newQty = Number(newAvailableQty);
    if (isNaN(newQty)) {
      return buildResponse(false, null, 'Corrected quantity must be a valid number.');
    }

    const reasonText = sanitizeString(reason || '', 'reason');
    if (!reasonText) {
      return buildResponse(false, null, 'A reason is required for manual stock corrections.');
    }

    const tag = sanitizeString(productTag || '', 'productTag');
    const colorVal = sanitizeString(color || '', 'color');

    const oldQty = _getWarehousePoolBucketAvailableQty(itemName, tag, colorVal);
    const delta = newQty - oldQty;
    if (delta === 0) {
      // oldQty here is read fresh from the sheet, so it's the authoritative
      // current value even if the caller's own on-screen number (e.g. a
      // stale client-side cache after another user's change) was different.
      // Surfacing it lets the UI reconcile its display instead of being
      // stuck showing a stale value that will keep rejecting this same edit.
      // buildResponse() forces data to null on failure, so this bypasses it
      // and returns the {success,data,message} shape directly.
      return { success: false, data: { oldAvailableQty: oldQty, newAvailableQty: oldQty }, message: 'New quantity is the same as the current value — nothing to adjust.' };
    }

    let openingSheet;
    try {
      openingSheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
    } catch (e) {
      initWarehousePoolOpeningSheet();
      openingSheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
    }

    openingSheet.appendRow([itemName, procId, tag, colorVal, delta, new Date(), `Correction: ${reasonText}`]);

    recalculateWarehousePool();
    SpreadsheetApp.flush();

    // Each field is URI-encoded before joining so a literal "|" inside an
    // item name, tag, or color can't be mistaken for the delimiter and
    // corrupt the 3-way split in getWarehousePoolAdjustmentHistory().
    const recordId = [itemName, tag, colorVal].map(encodeURIComponent).join('|');
    const details = `Old: ${oldQty}, New: ${newQty}. Reason: ${reasonText}`;
    logAction('ADJUST', APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING, recordId, details, 'SUCCESS');

    return buildResponse(true, { oldAvailableQty: oldQty, newAvailableQty: newQty }, 'Warehouse Pool stock adjusted successfully.');
  } catch (error) {
    console.error('[adjustWarehousePoolManually] Error:', error.message);
    logAction('ERROR', 'adjustWarehousePoolManually', String(outputItemName), error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to adjust Warehouse Pool stock: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Returns the history of manual Warehouse Pool corrections ('ADJUST'),
 * sourced from the audit Logs sheet (see adjustWarehousePoolManually). Used
 * by the Warehouse Pool Ledger view to surface these alongside Production
 * credits/debits and Dispatch debits.
 */
function getWarehousePoolAdjustmentHistory() {
  try {
    const logsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSheetName('LOGS'));
    if (!logsSheet) return buildResponse(true, []);

    const lastRow = logsSheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, []);

    const data = logsSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const records = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const action = String(row[LOGS_COL.ACTION - 1] || '').trim().toUpperCase();
      const sheetName = String(row[LOGS_COL.SHEET - 1] || '').trim();
      if (action !== 'ADJUST' || sheetName !== APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING) continue;

      const recordId = String(row[LOGS_COL.RECORD_ID - 1] || '');
      const idParts = recordId.split('|');
      if (idParts.length !== 3) continue;

      // recordId parts are URI-encoded (see adjustWarehousePoolManually) so a
      // literal "|" inside a value can't be mistaken for the delimiter.
      // Older log rows predate the encoding and decode back to themselves
      // unchanged as long as they contain no "%" sequences.
      let decodedParts;
      try {
        decodedParts = idParts.map(decodeURIComponent);
      } catch (e) {
        decodedParts = idParts;
      }

      const details = String(row[LOGS_COL.DETAILS - 1] || '');
      const detailsMatch = details.match(/Old:\s*([\d.\-]+),\s*New:\s*([\d.\-]+)\.\s*Reason:\s*(.*)$/);
      if (!detailsMatch) continue;

      records.push({
        date: row[LOGS_COL.TIMESTAMP - 1],
        outputItemName: decodedParts[0],
        productTag: decodedParts[1],
        color: decodedParts[2],
        oldValue: Number(detailsMatch[1]),
        newValue: Number(detailsMatch[2]),
        reason: detailsMatch[3],
        user: row[LOGS_COL.USER - 1]
      });
    }

    return buildResponse(true, records);
  } catch (error) {
    console.error('[getWarehousePoolAdjustmentHistory] Error:', error.message);
    return buildResponse(false, null, 'Failed to load Warehouse Pool adjustment history: ' + error.message);
  }
}

/**
 * Deletes an opening-balance entry and rebuilds the pool.
 */
function deleteWarehousePoolOpening(rowIdx) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(WAREHOUSE_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
    if (!sheet) throw new Error('Warehouse Pool Opening sheet not found.');

    const targetRow = parseInt(rowIdx, 10);
    if (isNaN(targetRow) || targetRow < 2 || targetRow > sheet.getLastRow()) {
      return buildResponse(false, null, 'Invalid opening stock entry selected for deletion.');
    }

    const outputItemName = String(sheet.getRange(targetRow, WAREHOUSE_POOL_OPENING_COL.OUTPUT_ITEM_NAME).getValue()).trim();

    sheet.deleteRow(targetRow);
    recalculateWarehousePool();
    SpreadsheetApp.flush();

    logAction('DELETE', APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING, outputItemName, `Opening stock entry for "${outputItemName}" deleted.`, 'SUCCESS');

    return buildResponse(true, null, 'Opening stock entry deleted successfully.');
  } catch (error) {
    console.error('[deleteWarehousePoolOpening] Error:', error.message);
    logAction('ERROR', 'deleteWarehousePoolOpening', String(rowIdx), error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete opening stock entry: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * @private
 * Reads the Warehouse Pool Opening sheet, returning each row in the same
 * shape recalculateWarehousePool()'s getBucket() expects.
 */
function _getWarehousePoolOpeningRows() {
  const rows = [];
  let sheet;
  try {
    sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
  } catch (e) {
    return rows; // Opening sheet not initialized yet — nothing to seed
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return rows;

  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  data.forEach(row => {
    const outputItemName = String(row[WAREHOUSE_POOL_OPENING_COL.OUTPUT_ITEM_NAME - 1] || '').trim();
    if (!outputItemName) return;
    // Zero is skipped (no-op); negative IS kept — a downward manual
    // correction (see adjustWarehousePoolManually) appends a negative delta
    // row here, and it must survive this read or the correction silently
    // reverts on the next rebuild.
    const qty = Number(row[WAREHOUSE_POOL_OPENING_COL.QTY - 1]) || 0;
    if (qty === 0) return;

    rows.push({
      outputItemName: outputItemName,
      processId: String(row[WAREHOUSE_POOL_OPENING_COL.PROCESS_ID - 1] || '').trim(),
      productTag: String(row[WAREHOUSE_POOL_OPENING_COL.PRODUCT_TAG - 1] || '').trim(),
      color: String(row[WAREHOUSE_POOL_OPENING_COL.COLOR - 1] || '').trim(),
      qty: qty
    });
  });

  return rows;
}

/**
 * Recalculates the entire Warehouse Pool sheet from scratch based on
 * manually-recorded Opening Balances + Completed Production lots (credits)
 * and downstream POOL-sourced component consumption + finished-stage
 * Dispatch consumption (debits). Mirrors recalculateStock()'s "always
 * rebuild from source data" approach.
 */
function recalculateWarehousePool() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(WAREHOUSE_LOCK_TIMEOUT_MS)) {
    console.error('[recalculateWarehousePool] Could not acquire lock.');
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    let poolSheet;
    try {
      poolSheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
    } catch (e) {
      initWarehousePoolSheet();
      poolSheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
    }
    ensureWarehousePoolColorColumn(poolSheet);

    let prodSheet;
    try {
      prodSheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
    } catch (e) {
      prodSheet = null;
    }

    const buckets = {}; // key -> { outputItemName, processId, productTag, color, producedQty, consumedQty }

    function getBucket(outputItemName, processId, productTag, color) {
      const key = _poolKey(outputItemName, productTag, color);
      if (!buckets[key]) {
        buckets[key] = {
          outputItemName: String(outputItemName || '').trim(),
          processId: String(processId || '').trim(),
          productTag: String(productTag || '').trim(),
          color: String(color || '').trim(),
          producedQty: 0,
          consumedQty: 0
        };
      }
      return buckets[key];
    }

    // Pass 0: seed buckets from manually-recorded Opening Balances (stock
    // that already existed before this app went live). This sheet is never
    // wiped by this rebuild, so it's the one durable source of "stock that
    // didn't come from a Production lot".
    _getWarehousePoolOpeningRows().forEach(r => {
      getBucket(r.outputItemName, r.processId, r.productTag, r.color).producedQty += r.qty;
    });

    if (prodSheet) {
      ensureProductionExtraColumns(prodSheet);
      ensureProductionProcessColumns(prodSheet);
      ensureProductionContractorColumns(prodSheet);
      ensureProductionWarehouseColumns(prodSheet);
      ensureProductionColorColumn(prodSheet);
      ensureProductionColorBreakdownColumn(prodSheet);

      const lastRow = prodSheet.getLastRow();
      if (lastRow >= 2) {
        const numCols = PRODUCTION_COL.COLOR_BREAKDOWN;
        const data = prodSheet.getRange(2, 1, lastRow - 1, numCols).getValues();

        // Pass 1: credit every Completed lot's own output to its pool
        // bucket(s). A multi-color lot's Color Breakdown is split into one
        // bucket per color so a downstream process can later draw a
        // specific color back out; a color-agnostic lot credits the single
        // blank-color bucket as before.
        data.forEach(row => {
          const status = String(row[PRODUCTION_COL.STATUS - 1] || '').trim().toLowerCase();
          if (status !== 'completed') return;

          const processId = String(row[PRODUCTION_COL.PROCESS_ID - 1] || '').trim();
          const outputItemName = String(row[PRODUCTION_COL.OUTPUT_ITEM_NAME - 1] || '').trim();
          if (!outputItemName) return;

          const productTag = String(row[PRODUCTION_COL.PRODUCT_ID - 1] || '').trim();

          let colorBreakdown = [];
          const colorBreakdownRaw = String(row[PRODUCTION_COL.COLOR_BREAKDOWN - 1] || '').trim();
          if (colorBreakdownRaw) {
            try {
              const parsed = JSON.parse(colorBreakdownRaw);
              if (Array.isArray(parsed)) colorBreakdown = parsed;
            } catch (e) { /* ignore malformed data */ }
          }

          if (colorBreakdown.length > 0) {
            colorBreakdown.forEach(entry => {
              const color = String(entry.color || '').trim();
              const qty = Number(entry.qty) || 0;
              if (!color || qty <= 0) return;
              getBucket(outputItemName, processId, productTag, color).producedQty += qty;
            });
          } else {
            const qty = Number(row[PRODUCTION_COL.QTY - 1]) || 0;
            getBucket(outputItemName, processId, productTag, '').producedQty += qty;
          }
        });

        // Pass 2: debit POOL-sourced components consumed by Completed lots
        // from the (untagged, intermediate) bucket of the upstream item. A
        // component scoped to a specific color (colorGroup other than
        // COMMON) debits that color's bucket specifically; a COMMON
        // component debits the blank-color bucket (the right bucket for an
        // upstream process that isn't itself multi-color).
        data.forEach(row => {
          const status = String(row[PRODUCTION_COL.STATUS - 1] || '').trim().toLowerCase();
          if (status !== 'completed') return;

          const rawComponents = String(row[PRODUCTION_COL.COMPONENTS_CONSUMED - 1] || '').trim();
          if (!rawComponents) return;

          let components = [];
          try {
            const parsed = JSON.parse(rawComponents);
            if (Array.isArray(parsed)) components = parsed;
          } catch (e) {
            return;
          }

          components.forEach(comp => {
            const sourceType = String(comp.sourceType || '').trim().toUpperCase();
            if (sourceType !== COMPONENT_SOURCE_TYPES.POOL) return;
            const itemName = String(comp.itemName || '').trim();
            if (!itemName) return;
            const qty = Number(comp.qty) || 0;

            const colorGroup = String(comp.colorGroup || '').trim();
            const color = colorGroup && colorGroup.toUpperCase() !== COMPONENT_COLOR_GROUP_COMMON ? colorGroup : '';

            const bucket = getBucket(itemName, '', '', color);
            bucket.consumedQty += qty;
          });
        });
      }
    }

    // Pass 3: debit finished-goods buckets by Dispatch quantity. A Product-
    // tagged bucket is matched by its tag; an untagged final-stage bucket
    // has no tag to match, so Dispatch's "Product ID" for that lot is the
    // Output Item Name itself (see _computeReadyToDispatchMap in
    // module_dispatch.js) — fall back to matching on that, restricted to
    // final-stage buckets so an untagged intermediate-WIP bucket sharing the
    // same Output Item Name from a non-final process is never touched.
    let dispatchSheet;
    try {
      dispatchSheet = getSheet(APP_CONFIG.SHEETS.DISPATCH);
    } catch (e) {
      dispatchSheet = null;
    }

    if (dispatchSheet) {
      const dLastRow = dispatchSheet.getLastRow();
      if (dLastRow >= 2) {
        const dData = dispatchSheet.getRange(2, 1, dLastRow - 1, DISPATCH_COL.QTY).getValues();
        const dispatchQtyByKey = {}; // productId(tag or outputItemName)Lower -> total dispatched qty
        dData.forEach(row => {
          const productId = String(row[DISPATCH_COL.PRODUCT_ID - 1] || '').trim();
          if (!productId) return;
          const qty = Number(row[DISPATCH_COL.QTY - 1]) || 0;
          const key = productId.toLowerCase();
          dispatchQtyByKey[key] = (dispatchQtyByKey[key] || 0) + qty;
        });

        const finalStageIds = new Set(
          (typeof _getAllProcessRecords === 'function' ? _getAllProcessRecords() : [])
            .filter(p => p.isFinalStage)
            .map(p => p.processId.toLowerCase())
        );

        // Dispatch carries no color of its own, so a Product Tag (or
        // untagged Output Item Name) credited across multiple color buckets
        // (a multi-color final-stage lot) can't be debited by color —
        // greedily drain whichever color buckets have stock first, dumping
        // any leftover (over-dispatch beyond total availability) on the
        // first bucket so the total consumedQty across all matching buckets
        // still equals the total dispatched qty.
        Object.keys(dispatchQtyByKey).forEach(key => {
          let remaining = dispatchQtyByKey[key];
          let matchingBuckets = Object.values(buckets).filter(b => b.productTag && b.productTag.toLowerCase() === key);
          if (matchingBuckets.length === 0) {
            matchingBuckets = Object.values(buckets).filter(b =>
              !b.productTag &&
              b.outputItemName.toLowerCase() === key &&
              b.processId && finalStageIds.has(b.processId.toLowerCase())
            );
          }
          if (matchingBuckets.length === 0) return;

          matchingBuckets.forEach(bucket => {
            if (remaining <= 0) return;
            const available = Math.max(bucket.producedQty - bucket.consumedQty, 0);
            const take = Math.min(remaining, available);
            bucket.consumedQty += take;
            remaining -= take;
          });

          if (remaining > 0) {
            matchingBuckets[0].consumedQty += remaining;
          }
        });
      }
    }

    const rows = Object.values(buckets)
      .filter(b => b.outputItemName)
      .map(b => [b.outputItemName, b.processId, b.productTag, b.producedQty, b.consumedQty, b.producedQty - b.consumedQty, b.color]);

    // Rewrite the sheet from scratch (small dataset — process count is tiny).
    const lastRow = poolSheet.getLastRow();
    if (lastRow >= 2) {
      poolSheet.getRange(2, 1, lastRow - 1, 7).clearContent();
    }
    if (rows.length > 0) {
      poolSheet.getRange(2, 1, rows.length, 7).setValues(rows);
    }

    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Warehouse Pool recalculated.');
  } catch (error) {
    console.error('[recalculateWarehousePool] Error:', error.message);
    logAction('ERROR', 'recalculateWarehousePool', 'N/A', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to recalculate Warehouse Pool: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Public read endpoint for the Warehouse Pool view.
 */
function getWarehousePoolData() {
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
    } catch (e) {
      initWarehousePoolSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
    }
    ensureWarehousePoolColorColumn(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, []);

    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const records = data
      .map((row, i) => ({
        rowIdx: i + 2,
        outputItemName: String(row[WAREHOUSE_POOL_COL.OUTPUT_ITEM_NAME - 1] || '').trim(),
        processId: String(row[WAREHOUSE_POOL_COL.PROCESS_ID - 1] || '').trim(),
        productTag: String(row[WAREHOUSE_POOL_COL.PRODUCT_TAG - 1] || '').trim(),
        producedQty: Number(row[WAREHOUSE_POOL_COL.PRODUCED_QTY - 1]) || 0,
        consumedQty: Number(row[WAREHOUSE_POOL_COL.CONSUMED_QTY - 1]) || 0,
        availableQty: Number(row[WAREHOUSE_POOL_COL.AVAILABLE_QTY - 1]) || 0,
        color: String(row[WAREHOUSE_POOL_COL.COLOR - 1] || '').trim()
      }))
      .filter(r => r.outputItemName);

    records.sort((a, b) => a.outputItemName.localeCompare(b.outputItemName) || a.color.localeCompare(b.color));

    return buildResponse(true, records);
  } catch (error) {
    console.error('[getWarehousePoolData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load Warehouse Pool data: ' + error.message);
  }
}

/**
 * Returns a lowercased-item-name -> availability map for every untagged
 * (intermediate WIP) Warehouse Pool bucket, built from a single batch read
 * of the sheet. Use this instead of calling getPoolAvailableQty() in a loop
 * — each call to that function re-reads the whole sheet, so looking up N
 * items one at a time costs N full-sheet reads instead of 1.
 *
 * Each entry's `total` sums every color bucket for that item (the qty a
 * COMMON-scoped component may draw, since it doesn't care which color);
 * `byColor[colorLower]` is that one color's own bucket (the qty a
 * component scoped to a specific Color Master name may draw), and
 * `byColor['']` is the blank/color-agnostic bucket.
 * @returns {Object} { [itemNameLower]: { total: number, byColor: { [colorLower]: number } } }
 */
function getPoolAvailableQtyMap() {
  const map = {};
  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
    ensureWarehousePoolColorColumn(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return map;

    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    data.forEach(row => {
      const itemName = String(row[WAREHOUSE_POOL_COL.OUTPUT_ITEM_NAME - 1] || '').trim().toLowerCase();
      const productTag = String(row[WAREHOUSE_POOL_COL.PRODUCT_TAG - 1] || '').trim();
      if (!itemName || productTag) return;
      const color = String(row[WAREHOUSE_POOL_COL.COLOR - 1] || '').trim().toLowerCase();
      const availableQty = Number(row[WAREHOUSE_POOL_COL.AVAILABLE_QTY - 1]) || 0;

      if (!map[itemName]) map[itemName] = { total: 0, byColor: {} };
      map[itemName].total += availableQty;
      map[itemName].byColor[color] = (map[itemName].byColor[color] || 0) + availableQty;
    });
    return map;
  } catch (e) {
    return map;
  }
}

/**
 * Returns the current total available qty (summed across every color
 * bucket) for an untagged (intermediate WIP) Warehouse Pool item.
 *
 * Looking up a single item is fine, but if you need availability for
 * several items (e.g. every POOL-sourced component in a recipe), call
 * getPoolAvailableQtyMap() once instead — each call here is a full-sheet
 * read.
 * @param {string} outputItemName
 * @returns {number}
 */
function getPoolAvailableQty(outputItemName) {
  const target = String(outputItemName || '').trim().toLowerCase();
  if (!target) return 0;
  const entry = getPoolAvailableQtyMap()[target];
  return entry ? entry.total : 0;
}
