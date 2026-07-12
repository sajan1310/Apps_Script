/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_stock.gs — STOCK MANAGEMENT MODULE
 * 
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Provides backend database operations for the Stock tab, including:
 * - Initialization of the Stock sheet (initStockSheet)
 * - Fetching stock inventory (getStockData)
 * - Inline editing of minimum threshold values (updateThreshold)
 * - Importing initial stock from Excel/CSV (importStockData)
 * - Real-time calculation of current stock (recalculateStock)
 * 
 * Stock Calculation Formula:
 * ───────────────────────────────────────────────────────────────────────────
 * Current Stock = Initial Stock + Billed Qty (from Bill Ledger)
 *                 - Returned Qty (from Return Ledger — goods sent back to
 *                   a vendor; see module_return.js)
 *                 - Wasted Qty (from Wastage Log; see module_wastage.js)
 *                 - Issued Qty (ad-hoc issuance outside any Process BOM /
 *                   Production lot; see module_issue.js)
 *                 - Consumed Qty (ITEM-sourced components on Completed
 *                   Production lots — see module_warehouse.js for POOL-
 *                   sourced consumption, which debits the Warehouse Pool
 *                   instead of raw-material Stock)
 * 
 * ═══════════════════════════════════════════════════════════════════════════
 */

const STOCK_LOCK_TIMEOUT_MS = 15000;

/**
 * Initializes the Stock sheet if it does not exist.
 * Sets the default headers and formats the sheet.
 */
function initStockSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.STOCK);
    if (!sheet) {
      sheet = ss.insertSheet(APP_CONFIG.SHEETS.STOCK);
    }
    
    const headers = [
      'Item Name',
      'Size',
      'Initial Stock',
      'Current Stock',
      'Threshold',
      'Dead Stock'
    ];
    
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
      
    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Stock sheet initialized successfully.');
  } catch (error) {
    Log.error('[initStockSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize Stock sheet: ' + error.message);
  }
}

/**
 * Builds the all-time Billed Qty and Consumed Qty maps (Item Name | Size -> Sum Qty)
 * used to derive Current Stock from Initial Stock. Shared by recalculateStock()
 * (to refresh every row) and importStockData() (to work out, ahead of time, what
 * Initial Stock must be set to so a freshly-imported quantity becomes the new
 * ground-truth Current Stock).
 * @returns {{billQtyMap: Object, consumedQtyMap: Object}}
 */
function _getBilledAndConsumedQtyMaps(ss) {
  // 1. Build Bill Qty Map (Item Name | Size -> Sum Qty)
  const billQtyMap = {};
  const billSheet = ss.getSheetByName(APP_CONFIG.SHEETS.BILL);
  if (billSheet) {
    const billLastRow = billSheet.getLastRow();
    if (billLastRow >= 2) {
      // Read columns PO_NUMBER (1) through BASE_RATE (16)
      const numCols = Math.max(billSheet.getLastColumn(), BILL_COL.BASE_RATE);
      const billData = billSheet.getRange(2, 1, billLastRow - 1, numCols).getValues();
      for (let i = 0; i < billData.length; i++) {
        const row = billData[i];
        const itemName = String(row[BILL_COL.ITEM_NAME - 1] || '').trim().toLowerCase();
        const size = String(row[BILL_COL.SIZE - 1] || '').trim().toLowerCase();
        // Stock is tracked in each item's Base Unit — use the converted
        // BASE_QTY. Legacy rows predating that column are implicitly Pcs,
        // so fall back to the as-entered QTY (factor 1, same number).
        const rawBaseQty = row[BILL_COL.BASE_QTY - 1];
        const qty = (rawBaseQty !== '' && rawBaseQty !== null && rawBaseQty !== undefined)
          ? Number(rawBaseQty) || 0
          : Number(row[BILL_COL.QTY - 1]) || 0;
        if (itemName) {
          const key = itemName + '|' + size;
          billQtyMap[key] = (billQtyMap[key] || 0) + qty;
        }
      }
    }
  }

  // 1b. Net out vendor returns (goods sent back) from the same map — a
  // return is the mirror image of a bill, debiting the same Base Unit
  // quantity that the original bill credited. See module_return.js.
  const returnSheet = ss.getSheetByName(APP_CONFIG.SHEETS.RETURN);
  if (returnSheet) {
    const returnLastRow = returnSheet.getLastRow();
    if (returnLastRow >= 2) {
      const numCols = Math.max(returnSheet.getLastColumn(), RETURN_COL.BASE_RATE);
      const returnData = returnSheet.getRange(2, 1, returnLastRow - 1, numCols).getValues();
      for (let i = 0; i < returnData.length; i++) {
        const row = returnData[i];
        const itemName = String(row[RETURN_COL.ITEM_NAME - 1] || '').trim().toLowerCase();
        const size = String(row[RETURN_COL.SIZE - 1] || '').trim().toLowerCase();
        const rawBaseQty = row[RETURN_COL.BASE_QTY - 1];
        const qty = (rawBaseQty !== '' && rawBaseQty !== null && rawBaseQty !== undefined)
          ? Number(rawBaseQty) || 0
          : Number(row[RETURN_COL.QTY - 1]) || 0;
        if (itemName) {
          const key = itemName + '|' + size;
          billQtyMap[key] = (billQtyMap[key] || 0) - qty;
        }
      }
    }
  }

  // 1c. Net out wastage (component losses) from the same map — wasted
  // items are destroyed in-house and never return to stock.
  const wastageSheet = ss.getSheetByName(APP_CONFIG.SHEETS.WASTAGE);
  if (wastageSheet) {
    const wastageLastRow = wastageSheet.getLastRow();
    if (wastageLastRow >= 2) {
      const numCols = Math.max(wastageSheet.getLastColumn(), WASTAGE_COL.BASE_QTY);
      const wastageData = wastageSheet.getRange(2, 1, wastageLastRow - 1, numCols).getValues();
      for (let i = 0; i < wastageData.length; i++) {
        const row = wastageData[i];
        const itemName = String(row[WASTAGE_COL.ITEM_NAME - 1] || '').trim().toLowerCase();
        const size = String(row[WASTAGE_COL.SIZE - 1] || '').trim().toLowerCase();
        const rawBaseQty = row[WASTAGE_COL.BASE_QTY - 1];
        const qty = (rawBaseQty !== '' && rawBaseQty !== null && rawBaseQty !== undefined)
          ? Number(rawBaseQty) || 0
          : Number(row[WASTAGE_COL.QTY - 1]) || 0;
        if (itemName) {
          const key = itemName + '|' + size;
          billQtyMap[key] = (billQtyMap[key] || 0) - qty;
        }
      }
    }
  }

  // 1d. Net out ad-hoc Issued Stock (module_issue.js) from the same map —
  // items issued outside a Process's recipe/BOM (and deliberately kept off
  // any Production lot's Components Consumed list) still leave the building,
  // so they debit Stock the same way Wastage does.
  const issueSheet = ss.getSheetByName(APP_CONFIG.SHEETS.ISSUE);
  if (issueSheet) {
    const issueLastRow = issueSheet.getLastRow();
    if (issueLastRow >= 2) {
      const numCols = Math.max(issueSheet.getLastColumn(), ISSUE_COL.BASE_QTY);
      const issueData = issueSheet.getRange(2, 1, issueLastRow - 1, numCols).getValues();
      for (let i = 0; i < issueData.length; i++) {
        const row = issueData[i];
        const itemName = String(row[ISSUE_COL.ITEM_NAME - 1] || '').trim().toLowerCase();
        const size = String(row[ISSUE_COL.SIZE - 1] || '').trim().toLowerCase();
        const rawBaseQty = row[ISSUE_COL.BASE_QTY - 1];
        const qty = (rawBaseQty !== '' && rawBaseQty !== null && rawBaseQty !== undefined)
          ? Number(rawBaseQty) || 0
          : Number(row[ISSUE_COL.QTY - 1]) || 0;
        if (itemName) {
          const key = itemName + '|' + size;
          billQtyMap[key] = (billQtyMap[key] || 0) - qty;
        }
      }
    }
  }

  // 2. Build Consumed Qty Map from each Completed Production lot's own
  // Components Consumed list (ITEM-sourced entries only — POOL-sourced
  // entries are debited from the Warehouse Pool instead, see
  // module_warehouse.js's recalculateWarehousePool()).
  const consumedQtyMap = {};
  const prodSheet = ss.getSheetByName(APP_CONFIG.SHEETS.PRODUCTION);
  if (prodSheet) {
    if (typeof ensureProductionWarehouseColumns === 'function') {
      ensureProductionWarehouseColumns(prodSheet);
    }
    const prodLastRow = prodSheet.getLastRow();
    if (prodLastRow >= 2) {
      const numCols = Math.max(prodSheet.getLastColumn(), PRODUCTION_COL.COMPONENTS_CONSUMED);
      const prodData = prodSheet.getRange(2, 1, prodLastRow - 1, numCols).getValues();
      for (let i = 0; i < prodData.length; i++) {
        const row = prodData[i];
        const status = String(row[PRODUCTION_COL.STATUS - 1] || '').trim().toLowerCase();
        if (status !== 'completed') continue;

        const rawComponents = String(row[PRODUCTION_COL.COMPONENTS_CONSUMED - 1] || '').trim();
        if (!rawComponents) continue;

        let components = [];
        try {
          const parsed = JSON.parse(rawComponents);
          if (Array.isArray(parsed)) components = parsed;
        } catch (e) {
          continue;
        }

        components.forEach(comp => {
          const sourceType = String(comp.sourceType || '').trim().toUpperCase();
          if (sourceType === COMPONENT_SOURCE_TYPES.POOL) return;
          const itemName = String(comp.itemName || '').trim().toLowerCase();
          const size = String(comp.size || '').trim().toLowerCase();
          const qty = Number(comp.qty) || 0;
          if (!itemName || qty <= 0) return;

          const compKey = itemName + '|' + size;
          consumedQtyMap[compKey] = (consumedQtyMap[compKey] || 0) + qty;
        });
      }
    }
  }

  return { billQtyMap, consumedQtyMap };
}

/**
 * Recalculates Current Stock for all items in the Stock sheet.
 * Formula: Current Stock = Initial Stock + Billed Qty - Consumed Qty (via finalized production).
 */
function recalculateStock() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(STOCK_LOCK_TIMEOUT_MS)) {
    Log.warn('[recalculateStock] Could not acquire lock. Skipping.');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let stockSheet = ss.getSheetByName(APP_CONFIG.SHEETS.STOCK);
    if (!stockSheet) {
      initStockSheet();
      stockSheet = ss.getSheetByName(APP_CONFIG.SHEETS.STOCK);
    }

    const stockLastRow = stockSheet.getLastRow();
    if (stockLastRow < 2) {
      lock.releaseLock();
      return;
    }

    const { billQtyMap, consumedQtyMap } = _getBilledAndConsumedQtyMaps(ss);

    // 4. Update Current Stock in Stock Sheet
    const stockRange = stockSheet.getRange(2, 1, stockLastRow - 1, 6);
    const stockData = stockRange.getValues();
    const updateValues = [];

    for (let i = 0; i < stockData.length; i++) {
      const row = stockData[i];
      const itemName = String(row[STOCK_COL.ITEM_NAME - 1] || '').trim().toLowerCase();
      const size = String(row[STOCK_COL.SIZE - 1] || '').trim().toLowerCase();
      const initialStock = Number(row[STOCK_COL.INITIAL_STOCK - 1]) || 0;
      
      const key = itemName + '|' + size;
      const billedQty = billQtyMap[key] || 0;
      const consumedQty = consumedQtyMap[key] || 0;
      const currentStock = initialStock + billedQty - consumedQty;

      updateValues.push([currentStock]);
    }

    if (updateValues.length > 0) {
      stockSheet.getRange(2, STOCK_COL.CURRENT_STOCK, updateValues.length, 1).setValues(updateValues);
    }
    
    SpreadsheetApp.flush();
    Logger.log('[recalculateStock] Successfully updated stock values.');
  } catch (error) {
    Log.error('[recalculateStock] Error:', error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Retrieves all stock data from the sheet.
 */
function getStockData() {
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.STOCK);
    } catch (e) {
      initStockSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.STOCK);
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return buildResponse(true, []);
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const records = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const name = String(row[STOCK_COL.ITEM_NAME - 1] || '').trim();
      const size = String(row[STOCK_COL.SIZE - 1] || '').trim();
      if (!name) continue;

      const initialStock = Number(row[STOCK_COL.INITIAL_STOCK - 1]) || 0;
      const currentStock = Number(row[STOCK_COL.CURRENT_STOCK - 1]) || 0;
      const threshold = Number(row[STOCK_COL.THRESHOLD - 1]) || 0;
      const rawDeadStock = row[STOCK_COL.DEAD_STOCK - 1];
      const deadStock = rawDeadStock === true || String(rawDeadStock).toLowerCase() === 'true';

      records.push({
        name,
        size,
        initialStock,
        currentStock,
        threshold,
        isLowStock: currentStock < threshold,
        deadStock
      });
    }
    
    // Sort stock items alphabetically by Item Name, then by Size
    records.sort((a, b) => {
      const nameComp = a.name.localeCompare(b.name);
      if (nameComp !== 0) return nameComp;
      return a.size.localeCompare(b.size);
    });
    
    return buildResponse(true, records);
  } catch (error) {
    Log.error('[getStockData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load stock data: ' + error.message);
  }
}

/**
 * Updates the threshold value for an item.
 */
function updateThreshold(itemName, size, threshold) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(STOCK_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.STOCK);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return buildResponse(false, null, 'No items found in stock database.');
    }

    const tName = String(itemName || '').trim().toLowerCase();
    const tSize = String(size || '').trim().toLowerCase();
    const thresholdVal = Number(threshold);

    if (isNaN(thresholdVal) || thresholdVal < 0) {
      return buildResponse(false, null, 'Threshold must be a valid non-negative number.');
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    let rowFound = -1;

    for (let i = 0; i < data.length; i++) {
      const nameVal = String(data[i][0] || '').trim().toLowerCase();
      const sizeVal = String(data[i][1] || '').trim().toLowerCase();
      if (nameVal === tName && sizeVal === tSize) {
        rowFound = i + 2; // 2-indexed row
        break;
      }
    }

    if (rowFound === -1) {
      return buildResponse(false, null, 'Item not found in Stock database.');
    }

    sheet.getRange(rowFound, STOCK_COL.THRESHOLD).setValue(thresholdVal);
    SpreadsheetApp.flush();

    logAction('UPDATE', APP_CONFIG.SHEETS.STOCK, `${itemName} (${size})`, `Updated threshold to ${thresholdVal}`, 'SUCCESS');
    return buildResponse(true, null, 'Threshold updated successfully.');
  } catch (error) {
    Log.error('[updateThreshold] Error:', error.message);
    return buildResponse(false, null, 'Failed to update threshold: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Marks (or unmarks) an item as dead stock by writing TRUE/FALSE to the
 * Dead Stock column. Only that single cell is written — no recalculation,
 * no lock needed beyond a quick point write.
 */
function updateDeadStock(itemName, size, isDeadStock) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(STOCK_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.STOCK);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return buildResponse(false, null, 'No items found in stock database.');
    }

    const tName = String(itemName || '').trim().toLowerCase();
    const tSize = String(size || '').trim().toLowerCase();
    const deadStockVal = isDeadStock === true || String(isDeadStock).toLowerCase() === 'true';

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    let rowFound = -1;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toLowerCase() === tName &&
          String(data[i][1] || '').trim().toLowerCase() === tSize) {
        rowFound = i + 2;
        break;
      }
    }

    if (rowFound === -1) {
      return buildResponse(false, null, 'Item not found in Stock database.');
    }

    sheet.getRange(rowFound, STOCK_COL.DEAD_STOCK).setValue(deadStockVal);
    SpreadsheetApp.flush();

    logAction('UPDATE', APP_CONFIG.SHEETS.STOCK, `${itemName} (${size})`,
      `Dead Stock marked as ${deadStockVal}`, 'SUCCESS');
    return buildResponse(true, { deadStock: deadStockVal }, 'Dead stock status updated.');
  } catch (error) {
    Log.error('[updateDeadStock] Error:', error.message);
    return buildResponse(false, null, 'Failed to update dead stock status: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Records a Current Stock change (manual adjustment or import-driven reset) to
 * the audit Logs sheet in a fixed, parseable format so getStockAdjustmentHistory()
 * can surface it in the Item Ledger's Transaction Ledger History.
 * @param {string} action - 'ADJUST' (manual correction) or 'RESET' (re-import overwrite)
 */
function _logStockChange(action, itemName, size, oldVal, newVal, reason) {
  const details = `Old: ${oldVal}, New: ${newVal}. Reason: ${reason}`;
  logAction(action, APP_CONFIG.SHEETS.STOCK, `${itemName} (${size})`, details, 'SUCCESS');
}

/**
 * Manually corrects an item's Current Stock (e.g. after a physical recount or
 * to fix a data error), setting it as the new ground-truth value. A reason is
 * required and the adjustment is recorded in the audit Logs sheet so manual
 * corrections stay traceable. Negative values are allowed — over-consumption
 * (e.g. production issuing more than was on hand) can legitimately leave
 * stock negative until the user reviews and corrects it.
 * @param {string} itemName
 * @param {string} size
 * @param {number} newCurrentStock - The corrected Current Stock value.
 * @param {string} reason - Why the adjustment was made (required).
 */
function adjustStockManually(itemName, size, newCurrentStock, reason) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(STOCK_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.STOCK);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return buildResponse(false, null, 'No items found in stock database.');
    }

    const newStockVal = Number(newCurrentStock);
    if (isNaN(newStockVal)) {
      return buildResponse(false, null, 'Corrected stock must be a valid number.');
    }

    const reasonText = String(reason || '').trim();
    if (!reasonText) {
      return buildResponse(false, null, 'A reason is required for manual stock adjustments.');
    }

    const tName = String(itemName || '').trim().toLowerCase();
    const tSize = String(size || '').trim().toLowerCase();

    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    let rowFound = -1;
    let oldCurrentStock = 0;

    for (let i = 0; i < data.length; i++) {
      const nameVal = String(data[i][0] || '').trim().toLowerCase();
      const sizeVal = String(data[i][1] || '').trim().toLowerCase();
      if (nameVal === tName && sizeVal === tSize) {
        rowFound = i + 2; // 2-indexed row
        oldCurrentStock = Number(data[i][3]) || 0;
        break;
      }
    }

    if (rowFound === -1) {
      return buildResponse(false, null, 'Item not found in Stock database.');
    }

    if (newStockVal === oldCurrentStock) {
      return buildResponse(false, null, 'New stock value is the same as the current value — nothing to adjust.');
    }

    // Back-solve Initial Stock so recalculateStock() derives Current Stock === newStockVal
    // (same ground-truth approach used by importStockData).
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const { billQtyMap, consumedQtyMap } = _getBilledAndConsumedQtyMaps(ss);
    const key = tName + '|' + tSize;
    const billedQty = billQtyMap[key] || 0;
    const consumedQty = consumedQtyMap[key] || 0;
    const newInitialStock = newStockVal - billedQty + consumedQty;

    sheet.getRange(rowFound, STOCK_COL.INITIAL_STOCK, 1, 2).setValues([[newInitialStock, newStockVal]]);
    SpreadsheetApp.flush();

    _logStockChange('ADJUST', itemName, size, oldCurrentStock, newStockVal, reasonText);

    return buildResponse(true, { oldCurrentStock, newCurrentStock: newStockVal }, 'Stock adjusted successfully.');
  } catch (error) {
    Log.error('[adjustStockManually] Error:', error.message);
    logAction('ERROR', 'adjustStockManually', `${itemName} (${size})`, error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to adjust stock: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Imports a stock list from Excel/CSV parsed JSON.
 * The imported quantity is treated as the new ground-truth Current Stock for
 * each item (a fresh stock count), for both existing items and newly added ones.
 * Since Current Stock is always re-derived as Initial Stock + Billed Qty - Consumed
 * Qty (see recalculateStock), Initial Stock is back-computed here so that derived
 * value lands exactly on the imported quantity instead of drifting by whatever
 * historical bills/production already apply to that item.
 * @param {Array} items - Array of items: [{ name, size, initialStock }]
 */
function importStockData(items) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(STOCK_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    if (!Array.isArray(items) || items.length === 0) {
      return buildResponse(false, null, 'Import file is empty or invalid.');
    }

    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.STOCK);
    } catch (e) {
      initStockSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.STOCK);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const { billQtyMap, consumedQtyMap } = _getBilledAndConsumedQtyMaps(ss);

    const lastRow = sheet.getLastRow();
    const existingItems = {};
    let stockData = [];
    if (lastRow >= 2) {
      stockData = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      for (let i = 0; i < stockData.length; i++) {
        const nameVal = String(stockData[i][0] || '').trim().toLowerCase();
        const sizeVal = String(stockData[i][1] || '').trim().toLowerCase();
        const rawDeadStock = stockData[i][5];
        existingItems[nameVal + '|' + sizeVal] = {
          rowNum: i + 2,
          threshold: Number(stockData[i][4]) || 0,
          oldCurrentStock: Number(stockData[i][3]) || 0,
          deadStock: rawDeadStock === true || String(rawDeadStock).toLowerCase() === 'true'
        };
      }
    }

    let updatedCount = 0;
    let addedCount = 0;
    const newRows = [];

    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      const name = String(item.name || '').trim();
      const size = String(item.size || '').trim();
      const qty = Number(item.initialStock);

      if (!name) continue;
      if (isNaN(qty) || qty < 0) continue;

      const key = name.toLowerCase() + '|' + size.toLowerCase();
      const billedQty = billQtyMap[key] || 0;
      const consumedQty = consumedQtyMap[key] || 0;
      // Back-solve Initial Stock so recalculateStock() derives Current Stock === qty.
      const newInitialStock = qty - billedQty + consumedQty;

      if (existingItems[key]) {
        const rowNum = existingItems[key].rowNum;
        const oldCurrentStock = existingItems[key].oldCurrentStock;
        sheet.getRange(rowNum, STOCK_COL.INITIAL_STOCK, 1, 2).setValues([[newInitialStock, qty]]);
        if (qty !== oldCurrentStock) {
          _logStockChange('RESET', name, size, oldCurrentStock, qty, 'Re-imported initial stock');
        }
        updatedCount++;
      } else {
        // Queue for bulk append
        newRows.push([
          name,
          size,
          newInitialStock,
          qty,   // New ground-truth Current Stock
          0,     // Default Threshold = 0
          false  // Default Dead Stock = false
        ]);
        addedCount++;
      }
    }

    // Append new rows if any
    if (newRows.length > 0) {
      const appendStart = sheet.getLastRow() + 1;
      sheet.getRange(appendStart, 1, newRows.length, 6).setValues(newRows);
    }

    SpreadsheetApp.flush();

    // Recalculate stock values dynamically to factor in Bills and Production Lots immediately
    recalculateStock();

    // Keep Items Master automatically in sync with any newly imported Stock items
    if (typeof importItemsFromStock === 'function') {
      importItemsFromStock();
    }

    const msg = `Stock import completed. Items updated: ${updatedCount}, Items added: ${addedCount}.`;
    logAction('CREATE', APP_CONFIG.SHEETS.STOCK, 'IMPORT', msg, 'SUCCESS');
    return buildResponse(true, { updatedCount, addedCount }, msg);
  } catch (error) {
    Log.error('[importStockData] Error:', error.message);
    logAction('ERROR', 'importStockData', 'IMPORT', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to import stock data: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Returns the history of manual Current Stock corrections ('ADJUST') and
 * import-driven stock resets ('RESET'), sourced from the audit Logs sheet
 * (see _logStockChange). Used by the Item Ledger view to surface these
 * alongside Bills/POs/Production in the Transaction Ledger History.
 */
function getStockAdjustmentHistory() {
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
      if ((action !== 'ADJUST' && action !== 'RESET') || sheetName !== APP_CONFIG.SHEETS.STOCK) continue;

      const recordId = String(row[LOGS_COL.RECORD_ID - 1] || '');
      const idMatch = recordId.match(/^(.*) \((.*)\)$/);
      if (!idMatch) continue;

      const details = String(row[LOGS_COL.DETAILS - 1] || '');
      const detailsMatch = details.match(/Old:\s*([\d.\-]+),\s*New:\s*([\d.\-]+)\.\s*Reason:\s*(.*)$/);
      if (!detailsMatch) continue;

      records.push({
        date: row[LOGS_COL.TIMESTAMP - 1],
        action: action, // 'ADJUST' | 'RESET'
        itemName: idMatch[1],
        size: idMatch[2],
        oldValue: Number(detailsMatch[1]),
        newValue: Number(detailsMatch[2]),
        reason: detailsMatch[3],
        user: row[LOGS_COL.USER - 1]
      });
    }

    return buildResponse(true, records);
  } catch (error) {
    Log.error('[getStockAdjustmentHistory] Error:', error.message);
    return buildResponse(false, null, 'Failed to load stock adjustment history: ' + error.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────
// ITEMS MASTER SYNC (Item Name + Size as common key)
// ─────────────────────────────────────────────────────────────────────────

/**
 * Finds the row of a Stock entry matching name + size (case-insensitive).
 * @returns {number} 1-based row number, or -1 if not found
 */
function _findStockRow(sheet, name, size) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const data = sheet.getRange(2, STOCK_COL.ITEM_NAME, lastRow - 1, 2).getValues();
  const tName = String(name || '').trim().toLowerCase();
  const tSize = String(size || '').trim().toLowerCase();

  for (let i = 0; i < data.length; i++) {
    const rowName = String(data[i][0] || '').trim().toLowerCase();
    const rowSize = String(data[i][1] || '').trim().toLowerCase();
    if (rowName === tName && rowSize === tSize) {
      return i + 2;
    }
  }
  return -1;
}

/**
 * Keeps the Stock sheet in sync with Items Master changes, using
 * Item Name + Size as the common key. Called from module_items.js
 * and module_vendors.js after items are created/renamed/removed.
 *
 * Runs under the caller's existing document lock — does not acquire its own.
 *
 * @param {string} action - 'ensure' | 'rename' | 'merge' | 'remove'
 * @param {Object} payload
 *   - ensure: { name, size, initialStock? } — adds a Stock row if one doesn't exist,
 *     seeding Initial/Current Stock with initialStock (default 0)
 *   - rename: { oldName, oldSize, newName, newSize } — renames the matching row in place
 *   - merge: { oldName, oldSize, newName, newSize } — adds the old row's Initial/Current
 *     Stock into the new row's, then deletes the old row (used when two item rows are
 *     consolidated into one; unlike rename, both rows may already exist)
 *   - remove: { name, size } — deletes the matching row
 */
function syncStockForItem(action, payload) {
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.STOCK);
    } catch (e) {
      initStockSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.STOCK);
    }

    if (action === 'ensure') {
      const { name, size, initialStock } = payload;
      if (_findStockRow(sheet, name, size) === -1) {
        const qty = Number(initialStock) || 0;
        sheet.appendRow([name, size, qty, qty, 0, false]);
      }
    } else if (action === 'rename') {
      const { oldName, oldSize, newName, newSize } = payload;
      const oldRow = _findStockRow(sheet, oldName, oldSize);
      const newRowExists = _findStockRow(sheet, newName, newSize) !== -1;

      if (oldRow === -1) {
        if (!newRowExists) {
          sheet.appendRow([newName, newSize, 0, 0, 0, false]);
        }
      } else if (!newRowExists) {
        // Safe to rename in place — preserves stock figures
        sheet.getRange(oldRow, STOCK_COL.ITEM_NAME, 1, 2).setValues([[newName, newSize]]);
      }
      // If the new key already has its own row, leave the old row as-is
      // (it may still belong to other items sharing the old name + size).
    } else if (action === 'merge') {
      const { oldName, oldSize, newName, newSize } = payload;
      const oldRow = _findStockRow(sheet, oldName, oldSize);
      const newRow = _findStockRow(sheet, newName, newSize);

      if (oldRow === -1) {
        // Nothing to merge in; ensure the target row exists.
        if (newRow === -1) {
          sheet.appendRow([newName, newSize, 0, 0, 0, false]);
        }
      } else if (newRow === -1) {
        // Target row doesn't exist yet — equivalent to a plain rename.
        sheet.getRange(oldRow, STOCK_COL.ITEM_NAME, 1, 2).setValues([[newName, newSize]]);
      } else {
        const oldVals = sheet.getRange(oldRow, STOCK_COL.INITIAL_STOCK, 1, 2).getValues()[0];
        const newVals = sheet.getRange(newRow, STOCK_COL.INITIAL_STOCK, 1, 2).getValues()[0];
        const mergedInitial = (Number(oldVals[0]) || 0) + (Number(newVals[0]) || 0);
        const mergedCurrent = (Number(oldVals[1]) || 0) + (Number(newVals[1]) || 0);

        sheet.getRange(newRow, STOCK_COL.INITIAL_STOCK, 1, 2).setValues([[mergedInitial, mergedCurrent]]);
        sheet.deleteRow(oldRow);
      }
    } else if (action === 'remove') {
      const { name, size } = payload;
      const row = _findStockRow(sheet, name, size);
      if (row !== -1) {
        sheet.deleteRow(row);
      }
    }

    SpreadsheetApp.flush();
  } catch (error) {
    Log.error('[syncStockForItem] Error:', error.message);
    logAction('ERROR', 'syncStockForItem', JSON.stringify(payload), error.message, 'ERROR');
  }
}

/**
 * backfillPoAndBillBaseQty()
 *
 * One-time migration for PO Tracker / Bill Ledger rows saved before
 * BASE_QTY/BASE_RATE existed. Ensures both sheets have headers for the new
 * columns, then sets BASE_QTY = QTY and BASE_RATE = PRICE for every row
 * missing them — legacy rows are implicitly Pcs (factor 1), so this keeps
 * recalculateStock()'s Current Stock numbers unchanged after migration.
 *
 * Run manually once from the Apps Script editor after deploying this
 * change, then run recalculateStock() (or let it run on its own trigger) to
 * confirm Stock totals didn't move.
 *
 * @returns {Object} API response with { poRowsUpdated, billRowsUpdated }
 */
function backfillPoAndBillBaseQty() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(STOCK_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    let poRowsUpdated = 0;
    let billRowsUpdated = 0;

    // ─── PO Tracker ───
    const poSheet = getSheet(APP_CONFIG.SHEETS.PO);
    if (poSheet) {
      const lastCol = poSheet.getLastColumn();
      if (lastCol < PO_COL.BASE_RATE) {
        poSheet.getRange(2, PO_COL.BASE_QTY, 1, 2).setValues([['Base Qty', 'Base Rate']]);
      } else {
        const header = String(poSheet.getRange(2, PO_COL.BASE_QTY).getValue() || '').trim();
        if (!header) {
          poSheet.getRange(2, PO_COL.BASE_QTY, 1, 2).setValues([['Base Qty', 'Base Rate']]);
        }
      }

      const startRow = APP_CONFIG.PO_SETTINGS.DATA_START_ROW;
      const lastRow = poSheet.getLastRow();
      if (lastRow >= startRow) {
        const range = poSheet.getRange(startRow, PO_COL.QTY, lastRow - startRow + 1, PO_COL.BASE_RATE - PO_COL.QTY + 1);
        const data = range.getValues();
        const qtyOff = 0;
        const priceOff = PO_COL.PRICE - PO_COL.QTY;
        const baseQtyOff = PO_COL.BASE_QTY - PO_COL.QTY;
        const baseRateOff = PO_COL.BASE_RATE - PO_COL.QTY;

        data.forEach(row => {
          const baseQtyRaw = row[baseQtyOff];
          if (baseQtyRaw === '' || baseQtyRaw === null || baseQtyRaw === undefined) {
            row[baseQtyOff] = Number(row[qtyOff]) || 0;
            row[baseRateOff] = Number(row[priceOff]) || 0;
            poRowsUpdated++;
          }
        });

        range.setValues(data);
      }
    }

    // ─── Bill Ledger ───
    const billSheet = getSheet(APP_CONFIG.SHEETS.BILL);
    if (billSheet) {
      if (typeof _ensureBillSheetColumns === 'function') {
        _ensureBillSheetColumns(billSheet);
      }

      const startRow = APP_CONFIG.BILL_SETTINGS.DATA_START_ROW;
      const lastRow = billSheet.getLastRow();
      if (lastRow >= startRow) {
        const range = billSheet.getRange(startRow, BILL_COL.QTY, lastRow - startRow + 1, BILL_COL.BASE_RATE - BILL_COL.QTY + 1);
        const data = range.getValues();
        const qtyOff = 0;
        const priceOff = BILL_COL.PRICE - BILL_COL.QTY;
        const baseQtyOff = BILL_COL.BASE_QTY - BILL_COL.QTY;
        const baseRateOff = BILL_COL.BASE_RATE - BILL_COL.QTY;

        data.forEach(row => {
          const baseQtyRaw = row[baseQtyOff];
          if (baseQtyRaw === '' || baseQtyRaw === null || baseQtyRaw === undefined) {
            row[baseQtyOff] = Number(row[qtyOff]) || 0;
            row[baseRateOff] = Number(row[priceOff]) || 0;
            billRowsUpdated++;
          }
        });

        range.setValues(data);
      }
    }

    SpreadsheetApp.flush();

    const msg = `Backfilled ${poRowsUpdated} PO row(s) and ${billRowsUpdated} Bill row(s) with Base Qty/Rate.`;
    logAction('UPDATE', 'PO_BILL', 'BACKFILL_BASE_QTY', msg, 'SUCCESS');
    return buildResponse(true, { poRowsUpdated, billRowsUpdated }, msg);
  } catch (error) {
    Log.error('[backfillPoAndBillBaseQty] Error:', error.message);
    logAction('ERROR', 'backfillPoAndBillBaseQty', 'BACKFILL_BASE_QTY', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to backfill Base Qty/Rate: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}
