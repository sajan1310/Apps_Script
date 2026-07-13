/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_clients.gs — CLIENTS & PI / ESTIMATES MODULE
 *
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * - Clients master CRUD (mirrors module_vendors.js)
 * - PI / Estimates ledger CRUD (Client Orders sheet, grouped by Order Number
 *   like BOM components are grouped by Product ID)
 * - PI → Production sync: when a PI/Estimate's status is set to
 *   "Order Confirmed", each not-yet-pushed line item is auto-queued as a
 *   Pending Production lot (one-time push, guarded by a per-line
 *   "Production Pushed" flag).
 *
 * Sheet Layout (Clients):
 * ───────────────────────────────────────────────────────────────────────────
 * Col A (1):  Client Name
 * Col B (2):  Contact Number
 * Col C (3):  Address
 * Col D (4):  GSTIN
 * Col E (5):  Remarks
 *
 * Sheet Layout (Client Orders / PI-Estimates):
 * ───────────────────────────────────────────────────────────────────────────
 * Col A (1):   Order Number (e.g. ORD-1001)
 * Col B (2):   Order Date (DD/MM/YYYY)
 * Col C (3):   Client Name
 * Col D (4):   Product_ID (must exist in BOM)
 * Col E (5):   Product Name
 * Col F (6):   Qty Ordered
 * Col G (7):   Status (Estimate | Order Confirmed | Cancelled)
 * Col H (8):   Order Remarks (repeated per line)
 * Col I (9):   Line Remarks
 * Col J (10):  Production Pushed ('Yes' once queued into Production)
 * ═══════════════════════════════════════════════════════════════════════════
 */

const CLIENT_LOCK_TIMEOUT_MS = 15000;
const CLIENT_ORDER_STATUSES = ['Estimate', 'Order Confirmed', 'Cancelled'];

// ─────────────────────────────────────────────────────────────────────────
// CLIENTS MASTER
// ─────────────────────────────────────────────────────────────────────────

/**
 * Run this once manually from the GAS Editor to setup the sheet headers
 */
function initClientsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.CLIENTS);
  if (!sheet) {
    sheet = ss.insertSheet(APP_CONFIG.SHEETS.CLIENTS);
  }
  sheet.getRange(1, 1, 1, 5).setValues([["Client Name", "Contact Number", "Address", "GSTIN", "Remarks"]]).setFontWeight("bold");
  SpreadsheetApp.flush();
}

/**
 * Retrieves all clients from the database.
 * @returns {Object} JSON Response with clients array.
 */
function getClientsData() {
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.CLIENTS);
    } catch (e) {
      initClientsSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.CLIENTS);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, []);

    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const clients = data.map(r => ({
      name: String(r[0] || "").trim(),
      contact: String(r[1] || "").trim(),
      address: String(r[2] || "").trim(),
      gstin: String(r[3] || "").trim(),
      remarks: String(r[4] || "").trim()
    })).filter(c => c.name !== "");

    clients.sort((a, b) => a.name.localeCompare(b.name));

    return buildResponse(true, clients);
  } catch (error) {
    Log.error('[getClientsData] Error:', error.message);
    return buildResponse(false, null, "Failed to load clients: " + error.message);
  }
}

/**
 * Renames a client across every sheet that holds a de-normalized snapshot of
 * their name as a plain string (no foreign key): Client Orders (PI/Estimates)
 * and Dispatch. Without this, a renamed client's existing PI/Estimates and
 * Dispatch history stays keyed under the old name, so
 * App.Client.calculateClientLedger (which filters by exact name match)
 * silently stops finding it — the client's order/dispatch history
 * effectively disappears from the Client Ledger even though the rows still
 * exist. Mirrors module_contractors.js's _renameContractorEverywhere. Runs
 * under the caller's existing document lock — does not acquire its own.
 * @private
 */
function _renameClientEverywhere(oldName, newName) {
  const tOldRaw = String(oldName || '').trim();
  const tOld = tOldRaw.toLowerCase();
  const tNew = String(newName || '').trim();
  // Plain string compare (not case-insensitive) so a casing-only edit
  // ("abc corp" -> "ABC Corp") still cascades — mirrors
  // module_vendors.js#saveClient's own guard fix; matching module_vendors.js/
  // module_contractors.js/module_units.js/module_tags.js, all of which
  // deliberately use case-SENSITIVE guards here for the same reason.
  if (!tOld || !tNew || tOldRaw === tNew) return;

  const renameInColumn = (sheetName, col) => {
    let sheet;
    try {
      sheet = getSheet(sheetName);
    } catch (e) {
      return; // Sheet doesn't exist yet, nothing to rename
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const range = sheet.getRange(2, col, lastRow - 1, 1);
    const values = range.getValues();
    let changed = false;

    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0] || '').trim().toLowerCase() === tOld) {
        values[i][0] = tNew;
        changed = true;
      }
    }

    if (changed) range.setValues(values);
  };

  renameInColumn(APP_CONFIG.SHEETS.CLIENT_ORDERS, CLIENT_ORDERS_COL.CLIENT_NAME);
  renameInColumn(APP_CONFIG.SHEETS.DISPATCH, DISPATCH_COL.CLIENT_NAME);
}

/**
 * @private
 * Returns an error message if clientName is referenced by a Client Order
 * (PI/Estimate) or a Dispatch record, or null if it's safe to delete.
 * Deleting a still-referenced client would remove it from the Clients
 * master while its order/dispatch rows stay behind — App.Client's ledger
 * modal looks the client up in the master first and bails if not found, so
 * its consolidated ledger becomes unreachable even though the underlying
 * data is intact. Same "blocked if referenced" pattern as
 * _contractorInUseMessage / deleteProcess.
 */
function _clientInUseMessage(clientName) {
  const name = String(clientName || '').trim().toLowerCase();
  if (!name) return null;

  const hasMatch = (sheetName, col) => {
    let sheet;
    try {
      sheet = getSheet(sheetName);
    } catch (e) {
      return false;
    }
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;
    const values = sheet.getRange(2, col, lastRow - 1, 1).getValues();
    return values.some(row => String(row[0] || '').trim().toLowerCase() === name);
  };

  if (hasMatch(APP_CONFIG.SHEETS.CLIENT_ORDERS, CLIENT_ORDERS_COL.CLIENT_NAME)) {
    return `Cannot delete "${clientName}": referenced by PI/Estimate records.`;
  }
  if (hasMatch(APP_CONFIG.SHEETS.DISPATCH, DISPATCH_COL.CLIENT_NAME)) {
    return `Cannot delete "${clientName}": referenced by Dispatch records.`;
  }
  return null;
}

/**
 * Creates or updates a client record.
 * @param {Object} formData The client data object
 * @returns {Object} Standardised API response
 */
function saveClient(formData) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CLIENT_LOCK_TIMEOUT_MS)) return buildResponse(false, null, "System busy. Please try again.");

  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.CLIENTS);
    } catch (e) {
      initClientsSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.CLIENTS);
    }

    const newName = sanitizeString(formData.clientName, "clientName");
    if (!newName) throw new Error("Client name cannot be empty.");

    const isEdit = !!formData.originalClientName;
    const originalName = isEdit ? sanitizeString(formData.originalClientName, "originalClientName") : newName;

    const data = sheet.getDataRange().getValues();
    let targetRow = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === originalName.toLowerCase()) {
        targetRow = i + 1;
        break;
      }
    }

    if (isEdit && targetRow === -1) throw new Error("Original client not found in sheet.");
    if (!isEdit && targetRow !== -1) throw new Error("A client with this name already exists.");

    if (isEdit && newName.toLowerCase() !== originalName.toLowerCase()) {
      const dupeCheck = data.findIndex((r, idx) => idx > 0 && String(r[0]).trim().toLowerCase() === newName.toLowerCase());
      if (dupeCheck !== -1) throw new Error("Another client with this new name already exists.");
    }

    const rowData = [
      newName,
      sanitizeString(formData.contact || "", "contact"),
      sanitizeString(formData.address || "", "address"),
      sanitizeString(formData.gstin || "", "gstin"),
      sanitizeString(formData.remarks || "", "remarks")
    ];

    if (isEdit) {
      sheet.getRange(targetRow, 1, 1, 5).setValues([rowData]);
      // Propagate the rename to every sheet holding a de-normalized copy of
      // the client name before it's out of scope (see _renameClientEverywhere).
      // Plain string compare (not case-insensitive) so a casing-only edit
      // still cascades — see _renameClientEverywhere's own comment.
      if (newName !== originalName) {
        _renameClientEverywhere(originalName, newName);
      }
    } else {
      sheet.appendRow(rowData);
    }

    SpreadsheetApp.flush();

    const action = isEdit ? 'UPDATE' : 'CREATE';
    const msg = isEdit ? "Client profile updated successfully." : "New client registered.";
    logAction(action, APP_CONFIG.SHEETS.CLIENTS, newName, JSON.stringify(formData), 'SUCCESS');

    return buildResponse(true, { name: newName }, msg);
  } catch (error) {
    logAction('ERROR', 'saveClient', formData ? formData.clientName : 'unknown', error.message, 'ERROR');
    return buildResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes a client record from the master sheet.
 * @param {string} clientName The name of the client to delete
 * @returns {Object} Standardised API response
 */
function deleteClient(clientName) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CLIENT_LOCK_TIMEOUT_MS)) return buildResponse(false, null, "System busy.");

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.CLIENTS);
    if (!sheet) throw new Error("Clients sheet not found.");

    const targetName = String(clientName || "").trim().toLowerCase();
    const data = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 1).getValues();
    let targetRow = -1;

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === targetName) {
        targetRow = i + 2;
        break;
      }
    }

    if (targetRow === -1) throw new Error("Client not found.");

    const inUseMessage = _clientInUseMessage(clientName);
    if (inUseMessage) throw new Error(inUseMessage);

    sheet.deleteRow(targetRow);
    SpreadsheetApp.flush();

    logAction('DELETE', APP_CONFIG.SHEETS.CLIENTS, clientName, `Client deleted`, 'SUCCESS');

    return buildResponse(true, null, `Client "${clientName}" deleted from master.`);
  } catch (error) {
    logAction('ERROR', 'deleteClient', clientName, error.message, 'ERROR');
    return buildResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes multiple client records from the master sheet in a single batch.
 * @param {Array<string>} clientNames - Names of the clients to delete
 * @returns {Object} Standardised API response
 */
function deleteClientsBulk(clientNames) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CLIENT_LOCK_TIMEOUT_MS)) return buildResponse(false, null, "System busy.");

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.CLIENTS);
    if (!sheet) throw new Error("Clients sheet not found.");

    const requestedNames = (clientNames || []).map(n => String(n || '').trim()).filter(Boolean);
    if (requestedNames.length === 0) {
      return buildResponse(true, null, 'No clients selected.');
    }

    // Skip any client with PI/Estimate or Dispatch history — deleting them
    // would orphan that history under a name no longer in the master list,
    // making it unreachable from the Client Ledger. Each referencing sheet
    // is read ONCE here into a name Set, instead of _clientInUseMessage()
    // re-reading both sheets from scratch for EVERY requested name (O(K x 2N)
    // instead of O(2N + K)) — same pattern as deleteClientOrdersBulk. The
    // bulk path only needs the yes/no in-use verdict, not a reason string.
    const inUseNames = new Set();
    const addNamesFromSheet = (sheetName, col) => {
      let sh;
      try { sh = getSheet(sheetName); } catch (e) { return; }
      const lastRow = sh.getLastRow();
      if (lastRow < 2) return;
      sh.getRange(2, col, lastRow - 1, 1).getValues().forEach(row => {
        const n = String(row[0] || '').trim().toLowerCase();
        if (n) inUseNames.add(n);
      });
    };
    addNamesFromSheet(APP_CONFIG.SHEETS.CLIENT_ORDERS, CLIENT_ORDERS_COL.CLIENT_NAME);
    addNamesFromSheet(APP_CONFIG.SHEETS.DISPATCH, DISPATCH_COL.CLIENT_NAME);

    const skipped = requestedNames.filter(n => inUseNames.has(n.toLowerCase()));
    const deletableNames = requestedNames.filter(n => !inUseNames.has(n.toLowerCase()));
    const targetSet = new Set(deletableNames.map(n => n.toLowerCase()));

    const lastRow = sheet.getLastRow();
    if (lastRow < 2 || targetSet.size === 0) {
      const msg = skipped.length
        ? `No clients deleted — all selected have PI/Estimate or Dispatch history: ${skipped.join(', ')}.`
        : 'No clients to delete.';
      return buildResponse(skipped.length === 0, null, msg);
    }

    const { rowsDeleted } = rewriteSheetExcludingRows(sheet, 2, row => {
      const name = String(row[CLIENTS_COL.CLIENT_NAME - 1] || '').trim().toLowerCase();
      return targetSet.has(name);
    });

    SpreadsheetApp.flush();

    let msg = `Deleted ${rowsDeleted} client(s) from master.`;
    if (skipped.length > 0) {
      msg += ` Skipped ${skipped.length} client(s) with PI/Estimate or Dispatch history: ${skipped.join(', ')}.`;
    }
    logAction('BULK_DELETE', APP_CONFIG.SHEETS.CLIENTS, 'multiple', msg, 'SUCCESS');

    return buildResponse(true, null, msg);
  } catch (error) {
    Log.error('[deleteClientsBulk] Error:', error.message);
    logAction('ERROR', 'deleteClientsBulk', 'multiple', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete clients: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────────────────
// PI / ESTIMATES (CLIENT ORDERS)
// ─────────────────────────────────────────────────────────────────────────

/**
 * Initializes the Client Orders (PI / Estimates) sheet with correct headers.
 */
function initClientOrdersSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.CLIENT_ORDERS);
    if (!sheet) {
      sheet = ss.insertSheet(APP_CONFIG.SHEETS.CLIENT_ORDERS);
    }

    const headers = [
      'Order Number',
      'Order Date',
      'Client Name',
      'Product_ID',
      'Product Name',
      'Qty Ordered',
      'Status',
      'Order Remarks',
      'Line Remarks',
      'Production Pushed'
    ];

    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Client Orders sheet initialized successfully.');
  } catch (error) {
    Log.error('[initClientOrdersSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize Client Orders sheet: ' + error.message);
  }
}

/**
 * Auto-generates the next sequential Order Number.
 * Format: ORD-1001, ORD-1002, ...
 */
function getNextOrderNumber() {
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.CLIENT_ORDERS);
    } catch (e) {
      initClientOrdersSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.CLIENT_ORDERS);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return 'ORD-1001';

    const ids = sheet.getRange(2, CLIENT_ORDERS_COL.ORDER_NUMBER, lastRow - 1, 1).getValues();
    let maxNum = 1000;

    ids.forEach(row => {
      const idStr = String(row[0] || '').trim();
      const match = idStr.match(/^ORD-(\d+)$/i);
      if (match) {
        const num = parseInt(match[1], 10);
        if (num > maxNum) maxNum = num;
      }
    });

    return 'ORD-' + (maxNum + 1);
  } catch (error) {
    Log.error('[getNextOrderNumber] Error:', error.message);
    return 'ORD-1001';
  }
}

/**
 * Retrieves all PI / Estimates grouped by Order Number.
 * @returns {Object} JSON Response with orders array:
 *   [{orderNumber, orderDate, clientName, orderRemarks, status,
 *     lines: [{productId, productName, qty, lineRemarks, productionPushed}]}]
 */
function getClientOrdersData() {
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.CLIENT_ORDERS);
    } catch (e) {
      initClientOrdersSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.CLIENT_ORDERS);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, []);

    const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const orderMap = {};

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const orderNumber = String(row[CLIENT_ORDERS_COL.ORDER_NUMBER - 1] || '').trim();
      if (!orderNumber) continue;

      if (!orderMap[orderNumber]) {
        const rawDate = row[CLIENT_ORDERS_COL.ORDER_DATE - 1];
        const dateStr = rawDate instanceof Date ? toSafeDateString(rawDate) : String(rawDate || '');

        orderMap[orderNumber] = {
          orderNumber: orderNumber,
          orderDate: dateStr,
          dateRaw: rawDate instanceof Date ? rawDate.toISOString() : null,
          clientName: String(row[CLIENT_ORDERS_COL.CLIENT_NAME - 1] || '').trim(),
          status: String(row[CLIENT_ORDERS_COL.STATUS - 1] || '').trim() || 'Estimate',
          orderRemarks: String(row[CLIENT_ORDERS_COL.ORDER_REMARKS - 1] || '').trim(),
          lines: []
        };
      }

      const productId = String(row[CLIENT_ORDERS_COL.PRODUCT_ID - 1] || '').trim();
      if (!productId) continue;

      const pushedRaw = String(row[CLIENT_ORDERS_COL.PRODUCTION_PUSHED - 1] || '').trim().toLowerCase();
      orderMap[orderNumber].lines.push({
        productId: productId,
        productName: String(row[CLIENT_ORDERS_COL.PRODUCT_NAME - 1] || '').trim(),
        qty: Number(row[CLIENT_ORDERS_COL.QTY_ORDERED - 1]) || 0,
        lineRemarks: String(row[CLIENT_ORDERS_COL.LINE_REMARKS - 1] || '').trim(),
        productionPushed: pushedRaw === 'yes',
        needsManualProduction: pushedRaw === 'manual'
      });
    }

    const orders = Object.values(orderMap);
    // Sort by Order Number descending (newest first)
    orders.sort((a, b) => b.orderNumber.localeCompare(a.orderNumber, undefined, { numeric: true }));

    return buildResponse(true, orders);
  } catch (error) {
    Log.error('[getClientOrdersData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load PI / Estimates: ' + error.message);
  }
}

/**
 * Creates a new PI / Estimate or updates an existing one (edit-by-replace,
 * same approach as saveBOM). Qty-only line items — no pricing fields.
 *
 * On status === 'Order Confirmed', any line item not yet pushed to
 * Production is auto-queued as a new Pending Production lot, and marked
 * as pushed (sticky flag, never reset by later status changes).
 *
 * @param {Object} formData {orderNumber (if edit), orderDate, clientName,
 *   orderRemarks, status, lines: [...] (JSON or array)}
 */
function saveClientOrder(formData) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CLIENT_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.CLIENT_ORDERS);
    } catch (e) {
      initClientOrdersSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.CLIENT_ORDERS);
    }

    const clientName = sanitizeString(formData.clientName, 'clientName');
    if (!clientName) {
      return buildResponse(false, null, 'Client is required.');
    }

    const status = sanitizeString(formData.status || 'Estimate', 'status');
    if (CLIENT_ORDER_STATUSES.indexOf(status) === -1) {
      return buildResponse(false, null, 'Invalid status value.');
    }

    const dateObj = toSafeDateObject(formData.orderDate) || new Date();

    const orderRemarks = sanitizeString(formData.orderRemarks || '', 'orderRemarks').slice(0, APP_CONFIG.VALIDATION.MAX_REMARKS_LENGTH);

    let lines;
    try {
      lines = typeof formData.lines === 'string' ? JSON.parse(formData.lines) : (formData.lines || []);
    } catch (e) {
      return buildResponse(false, null, 'Invalid product line data format.');
    }

    if (!Array.isArray(lines) || lines.length === 0) {
      return buildResponse(false, null, 'PI / Estimate must contain at least one product line.');
    }

    // Product lines must reference products that exist in the BOM
    const bomResponse = getBOMProductionData();
    const bomProducts = (bomResponse && bomResponse.data) || [];
    const validProductIds = new Set(bomProducts.map(p => p.productId.toLowerCase()));
    const productNameMap = {};
    bomProducts.forEach(p => { productNameMap[p.productId.toLowerCase()] = p.productName; });

    const isEdit = !!formData.orderNumber;
    let orderNumber = isEdit ? sanitizeString(formData.orderNumber, 'orderNumber') : '';

    // Carry forward "Production Pushed" flags from existing rows, counted
    // per Product ID rather than a single boolean — two lines in the same
    // PI/Estimate can share a Product ID (e.g. a second batch of the same
    // product), and a boolean keyed only by productId can't tell them
    // apart: it would mark BOTH the already-pushed line AND a brand-new,
    // never-pushed line for that same product as pushed, permanently
    // skipping the new line from ever reaching Production. Each new line
    // below "claims" one pushed credit (if any remain) in the same order
    // the lines were originally entered, so the count of pushed vs.
    // not-yet-pushed lines per product is preserved across the edit.
    const pushedCountByProduct = {};
    if (isEdit) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const existingData = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
        const exists = existingData.some(row => String(row[CLIENT_ORDERS_COL.ORDER_NUMBER - 1] || '').trim().toLowerCase() === orderNumber.toLowerCase());
        if (!exists) {
          return buildResponse(false, null, `PI / Estimate "${orderNumber}" not found.`);
        }

        existingData.forEach(row => {
          if (String(row[CLIENT_ORDERS_COL.ORDER_NUMBER - 1] || '').trim().toLowerCase() === orderNumber.toLowerCase()) {
            const pid = String(row[CLIENT_ORDERS_COL.PRODUCT_ID - 1] || '').trim().toLowerCase();
            const pushed = String(row[CLIENT_ORDERS_COL.PRODUCTION_PUSHED - 1] || '').trim().toLowerCase() === 'yes';
            if (pid && pushed) pushedCountByProduct[pid] = (pushedCountByProduct[pid] || 0) + 1;
          }
        });
      }

      deleteRowsById(orderNumber, sheet, 2, CLIENT_ORDERS_COL.ORDER_NUMBER);
    } else {
      orderNumber = getNextOrderNumber();
    }

    // Build + validate clean line items
    const cleanLines = [];
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const productId = sanitizeString(line.productId || '', 'productId');
      if (!productId) continue;

      if (!validProductIds.has(productId.toLowerCase())) {
        return buildResponse(false, null, `Product "${productId}" is not defined in BOM and cannot be added to a PI / Estimate.`);
      }

      const qty = validateNumber(line.qty, 0.001, 10000000);
      if (qty <= 0) continue;

      const pidKey = productId.toLowerCase();
      let alreadyPushed = false;
      if (pushedCountByProduct[pidKey] > 0) {
        pushedCountByProduct[pidKey]--;
        alreadyPushed = true;
      }

      cleanLines.push({
        productId: productId,
        productName: productNameMap[productId.toLowerCase()] || sanitizeString(line.productName || '', 'productName'),
        qty: qty,
        lineRemarks: sanitizeString(line.lineRemarks || '', 'lineRemarks'),
        productionPushed: alreadyPushed,
        needsManualProduction: false
      });
    }

    if (cleanLines.length === 0) {
      return buildResponse(false, null, 'Please specify at least one valid product line with quantity greater than zero.');
    }

    // PI -> Production sync: queue unpushed lines once status is Order Confirmed.
    // Lines that don't resolve to an unambiguous final-stage Process are
    // retried on every subsequent save (not marked pushed) instead of being
    // stuck — see _pushOrderLinesToProduction.
    let pushedCount = 0;
    let manualCount = 0;
    if (status === 'Order Confirmed') {
      const linesToPush = cleanLines.filter(l => !l.productionPushed);
      if (linesToPush.length > 0) {
        const pushResult = _pushOrderLinesToProduction(orderNumber, clientName, linesToPush);
        linesToPush.forEach(l => {
          if (pushResult.autoQueued.has(l.productId.toLowerCase())) {
            l.productionPushed = true;
          } else {
            l.needsManualProduction = true;
          }
        });
        pushedCount = pushResult.lotsCreated;
        manualCount = linesToPush.length - pushedCount;
      }
    }

    const rowsToWrite = cleanLines.map(line => [
      orderNumber,
      dateObj,
      clientName,
      line.productId,
      line.productName,
      line.qty,
      status,
      orderRemarks,
      line.lineRemarks,
      line.productionPushed ? 'Yes' : (line.needsManualProduction ? 'Manual' : '')
    ]);

    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToWrite.length, 10).setValues(rowsToWrite);

    SpreadsheetApp.flush();

    let logMsg = isEdit
      ? `PI / Estimate updated: ${orderNumber} for ${clientName} (${cleanLines.length} line(s)), status: ${status}.`
      : `PI / Estimate created: ${orderNumber} for ${clientName} (${cleanLines.length} line(s)), status: ${status}.`;
    if (pushedCount > 0) {
      logMsg += ` ${pushedCount} line(s) auto-queued into Production.`;
    }
    if (manualCount > 0) {
      logMsg += ` ${manualCount} line(s) need manual Production setup (no unambiguous final-stage Process found).`;
    }
    logAction(isEdit ? 'UPDATE' : 'CREATE', APP_CONFIG.SHEETS.CLIENT_ORDERS, orderNumber, logMsg, 'SUCCESS');

    let successMsg = isEdit ? 'PI / Estimate updated successfully.' : 'PI / Estimate saved successfully.';
    if (pushedCount > 0) {
      successMsg += ` ${pushedCount} product line(s) queued into Production.`;
    }
    if (manualCount > 0) {
      successMsg += ` ${manualCount} line(s) need manual Production setup — log them yourself in the Production tab (BOM doesn't unambiguously map to one final-stage Process, or it's multi-color).`;
    }

    return buildResponse(true, { orderNumber: orderNumber }, successMsg);
  } catch (error) {
    Log.error('[saveClientOrder] Error:', error.message);
    logAction('ERROR', 'saveClientOrder', formData.orderNumber || 'NEW', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to save PI / Estimate: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Appends a new Pending Production lot for each PI line whose product
 * resolves to exactly one final-stage Process (via getBOMFinalStageProcessMap)
 * that isn't itself multi-color — i.e. only when there's enough information
 * to write a row with a real Process ID, Lot Number, Output Item Name and a
 * COMMON-components recipe, so it actually shows up in the Production tab
 * and can credit the Warehouse Pool once completed. (A plain Product+Qty
 * row with no Process ID is invisible to getProductionData() and
 * recalculateWarehousePool() — both require it.)
 *
 * Lines that don't resolve (no BOM, no single final-stage Process, or a
 * multi-color process where the per-color split isn't known from a PI
 * line) are left unqueued for the operator to log manually — see
 * CLIENT_ORDERS_COL.PRODUCTION_PUSHED's 'Manual' value.
 *
 * Does not call recalculateStock()/recalculateWarehousePool() — a Pending
 * lot doesn't affect Stock or the Warehouse Pool until it's Completed.
 * @param {string} orderNumber
 * @param {string} clientName
 * @param {Array<Object>} lines - [{productId, productName, qty}]
 * @returns {Object} { autoQueued: Set<productIdLower>, lotsCreated: number }
 * @private
 */
function _pushOrderLinesToProduction(orderNumber, clientName, lines) {
  const result = { autoQueued: new Set(), lotsCreated: 0 };

  let prodSheet;
  try {
    prodSheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
  } catch (e) {
    initProductionSheet();
    prodSheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
  }

  ensureProductionExtraColumns(prodSheet);
  ensureProductionProcessColumns(prodSheet);
  ensureProductionContractorColumns(prodSheet);
  ensureProductionWarehouseColumns(prodSheet);
  ensureProductionColorColumn(prodSheet);
  ensureProductionColorBreakdownColumn(prodSheet);
  ensureProductionOrderNumberColumn(prodSheet);

  const finalStageByProduct = getBOMFinalStageProcessMap();
  const dateObj = new Date();
  const rowsToWrite = [];

  lines.forEach(line => {
    const process = finalStageByProduct[line.productId.toLowerCase()];
    if (!process) return; // no BOM, or 0/2+ final-stage processes referenced — ambiguous

    const colorGroupsResp = getProcessColorGroups(process.processId);
    const colorGroups = (colorGroupsResp && colorGroupsResp.data) || [];
    if (colorGroups.length > 0) return; // per-color split isn't known from a PI line

    const componentsResp = getProcessComponentsData(process.processId);
    const recipeComponents = (componentsResp && componentsResp.data) || [];
    const componentsConsumed = recipeComponents
      .filter(c => c.colorGroup === COMPONENT_COLOR_GROUP_COMMON)
      .map(c => ({
        itemName: c.itemName,
        size: c.size || '',
        color: '',
        sourceType: c.sourceType === COMPONENT_SOURCE_TYPES.POOL ? COMPONENT_SOURCE_TYPES.POOL : COMPONENT_SOURCE_TYPES.ITEM,
        qty: (Number(c.qtyPerUnit) || 0) * line.qty,
        colorGroup: COMPONENT_COLOR_GROUP_COMMON,
        // Blank = "already in the item's Base Unit" — see PROCESS_COMPONENTS_COL.UNIT.
        unit: c.unit || ''
      }))
      .filter(c => c.itemName && c.qty > 0);

    const lotNumber = _generateLotNumber(prodSheet, process.lotPrefix);

    const row = new Array(PRODUCTION_COL.ORDER_NUMBER).fill('');
    row[PRODUCTION_COL.DATE - 1] = dateObj;
    row[PRODUCTION_COL.PRODUCT_ID - 1] = line.productId;
    row[PRODUCTION_COL.PRODUCT_NAME - 1] = line.productName;
    row[PRODUCTION_COL.QTY - 1] = line.qty;
    row[PRODUCTION_COL.STATUS - 1] = 'Pending';
    row[PRODUCTION_COL.REMARKS - 1] = `Auto-queued from PI ${orderNumber} - ${clientName}`;
    row[PRODUCTION_COL.PROCESS_ID - 1] = process.processId;
    row[PRODUCTION_COL.LOT_NUMBER - 1] = lotNumber;
    row[PRODUCTION_COL.OUTPUT_ITEM_NAME - 1] = process.outputItemName;
    row[PRODUCTION_COL.COMPONENTS_CONSUMED - 1] = JSON.stringify(componentsConsumed);
    row[PRODUCTION_COL.ORDER_NUMBER - 1] = orderNumber;

    rowsToWrite.push(row);
    result.autoQueued.add(line.productId.toLowerCase());
  });

  if (rowsToWrite.length > 0) {
    const startRow = prodSheet.getLastRow() + 1;
    prodSheet.getRange(startRow, 1, rowsToWrite.length, PRODUCTION_COL.ORDER_NUMBER).setValues(rowsToWrite);
    result.lotsCreated = rowsToWrite.length;
    logAction('CREATE', APP_CONFIG.SHEETS.PRODUCTION, orderNumber, `Auto-queued ${rowsToWrite.length} production lot(s) from PI ${orderNumber} (${clientName}).`, 'SUCCESS');
  }

  return result;
}

/**
 * Deletes all line items of a PI / Estimate. Blocks deletion if the order
 * is referenced by any Dispatch record, or already has a Production lot
 * auto-queued from it (PRODUCTION_COL.ORDER_NUMBER).
 * @param {string} orderNumber
 */
function deleteClientOrder(orderNumber) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CLIENT_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.CLIENT_ORDERS);
    if (!sheet) throw new Error('Client Orders sheet not found.');

    const orderClean = String(orderNumber || '').trim();
    if (!orderClean) return buildResponse(false, null, 'Order number is required.');

    let dispatchSheet;
    try {
      dispatchSheet = getSheet(APP_CONFIG.SHEETS.DISPATCH);
    } catch (e) {
      // Dispatch sheet doesn't exist yet, so nothing references this order
    }

    if (dispatchSheet) {
      const dLastRow = dispatchSheet.getLastRow();
      if (dLastRow >= 2) {
        const refs = dispatchSheet.getRange(2, DISPATCH_COL.ORDER_NUMBER, dLastRow - 1, 1).getValues();
        const inUse = refs.some(row => String(row[0]).trim().toLowerCase() === orderClean.toLowerCase());
        if (inUse) {
          return buildResponse(false, null, `Cannot delete PI / Estimate "${orderClean}": it has dispatch records against it.`);
        }
      }
    }

    // Also block if a Production lot was already auto-queued from this order
    // (PRODUCTION_COL.ORDER_NUMBER — see _pushOrderLinesToProduction). This
    // used to be un-checkable: the only trace was a free-text Remarks note.
    let prodSheet;
    try {
      prodSheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
    } catch (e) {
      // Production sheet doesn't exist yet, so nothing references this order
    }

    if (prodSheet) {
      ensureProductionOrderNumberColumn(prodSheet);
      const pLastRow = prodSheet.getLastRow();
      if (pLastRow >= 2 && prodSheet.getLastColumn() >= PRODUCTION_COL.ORDER_NUMBER) {
        const prodRefs = prodSheet.getRange(2, PRODUCTION_COL.ORDER_NUMBER, pLastRow - 1, 1).getValues();
        const inProduction = prodRefs.some(row => String(row[0]).trim().toLowerCase() === orderClean.toLowerCase());
        if (inProduction) {
          return buildResponse(false, null, `Cannot delete PI / Estimate "${orderClean}": it already has a Production lot queued from it.`);
        }
      }
    }

    const rowsDeleted = deleteRowsById(orderClean, sheet, 2, CLIENT_ORDERS_COL.ORDER_NUMBER);
    if (rowsDeleted === 0) {
      return buildResponse(false, null, `PI / Estimate "${orderClean}" not found.`);
    }

    SpreadsheetApp.flush();

    const msg = `PI / Estimate "${orderClean}" deleted (${rowsDeleted} line(s) removed).`;
    logAction('DELETE', APP_CONFIG.SHEETS.CLIENT_ORDERS, orderClean, msg, 'SUCCESS');

    return buildResponse(true, null, msg);
  } catch (error) {
    Log.error('[deleteClientOrder] Error:', error.message);
    logAction('ERROR', 'deleteClientOrder', orderNumber, error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete PI / Estimate: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes multiple PI / Estimates in a single batch. Orders referenced by
 * any Dispatch record, or already auto-queued into Production, are skipped
 * and reported back rather than failing the whole batch.
 * @param {Array<string>} orderNumbers
 */
function deleteClientOrdersBulk(orderNumbers) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CLIENT_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.CLIENT_ORDERS);
    if (!sheet) throw new Error('Client Orders sheet not found.');

    const requested = (orderNumbers || []).map(o => String(o || '').trim()).filter(Boolean);
    if (requested.length === 0) {
      return buildResponse(true, null, 'No PI / Estimates selected.');
    }

    let inUseSet = new Set();
    let dispatchSheet;
    try {
      dispatchSheet = getSheet(APP_CONFIG.SHEETS.DISPATCH);
    } catch (e) {
      // Dispatch sheet doesn't exist yet, so nothing is in use
    }

    if (dispatchSheet) {
      const dLastRow = dispatchSheet.getLastRow();
      if (dLastRow >= 2) {
        const refs = dispatchSheet.getRange(2, DISPATCH_COL.ORDER_NUMBER, dLastRow - 1, 1).getValues();
        const usedOrders = new Set(refs.map(row => String(row[0]).trim().toLowerCase()));
        requested.forEach(o => {
          if (usedOrders.has(o.toLowerCase())) inUseSet.add(o);
        });
      }
    }

    // Also block any order already auto-queued into Production (see
    // deleteClientOrder's matching check — same PRODUCTION_COL.ORDER_NUMBER
    // reference, previously un-checkable here).
    let prodSheet;
    try {
      prodSheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
    } catch (e) {
      // Production sheet doesn't exist yet, so nothing is in use
    }

    if (prodSheet) {
      ensureProductionOrderNumberColumn(prodSheet);
      const pLastRow = prodSheet.getLastRow();
      if (pLastRow >= 2 && prodSheet.getLastColumn() >= PRODUCTION_COL.ORDER_NUMBER) {
        const prodRefs = prodSheet.getRange(2, PRODUCTION_COL.ORDER_NUMBER, pLastRow - 1, 1).getValues();
        const inProductionOrders = new Set(prodRefs.map(row => String(row[0]).trim().toLowerCase()));
        requested.forEach(o => {
          if (inProductionOrders.has(o.toLowerCase())) inUseSet.add(o);
        });
      }
    }

    const toDelete = requested.filter(o => !inUseSet.has(o));
    const targetSet = new Set(toDelete);

    let rowsDeleted = 0;
    let deletedIds = [];
    if (targetSet.size > 0) {
      ({ rowsDeleted, deletedIds } = _rewriteWithoutMatchingRowsBulk(sheet, 2, CLIENT_ORDERS_COL.ORDER_NUMBER, targetSet));
      SpreadsheetApp.flush();
    }

    let msg = `Deleted ${toDelete.length} PI / Estimate(s) (${rowsDeleted} line(s) removed).`;
    if (inUseSet.size > 0) {
      msg += ` Skipped ${inUseSet.size} PI / Estimate(s) with dispatch records or a queued Production lot against them: ${Array.from(inUseSet).join(', ')}.`;
    }
    logAction('BULK_DELETE', APP_CONFIG.SHEETS.CLIENT_ORDERS, 'multiple', msg, 'SUCCESS');

    // Matches deleteClientsBulk/deleteContractorsBulk's convention: if every
    // selected order was skipped (still has dispatch records), this call
    // accomplished nothing and must not present as a green "success" toast.
    return buildResponse(toDelete.length > 0, { deletedIds, skipped: Array.from(inUseSet) }, msg);
  } catch (error) {
    Log.error('[deleteClientOrdersBulk] Error:', error.message);
    logAction('ERROR', 'deleteClientOrdersBulk', 'multiple', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete PI / Estimates: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}
