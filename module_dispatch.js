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
 * - Dispatched Goods ledger CRUD (Dispatch sheet). Each entry is a "bill" —
 *   one Dispatch Number shared by one-or-more line-item rows (same
 *   header+lines pattern as Client Orders/PI-Estimates), so a single
 *   dispatch/delivery can carry several different Ready-to-Dispatch
 *   products at once, each with its own qty and an optional per-unit Rate
 *   for record-keeping. Entries may optionally reference a PI / Estimate
 *   (Order Number, shared by the whole bill) or be standalone "direct
 *   supply" dispatches.
 *
 * Sheet Layout (Dispatch) — one row per line item, DISPATCH_NUMBER repeated
 * across every line of the same bill; header-level fields (date, order,
 * client, transport, remarks, invoice/private-mark/GR, logistics contractor)
 * are duplicated onto every line row, same convention as Client Orders'
 * orderRemarks:
 * ───────────────────────────────────────────────────────────────────────────
 * Col A (1):   Dispatch Number (e.g. DSP-1001) — shared by every line of one bill
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
 * Col M (13):  Logistics Contractor (optional)
 * Col N (14):  Logistics Rate (snapshot)
 * Col O (15):  Logistics Cost (= Logistics Rate x this line's Qty)
 * Col P (16):  Rate (optional per-unit billing/sale rate, record-keeping only)
 * Col Q (17):  Amount (= Qty x Rate)
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
      'Logistics Cost',
      'Rate',
      'Amount'
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
 * Backfills the "Rate" and "Amount" columns on Dispatch sheets created
 * before per-line billing rate existed, so legacy rows don't throw when
 * read/written.
 */
function ensureDispatchRateColumns(sheet) {
  try {
    if (sheet.getLastColumn() < DISPATCH_COL.AMOUNT) {
      const startCol = sheet.getLastColumn() + 1;
      sheet.insertColumnsAfter(sheet.getLastColumn(), DISPATCH_COL.AMOUNT - sheet.getLastColumn());
      sheet.getRange(1, startCol, 1, DISPATCH_COL.AMOUNT - startCol + 1)
        .setValues([['Rate', 'Amount']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    Log.error('[ensureDispatchRateColumns] Error:', error.message);
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
        dispatchedQty: 0,
        // Same Product Tag can accumulate credits from several Completed
        // lots logged at different times with different Colors to Produce
        // combinations (e.g. one batch "Blue-White / BCP", another
        // "Red-White / Black") — the aggregate above is color-blind by
        // design (see getReadyToDispatchData), but this keeps each color's
        // own produced/dispatched numbers so the operator can still tell
        // which color batch they're actually dispatching.
        colors: {} // colorLabel ('' = untagged/blank) -> { producedQty, dispatchedQty }
      };
    }

    map[key].producedQty += r.producedQty;
    map[key].dispatchedQty += r.consumedQty;

    const colorLabel = String(r.color || '').trim();
    if (!map[key].colors[colorLabel]) {
      map[key].colors[colorLabel] = { producedQty: 0, dispatchedQty: 0 };
    }
    map[key].colors[colorLabel].producedQty += r.producedQty;
    map[key].colors[colorLabel].dispatchedQty += r.consumedQty;
  });

  return map;
}

/**
 * Retrieves the "Ready to Dispatch" view: one row per Product ID with a
 * fully-packed, Product-tagged Warehouse Pool credit, with produced/
 * dispatched/ready quantities. The row itself stays one-per-product (a
 * Product Tag can legitimately span several Completed lots logged under
 * different Colors to Produce combinations) — colorBreakdown carries each
 * of those colors' own produced/dispatched/ready numbers alongside the
 * aggregate, for a detail view (see App.Dispatch.openColorBreakdown) rather
 * than splitting the main list into one row per color. Pure read — no
 * persistence.
 */
function getReadyToDispatchData() {
  try {
    const map = _computeReadyToDispatchMap();
    const records = Object.values(map).map(r => ({
      productId: r.productId,
      productName: r.productName,
      producedQty: r.producedQty,
      dispatchedQty: r.dispatchedQty,
      readyQty: r.producedQty - r.dispatchedQty,
      colorBreakdown: Object.keys(r.colors)
        .map(color => ({
          color,
          producedQty: r.colors[color].producedQty,
          dispatchedQty: r.colors[color].dispatchedQty,
          readyQty: r.colors[color].producedQty - r.colors[color].dispatchedQty
        }))
        .sort((a, b) => a.color.localeCompare(b.color))
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
/**
 * Maps one raw Dispatch sheet row into the record shape the client expects,
 * given its 2-indexed sheet row number, or null for a blank row (missing
 * dispatchNumber). Shared by getDispatchData's bulk read and saveDispatch's
 * single fresh-row read-back (used to patch just the saved record into the
 * client's already-loaded table in place instead of a full list reload).
 * @private
 */
function _mapDispatchRow(row, rowIdx) {
  const dispatchNumber = String(row[DISPATCH_COL.DISPATCH_NUMBER - 1] || '').trim();
  if (!dispatchNumber) return null;

  const rawDate = row[DISPATCH_COL.DISPATCH_DATE - 1];
  const dateStr = rawDate instanceof Date ? toSafeDateString(rawDate) : String(rawDate || '');

  return {
    rowIdx: rowIdx,
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
    logisticsCost: Number(row[DISPATCH_COL.LOGISTICS_COST - 1]) || 0,
    rate: Number(row[DISPATCH_COL.RATE - 1]) || 0,
    amount: Number(row[DISPATCH_COL.AMOUNT - 1]) || 0
  };
}

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
    ensureDispatchRateColumns(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, []);

    const data = sheet.getRange(2, 1, lastRow - 1, DISPATCH_COL.AMOUNT).getValues();
    const records = [];

    for (let i = 0; i < data.length; i++) {
      const record = _mapDispatchRow(data[i], i + 2);
      if (record) records.push(record);
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
 * (orderNumber, productId), optionally excluding every line belonging to
 * one Dispatch Number (the bill currently being edited, so its own
 * not-yet-rewritten lines aren't double-counted against themselves).
 * @private
 */
function _getDispatchedQtyForOrder(dispatchSheet, orderNumber, productId, excludeDispatchNumber) {
  const lastRow = dispatchSheet.getLastRow();
  if (lastRow < 2) return 0;

  const data = dispatchSheet.getRange(2, 1, lastRow - 1, DISPATCH_COL.QTY).getValues();
  const targetOrder = orderNumber.toLowerCase();
  const targetProduct = productId.toLowerCase();
  const excludeKey = excludeDispatchNumber ? String(excludeDispatchNumber).trim().toLowerCase() : null;
  let total = 0;

  data.forEach(row => {
    if (excludeKey && String(row[DISPATCH_COL.DISPATCH_NUMBER - 1] || '').trim().toLowerCase() === excludeKey) return;
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
    ensureDispatchRateColumns(sheet);

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

    let lines;
    try {
      lines = typeof formData.lines === 'string' ? JSON.parse(formData.lines) : (formData.lines || []);
    } catch (e) {
      return buildResponse(false, null, 'Invalid item line data format.');
    }
    if (!Array.isArray(lines) || lines.length === 0) {
      return buildResponse(false, null, 'A dispatch bill must contain at least one item.');
    }

    const cleanLines = [];
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i] || {};
      const productId = sanitizeString(line.productId || '', 'productId');
      const productName = sanitizeString(line.productName || '', 'productName');
      if (!productId || !productName) continue;

      const qty = validateNumber(line.qty, 0.001, 10000000);
      if (qty <= 0) continue;

      const rate = validateNumber(line.rate, 0, 10000000);

      cleanLines.push({ productId, productName, qty, rate });
    }

    if (cleanLines.length === 0) {
      return buildResponse(false, null, 'Please specify at least one valid item with a Product and Quantity greater than zero.');
    }

    const isEdit = !!formData.dispatchNumber;
    let dispatchNumber = isEdit ? sanitizeString(formData.dispatchNumber, 'dispatchNumber') : '';

    // This bill's own previously-saved per-product quantities, captured
    // BEFORE any write so they can be added back onto the Ready-to-Dispatch
    // availability check below — otherwise editing a bill would fail
    // against its own prior consumption (the Warehouse Pool snapshot
    // _computeReadyToDispatchMap() reads from still reflects the old rows
    // until recalculateWarehousePool() runs, further down, after the write).
    const originalQtyByProduct = {};
    if (isEdit) {
      const lastRow = sheet.getLastRow();
      let found = false;
      if (lastRow >= 2) {
        const existingData = sheet.getRange(2, 1, lastRow - 1, DISPATCH_COL.QTY).getValues();
        existingData.forEach(row => {
          if (String(row[DISPATCH_COL.DISPATCH_NUMBER - 1] || '').trim().toLowerCase() === dispatchNumber.toLowerCase()) {
            found = true;
            const pid = String(row[DISPATCH_COL.PRODUCT_ID - 1] || '').trim().toLowerCase();
            if (pid) originalQtyByProduct[pid] = (originalQtyByProduct[pid] || 0) + (Number(row[DISPATCH_COL.QTY - 1]) || 0);
          }
        });
      }
      if (!found) {
        return buildResponse(false, null, `Dispatch "${dispatchNumber}" not found.`);
      }
    } else {
      dispatchNumber = getNextDispatchNumber();
    }

    // Validate against currently available Ready to Dispatch quantity,
    // reserved cumulatively across this bill's own lines (two lines of the
    // same product in one bill both draw from the same pool).
    // _computeReadyToDispatchMap() keys untagged final-stage rows under a
    // '__output__' prefixed key (see there) to keep them from colliding with
    // a differently-named Product Tag, but the productId the client echoes
    // back is always the unprefixed display value for both cases — so an
    // untagged product's lookup must fall back to the prefixed key.
    const readyMap = _computeReadyToDispatchMap();
    const reservedByProduct = {};
    for (let i = 0; i < cleanLines.length; i++) {
      const line = cleanLines[i];
      const key = line.productId.toLowerCase();
      const entry = readyMap[key] || readyMap['__output__' + key];
      const currentReadyQty = entry ? (entry.producedQty - entry.dispatchedQty) : 0;
      const availableQty = currentReadyQty + (originalQtyByProduct[key] || 0);
      const reservedSoFar = reservedByProduct[key] || 0;

      if (reservedSoFar + line.qty > availableQty + 0.0001) {
        return buildResponse(false, null, `Only ${availableQty} unit(s) of "${line.productName}" are Ready to Dispatch (requested ${reservedSoFar + line.qty} across this bill).`);
      }
      reservedByProduct[key] = reservedSoFar + line.qty;
    }

    // Validate against the SPECIFIC PI/Estimate line's own remaining qty, not
    // just aggregate product stock — without this, two different orders for
    // the same product share one pool with no guard, letting one dispatch
    // over-fulfill a small order's line using stock meant for a different
    // order of the same product. Only checked when an Order Number is
    // actually referenced; a "Direct" dispatch (no orderNumber) has no order
    // line to check against. This bill's own existing rows are excluded from
    // "already dispatched" (they're about to be replaced), and reservations
    // are cumulative across this bill's own lines the same way as above.
    const linesWithNoOrderMatch = [];
    if (orderNumber) {
      const reservedForOrderByProduct = {};
      for (let i = 0; i < cleanLines.length; i++) {
        const line = cleanLines[i];
        const orderLineQty = _getClientOrderLineQty(orderNumber, line.productId);
        if (orderLineQty === null) {
          linesWithNoOrderMatch.push(line.productName);
          continue;
        }
        const key = line.productId.toLowerCase();
        const alreadyDispatchedForOrder = _getDispatchedQtyForOrder(sheet, orderNumber, line.productId, isEdit ? dispatchNumber : null);
        const availableForOrder = orderLineQty - alreadyDispatchedForOrder;
        const reservedSoFar = reservedForOrderByProduct[key] || 0;

        if (reservedSoFar + line.qty > availableForOrder + 0.0001) {
          return buildResponse(false, null,
            `Only ${availableForOrder} unit(s) of "${line.productName}" remain pending on PI/Estimate "${orderNumber}" (ordered ${orderLineQty}, already dispatched ${alreadyDispatchedForOrder} against it). Use a different order reference, or Direct, if this dispatch is really for other stock.`);
        }
        reservedForOrderByProduct[key] = reservedSoFar + line.qty;
      }
    }

    // Snapshot the logistics contractor's rate at save time, same pattern
    // as Production's contractor payable. Optional — 0 if no rate card entry.
    // Applied per-line (rate x that line's own qty) so the bill's per-row
    // Logistics Cost values sum correctly to rate x total bill qty, instead
    // of repeating one bill-wide total on every row.
    const logisticsRate = typeof _getContractorRate === 'function'
      ? _getContractorRate(logisticsContractor, LOGISTICS_PROCESS_NAME)
      : 0;

    if (isEdit) {
      deleteRowsById(dispatchNumber, sheet, 2, DISPATCH_COL.DISPATCH_NUMBER);
    }

    const rowsToWrite = cleanLines.map(line => [
      dispatchNumber,
      dateStr,
      orderNumber,
      clientName,
      line.productId,
      line.productName,
      line.qty,
      transport,
      remarks,
      invoiceNumber,
      privateMark,
      grNumber,
      logisticsContractor,
      logisticsRate,
      logisticsRate * line.qty,
      line.rate,
      line.qty * line.rate
    ]);

    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);

    if (typeof recalculateWarehousePool === 'function') {
      recalculateWarehousePool();
    }

    SpreadsheetApp.flush();

    const totalQty = cleanLines.reduce((sum, l) => sum + l.qty, 0);
    const logMsg = isEdit
      ? `Dispatch bill updated: ${dispatchNumber} (${cleanLines.length} item(s), total qty ${totalQty})`
      : `Dispatch bill recorded: ${dispatchNumber} (${cleanLines.length} item(s), total qty ${totalQty})`;

    logAction(isEdit ? 'UPDATE' : 'CREATE', APP_CONFIG.SHEETS.DISPATCH, dispatchNumber, logMsg, 'SUCCESS');

    let successMsg = isEdit ? 'Dispatch bill updated successfully.' : 'Dispatch bill recorded successfully.';

    // Soft-validate the optional Order Number reference (warn, don't block) —
    // reuses the per-line lookups above instead of re-reading the sheet.
    // A product landing here means no line for THAT product was found under
    // that order number at all (order number itself may still exist with
    // other products on it, or may not exist/may have been removed).
    if (orderNumber && linesWithNoOrderMatch.length > 0) {
      successMsg += ` Note: PI / Estimate "${orderNumber}" has no line for ${linesWithNoOrderMatch.map(n => `"${n}"`).join(', ')} (it may have been edited or removed).`;
    }

    // Read this bill's own just-written rows back (cheap — only its own
    // block, not the whole sheet) so the client can patch it into an
    // already-loaded Dispatch table in place instead of a full
    // getDispatchData() reload.
    const freshRawRows = sheet.getRange(startRow, 1, rowsToWrite.length, DISPATCH_COL.AMOUNT).getValues();
    const freshRows = freshRawRows.map((row, i) => _mapDispatchRow(row, startRow + i));

    return buildResponse(true, { dispatchNumber: dispatchNumber, rows: freshRows }, successMsg);
  } catch (error) {
    Log.error('[saveDispatch] Error:', error.message);
    logAction('ERROR', 'saveDispatch', formData.dispatchNumber || 'NEW', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to save dispatch: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes every line item of one dispatch bill (identified by Dispatch
 * Number).
 */
function deleteDispatch(dispatchNumber, expectedItemCount, expectedTotalQty) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(DISPATCH_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.DISPATCH);
    if (!sheet) throw new Error('Dispatch sheet not found.');

    const dispatchClean = sanitizeString(dispatchNumber, 'dispatchNumber');

    // Optional guard (same intent as the old rowIdx-based mismatch check):
    // skip the delete, rather than silently removing the wrong bill, if the
    // bill's item count/total qty has changed since the client loaded it.
    if (expectedItemCount !== undefined && expectedTotalQty !== undefined) {
      const lastRow = sheet.getLastRow();
      let itemCount = 0;
      let totalQty = 0;
      if (lastRow >= 2) {
        const data = sheet.getRange(2, 1, lastRow - 1, DISPATCH_COL.QTY).getValues();
        data.forEach(row => {
          if (String(row[DISPATCH_COL.DISPATCH_NUMBER - 1] || '').trim().toLowerCase() === dispatchClean.toLowerCase()) {
            itemCount++;
            totalQty += Number(row[DISPATCH_COL.QTY - 1]) || 0;
          }
        });
      }
      if (itemCount !== Number(expectedItemCount) || Math.abs(totalQty - Number(expectedTotalQty)) > 0.0001) {
        return buildResponse(false, null, 'Data mismatch: The bill has been modified since it was loaded. Please refresh.');
      }
    }

    const rowsDeleted = deleteRowsById(dispatchClean, sheet, 2, DISPATCH_COL.DISPATCH_NUMBER);
    if (rowsDeleted === 0) {
      return buildResponse(false, null, `Dispatch "${dispatchClean}" not found.`);
    }

    if (typeof recalculateWarehousePool === 'function') {
      recalculateWarehousePool();
    }

    SpreadsheetApp.flush();

    const msg = `Dispatch "${dispatchClean}" deleted (${rowsDeleted} item(s) removed).`;
    logAction('DELETE', APP_CONFIG.SHEETS.DISPATCH, dispatchClean, msg, 'SUCCESS');

    return buildResponse(true, null, msg);
  } catch (error) {
    Log.error('[deleteDispatch] Error:', error.message);
    logAction('ERROR', 'deleteDispatch', String(dispatchNumber), error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete dispatch: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes multiple dispatch bills (every line item of each) in a single
 * batch.
 * @param {Array<string>} dispatchNumbers
 */
function deleteDispatchBulk(dispatchNumbers) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(DISPATCH_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.DISPATCH);
    if (!sheet) throw new Error('Dispatch sheet not found.');

    const requested = (dispatchNumbers || []).map(d => String(d || '').trim()).filter(Boolean);
    if (requested.length === 0) {
      return buildResponse(true, null, 'No dispatch bills selected.');
    }

    const targetSet = new Set(requested);
    const { rowsDeleted } = _rewriteWithoutMatchingRowsBulk(sheet, 2, DISPATCH_COL.DISPATCH_NUMBER, targetSet);

    if (typeof recalculateWarehousePool === 'function') {
      recalculateWarehousePool();
    }

    SpreadsheetApp.flush();

    const msg = `Deleted ${requested.length} dispatch bill(s) (${rowsDeleted} item(s) removed).`;
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
