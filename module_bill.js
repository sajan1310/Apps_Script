/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_bill.gs — BILL LEDGER MODULE
 * 
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Complete CRUD operations for Bill Ledger including:
 * - Retrieval and aggregation (getBillData)
 * - Creation (saveBill)
 * - Deletion (deleteBill)
 * - GST calculation and formatting
 * - Document-level locking for concurrency safety
 * - Comprehensive date parsing and formatting
 * 
 * Key Features:
 * ───────────────────────────────────────────────────────────────────────────
 * - Structured item data (not pre-formatted strings)
 * - Accurate GST computation (percentage-based)
 * - Date parsing with multiple format support
 * - Formula injection prevention
 * - Atomic operations with rollback support
 * - Detailed error messages for user feedback
 * 
 * Dependencies:
 * ───────────────────────────────────────────────────────────────────────────
 * - config.gs (APP_CONFIG, BILL_COL, column mappings)
 * - utils.gs (getSheet, buildResponse, sanitizeString, toSafeDateString)
 * - module_po.gs (shared validation functions)
 * 
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ─────────────────────────────────────────────────────────────────────────
// CONSTANTS & CONFIGURATION
// ─────────────────────────────────────────────────────────────────────────

/**
 * Bill lock timeout (milliseconds)
 * Prevents deadlocks and concurrent modification conflicts
 * Timeout: 15 seconds (sufficient for most operations)
 */
const BILL_LOCK_TIMEOUT_MS = 15000;

/**
 * Default GST rate (percentage)
 * Used when item doesn't specify GST rate
 * India: 18% standard rate
 */
const DEFAULT_GST_RATE_PCT = 18;

// ─────────────────────────────────────────────────────────────────────────
// VALIDATION HELPERS
// ─────────────────────────────────────────────────────────────────────────

/**
 * Validates that a bill number is non-empty and safe
 * 
 * @param {*} billNumber - Bill number to validate
 * @returns {string} Validated bill number
 * @throws {Error} If invalid
 * 
 * @example
 * const billNum = validateBillNumber('INV-998');  // "INV-998"
 * const billNum = validateBillNumber('');         // throws
 */
function _validateBillNumber(billNumber) {
  const s = String(billNumber || '').trim();
  if (!s) {
    throw new Error('Bill number must not be empty.');
  }

  // Reuse injection check from utils
  sanitizeString(s, 'billNumber');

  return s;
}

/**
 * Validates that a value is a valid non-negative number
 * 
 * @param {*} value - Value to validate
 * @param {string} fieldLabel - Field name for error message
 * @param {boolean} allowZero - Whether zero is acceptable
 * @returns {number} Validated number
 * @throws {Error} If invalid
 */
function _toValidNumber(value, fieldLabel, allowZero = true) {
  const n = Number(value);

  if (!Number.isFinite(n)) {
    throw new Error(`${fieldLabel} must be a valid number.`);
  }

  if (!allowZero && n === 0) {
    throw new Error(`${fieldLabel} must be greater than zero.`);
  }

  if (n < 0) {
    throw new Error(`${fieldLabel} must not be negative.`);
  }

  return n;
}

// ─────────────────────────────────────────────────────────────────────────
// GST COMPUTATION
// ─────────────────────────────────────────────────────────────────────────

/**
 * Computes GST amount and line total from quantity, unit price, and GST rate
 * 
 * Formula:
 * - Subtotal = Qty × Price
 * - GST Amount = Subtotal × (GST% / 100)
 * - Line Total = Subtotal + GST Amount
 * 
 * Important Contract:
 * - gstRatePct MUST be a percentage (e.g., 18 for 18%)
 * - NOT a decimal fraction (e.g., NOT 0.18)
 * - This prevents the bug where 18% was added as ₹18 flat
 * 
 * @param {number} qty - Quantity (non-negative)
 * @param {number} price - Unit price (non-negative)
 * @param {number} gstRatePct - GST rate as percentage (e.g., 18 for 18%)
 * @returns {Object} Result object
 *   @returns {number} subtotal - Pre-tax subtotal
 *   @returns {number} gstAmount - GST rupee amount (computed)
 *   @returns {number} lineTotal - Total including GST
 * 
 * @example
 * // Qty: 10, Price: ₹100, GST: 18%
 * const result = computeLineTotals(10, 100, 18);
 * // Result:
 * // {
 * //   subtotal: 1000,       (10 × 100)
 * //   gstAmount: 180,       (1000 × 0.18)
 * //   lineTotal: 1180       (1000 + 180)
 * // }
 */
function _computeLineTotals(qty, price, gstRatePct) {
  try {
    const qty_val = _toValidNumber(qty, 'Quantity');
    const price_val = _toValidNumber(price, 'Price');
    const gstRate_val = _toValidNumber(gstRatePct, 'GST rate', true);  // Allow 0% GST

    if (gstRate_val < 0 || gstRate_val > 100) {
      throw new Error('GST rate must be between 0 and 100 percent.');
    }

    const subtotal = qty_val * price_val;
    const gstAmount = subtotal * (gstRate_val / 100);
    const lineTotal = subtotal + gstAmount;

    return {
      subtotal: parseFloat(subtotal.toFixed(2)),
      gstAmount: parseFloat(gstAmount.toFixed(2)),
      lineTotal: parseFloat(lineTotal.toFixed(2))
    };
  } catch (error) {
    Log.error('[_computeLineTotals] Error:', error.message);
    throw error;
  }
}

/**
 * Validates GST rate is within acceptable range
 * 
 * @param {number} gstRate - GST rate as percentage
 * @returns {number} Validated rate
 * @throws {Error} If invalid
 */
function _validateGstRate(gstRate) {
  const rate = _toValidNumber(gstRate, 'GST rate', true);
  if (rate < 0 || rate > 100) {
    throw new Error('GST rate must be between 0 and 100 percent.');
  }
  return rate;
}

// ────────────────────────────────────────────────────────��────────────────
// BILL DATA RETRIEVAL
// ─────────────────────────────────────────────────────────────────────────

/**
 * getBillData()
 * 
 * Retrieves all bills from sheet and aggregates by bill number
 * 
 * Data Transformation:
 * 1. Fetch all rows from Bill sheet
 * 2. Group rows by bill number
 * 3. Aggregate items under each bill
 * 4. Recompute totals from item data (never from pre-aggregated columns)
 * 5. Sort by bill date (newest first)
 * 
 * Key Design Principles:
 * - Structured items (not pre-formatted strings)
 * - Totals always recomputed (single source of truth = item data)
 * - Robust date parsing (multiple formats)
 * - Handles out-of-order rows gracefully
 * - Logs unparseable dates for debugging
 * 
 * Performance:
 * - Single batch API call for all data
 * - O(n) aggregation algorithm
 * - Returns only what frontend needs
 * 
 * @returns {Object} API response
 *   - success: true/false
 *   - data: Array of bill objects (sorted by date descending)
 *   - message: Status message
 * 
 * Bill Object Structure:
 * {
 *   poNumber: "1001",              // Reference to PO (may be "DIRECT")
 *   billNumber: "INV-998",         // Unique bill identifier
 *   billDate: "25/05/2026",        // Display format
 *   billDateRaw: "2026-05-25T...", // ISO format
 *   vendor: "Hero MotoCorp",       // Vendor name
 *   contact: "Ludhiana",           // Contact location
 *   remarks: "Included GST",       // Bill remarks
 *   items: [
 *     {
 *       name: "Brake Pads",
 *       size: "Standard",
 *       qty: 50,
 *       price: 150,
 *       gstRatePct: 18,
 *       lineTotal: 8850  // Qty × Price × (1 + GST%)
 *     }
 *   ],
 *   totalQty: 50,                  // Sum of item quantities
 *   totalAmount: 8850              // Sum of line totals (with GST)
 * }
 */
function getBillData() {
  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.BILL);
    if (!sheet) {
      return buildResponse(false, null, 'Bill sheet not found.');
    }

    _ensureBillSheetColumns(sheet);

    const startRow = APP_CONFIG.BILL_SETTINGS.DATA_START_ROW;
    const lastRow = sheet.getLastRow();

    if (lastRow < startRow) {
      return buildResponse(true, []);
    }

    // Read from PO_NUMBER column through AFFECTS_STOCK column
    // Avoids reading unrelated columns to the left
    const firstCol = BILL_COL.PO_NUMBER;
    const numCols = BILL_COL.AFFECTS_STOCK - firstCol + 1;
    const data = sheet
      .getRange(startRow, firstCol, lastRow - startRow + 1, numCols)
      .getValues();

    // Calculate zero-based offsets relative to firstCol
    const OFF = {
      poNum:    BILL_COL.PO_NUMBER - firstCol,
      billNum:  BILL_COL.BILL_NUMBER - firstCol,
      billDate: BILL_COL.BILL_DATE - firstCol,
      vendor:   BILL_COL.VENDOR - firstCol,
      contact:  BILL_COL.CONTACT - firstCol,
      itemName: BILL_COL.ITEM_NAME - firstCol,
      size:     BILL_COL.SIZE - firstCol,
      narration: BILL_COL.NARRATION - firstCol,
      qty:      BILL_COL.QTY - firstCol,
      unit:     BILL_COL.UNIT - firstCol,
      price:    BILL_COL.PRICE - firstCol,
      gst:      BILL_COL.GST - firstCol,
      remarks:  BILL_COL.REMARKS - firstCol,
      baseQty:  BILL_COL.BASE_QTY - firstCol,
      baseRate: BILL_COL.BASE_RATE - firstCol,
      affectsStock: BILL_COL.AFFECTS_STOCK - firstCol
    };

    // Aggregate rows into bill objects
    const billMap = {};

    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const billNum = String(r[OFF.billNum] || '').trim();

      if (!billNum) continue;

      // The same bill number can legitimately recur across different vendors,
      // so bills are aggregated by the (vendor, bill number) pair, not the
      // bill number alone — otherwise two different vendors' bills sharing a
      // number would incorrectly merge into a single bill object here.
      const vendorName = String(r[OFF.vendor] || '').trim();
      const key = vendorName.toLowerCase() + '|' + billNum.toLowerCase();

      // Create bill entry if first time seeing this (vendor, bill number) pair
      if (!billMap[key]) {
        const rawDate = r[OFF.billDate];
        const parsedDate = toSafeDateString(rawDate); // returns DD/MM/YYYY

        // Log unparseable dates for debugging
        if (!parsedDate && rawDate) {
          Log.warn(
            `[getBillData] Unparseable date in bill #${billNum}: "${rawDate}"`
          );
        }

        // Convert DD/MM/YYYY to YYYY-MM-DD for HTML date inputs and date sorting
        const isoDateStr = parsedDate
          ? parsedDate.split('/').reverse().join('-')
          : '';

        billMap[key] = {
          poNumber: String(r[OFF.poNum] || ''),
          poNumbers: [],
          billNumber: billNum,
          billDate: parsedDate || (rawDate ? String(rawDate) : 'N/A'),
          billDateRaw: isoDateStr,
          vendor: vendorName,
          contact: String(r[OFF.contact] || ''),
          remarks: String(r[OFF.remarks] || ''),
          items: [],  // Structured, not pre-formatted
          totalQty: 0,
          totalAmount: 0,
          _rowIdx: i
        };
      }

      // Track unique PO numbers for this bill
      const rowPoNum = String(r[OFF.poNum] || '').trim();
      if (rowPoNum && !billMap[key].poNumbers.includes(rowPoNum)) {
        billMap[key].poNumbers.push(rowPoNum);
      }

      // Parse item data
      const qty = _toValidNumber(r[OFF.qty], 'Qty');
      const price = _toValidNumber(r[OFF.price], 'Price');
      const gstRatePct = _validateGstRate(r[OFF.gst]);

      // Compute line total from item data
      const { lineTotal } = _computeLineTotals(qty, price, gstRatePct);

      // Add structured item (not pre-formatted string)
      billMap[key].items.push({
        name: String(r[OFF.itemName] || ''),
        size: String(r[OFF.size] || ''),
        narration: String(r[OFF.narration] || ''),
        unit: String(r[OFF.unit] || 'Pcs'),
        qty: qty,
        price: price,
        gstRatePct: gstRatePct,
        lineTotal: lineTotal,
        poNumber: rowPoNum,
        baseQty: _toValidNumber(r[OFF.baseQty], 'Base Qty', true) || qty,
        baseRate: _toValidNumber(r[OFF.baseRate], 'Base Rate', true) || price,
        // Blank (legacy rows predating this column) defaults to affecting Stock.
        affectsStock: String(r[OFF.affectsStock] || 'Y').trim().toUpperCase() !== 'N'
      });

      // Update totals (always recomputed, never from sheet columns)
      billMap[key].totalQty += qty;
      billMap[key].totalAmount += lineTotal;
    }

    // Convert to array and sort by bill date (newest first), then by row
    // order (newest first) to break ties between same-day bills so the most
    // recently entered one still comes first rather than falling back to
    // ascending sheet order.
    const sorted = Object.values(billMap).sort(function(a, b) {
      const timeA = a.billDateRaw ? new Date(a.billDateRaw).getTime() : 0;
      const timeB = b.billDateRaw ? new Date(b.billDateRaw).getTime() : 0;
      if (timeB !== timeA) return timeB - timeA;  // Descending (newest first)
      return b._rowIdx - a._rowIdx;
    });
    sorted.forEach(function(b) { delete b._rowIdx; });

    return buildResponse(true, sorted);
  } catch (error) {
    Log.error('[getBillData] Error:', error.message);
    logAction('ERROR', 'getBillData', 'module_bill', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to load bills: ' + error.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────
// BILLED-QTY AGGREGATION (used by PO auto-matching)
// ─────────────────────────────────────────────────────────────────────────

/**
 * _buildPoLineKey(poNumber, name, size, narration)
 *
 * The ONE canonical key for aggregating/looking up billed quantity against a
 * PO line. _aggregateBilledBaseQtyByPo() (below) builds this map; module_po.js's
 * _attachPoStatus/suggestPoAllocations and module_dashboard.js's
 * _getOpenPoSummary all look it up — every caller MUST derive an identical
 * key for the same (PO, item, size, narration), or status/pending-qty math
 * silently breaks. Previously hand-built in all 4 places (drift risk); now
 * every part (including a missing/undefined/null one) normalizes the same
 * way: coerced to '', then trimmed and lowercased.
 *
 * @param {string} poNumber
 * @param {string} name
 * @param {string} size
 * @param {string} narration
 * @returns {string}
 */
function _buildPoLineKey(poNumber, name, size, narration) {
  const part = v => String(v == null ? '' : v).trim().toLowerCase();
  return part(poNumber) + '|' + part(name) + '|' + part(size) + '|' + part(narration);
}

/**
 * _aggregateBilledBaseQtyByPo()
 *
 * Sums already-billed base quantity per (PO_NUMBER, item name, size, narration)
 * across the entire Bill Ledger. Used to compute how much of a PO line is
 * still "open" when suggesting PO matches for a new, unlinked bill line
 * (see suggestPoAllocations() in module_po.js).
 *
 * @param {Array} [preloadedBillRecords] - getBillData().data, when the caller
 *   already has it (e.g. getDashboardData loads the Bill Ledger once for its
 *   own KPIs) — skips a second full Bill-sheet read. Falls back to reading
 *   it here when omitted, unchanged for existing callers.
 * @returns {Object} map keyed by _buildPoLineKey(poNumber, name, size, narration) -> baseQty
 */
function _aggregateBilledBaseQtyByPo(preloadedBillRecords) {
  const map = {};
  let billRecords;
  if (preloadedBillRecords) {
    billRecords = preloadedBillRecords;
  } else {
    const billResponse = getBillData();
    if (!billResponse.success) return map;
    billRecords = billResponse.data || [];
  }

  (billRecords || []).forEach(function(bill) {
    (bill.items || []).forEach(function(item) {
      const poNum = String(item.poNumber || '').trim();
      if (!poNum || poNum.toUpperCase() === 'DIRECT') return;

      const key = _buildPoLineKey(poNum, item.name, item.size, item.narration);

      map[key] = (map[key] || 0) + (Number(item.baseQty) || 0);
    });
  });

  return map;
}

/**
 * _computeBillOverageWarnings(items)
 *
 * Advisory, non-blocking check run AFTER a bill has been saved: for every
 * bill line linked to a real PO (not "DIRECT"/blank), re-aggregates the
 * Bill Ledger fresh from the sheet (so it reflects this save's own effect)
 * and flags any PO line whose cumulative billed base qty now exceeds what
 * was ordered. Bills are intentionally allowed to exceed a PO line (freight
 * corrections, renegotiated quantities, etc.) — this only surfaces the
 * overage in the save's response message rather than leaving it invisible
 * (see module_po.js#_attachPoStatus, which no longer clamps pendingQty to 0).
 *
 * @param {Array} items - the same items array saveBill() just wrote, each
 *   already carrying .poNumber/.baseQty (set during saveBill's own item map).
 * @returns {string[]} human-readable warning lines, empty if nothing is over.
 */
function _computeBillOverageWarnings(items) {
  try {
    const poNumbers = new Set();
    (items || []).forEach(function(it) {
      const po = String(it.poNumber || '').trim();
      if (po && po.toUpperCase() !== 'DIRECT') poNumbers.add(po);
    });
    if (poNumbers.size === 0) return [];

    const billedMap = _aggregateBilledBaseQtyByPo();
    const poResponse = typeof getPOData === 'function' ? getPOData(billedMap) : null;
    if (!poResponse || !poResponse.success) return [];

    const poByNumber = new Map();
    (poResponse.data || []).forEach(function(po) { poByNumber.set(String(po.poNumber), po); });

    const warnings = [];
    const seenKeys = new Set();
    (items || []).forEach(function(it) {
      const poNum = String(it.poNumber || '').trim();
      if (!poNum || poNum.toUpperCase() === 'DIRECT') return;

      const key = _buildPoLineKey(poNum, it.name, it.size || '', it.narration || '');
      if (seenKeys.has(key)) return;
      seenKeys.add(key);

      const po = poByNumber.get(poNum);
      if (!po) return; // PO no longer exists — nothing to check against

      const poItem = (po.items || []).find(function(pi) {
        return _buildPoLineKey(poNum, pi.name, pi.size, pi.narration) === key;
      });
      if (!poItem) return; // this line doesn't match any existing PO line

      const billedBaseQty = billedMap[key] || 0;
      const orderedBaseQty = poItem.baseQty || 0;
      if (billedBaseQty > orderedBaseQty + 0.0001) {
        const overBy = billedBaseQty - orderedBaseQty;
        warnings.push(
          `"${it.name}" on PO ${poNum} is now billed ${billedBaseQty.toFixed(2)} vs ordered ${orderedBaseQty.toFixed(2)} (${overBy.toFixed(2)} over).`
        );
      }
    });
    return warnings;
  } catch (e) {
    // Advisory only — never let this check block or fail a bill save.
    return [];
  }
}

// ─────────────────────────────────────────────────────────────────────────
// BILL CREATION
// ─────────────────────────────────────────────────────────────────────────

/**
 * saveBill(formData)
 * 
 * Records a new bill to the sheet
 * 
 * Operations:
 * 1. Acquire document lock (concurrency safety)
 * 2. Validate all input data
 * 3. Parse and validate bill date
 * 4. Validate items array (must have at least one item)
 * 5. Compute GST and line totals for each item
 * 6. Build sheet rows
 * 7. Append rows to sheet
 * 8. Flush changes
 * 
 * GST Contract:
 * - gstRatePct MUST be a percentage (e.g., 18 for 18%)
 * - NOT a decimal (e.g., NOT 0.18)
 * - Computed as: LineTotal = (Qty × Price) × (1 + GST% / 100)
 * 
 * Date Handling:
 * - Accepts multiple formats: DD/MM/YYYY, YYYY-MM-DD, Date objects
 * - Requires valid date (not optional)
 * - Stored as native Date object in sheet
 * 
 * @param {Object} formData - Form submission data
 *   @param {string} formData.poNumber - Reference to PO (may be "DIRECT")
 *   @param {string} formData.billNumber - Unique bill identifier
 *   @param {string} formData.billDate - Date (DD/MM/YYYY or YYYY-MM-DD)
 *   @param {string} formData.vendor - Vendor name
 *   @param {string} formData.contact - Contact location
 *   @param {string} formData.remarks - Optional remarks
 *   @param {Array|string} formData.items - Items array or JSON string
 *     Each item: { name, size, unit, qty, price, gst }
 * 
 * @returns {Object} API response
 *   - success: true/false
 *   - data: { billNumber } if successful
 *   - message: Status/error message
 */
function saveBill(formData) {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(BILL_LOCK_TIMEOUT_MS)) {
    return buildResponse(
      false,
      null,
      'System is busy. Please try again in a moment.'
    );
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.BILL);
    if (!sheet) {
      throw new Error('Bill sheet not found.');
    }

    _ensureBillSheetColumns(sheet);

    // ─────────────────────────────────────────────────────────────────
    // Parse and validate items
    // ─────────────────────────────────────────────────────────────────
    let items;
    try {
      items = typeof formData.items === 'string'
        ? JSON.parse(formData.items)
        : (formData.items || []);
    } catch (error) {
      return buildResponse(
        false,
        null,
        'Invalid items data: could not parse JSON.'
      );
    }

    if (!Array.isArray(items) || items.length === 0) {
      return buildResponse(
        false,
        null,
        'Cannot save a bill with zero items. Add at least one item.'
      );
    }

    // ─────────────────────────────────────────────────────────────────
    // Validate identifiers & check duplicates
    // ─────────────────────────────────────────────────────────────────
    const oldBillNum = String(formData.existingBillNumber || '').trim();
    const oldVendor = String(formData.existingVendor || '').trim();
    const isEdit = !!oldBillNum;

    const billNumber = _validateBillNumber(formData.billNumber);
    const vendor = sanitizeString(formData.vendor, 'vendor');
    const startRow = APP_CONFIG.BILL_SETTINGS.DATA_START_ROW;

    // Uniqueness is on the (vendor, bill number) pair — the same bill number
    // can legitimately recur across different vendors.
    if (_billExists(sheet, startRow, vendor, billNumber, isEdit ? oldVendor : null, isEdit ? oldBillNum : null)) {
      return buildResponse(
        false,
        null,
        isEdit
          ? `Bill #${billNumber} for vendor "${vendor}" already exists in another record. Edit aborted.`
          : `Bill #${billNumber} already exists for vendor "${vendor}". Please use a unique bill number for this vendor.`
      );
    }

    let firstMatchRow = -1;
    if (isEdit) {
      firstMatchRow = _findFirstBillRow(sheet, startRow, oldVendor, oldBillNum);
      if (firstMatchRow === -1) {
        throw new Error(
          `Original Bill #${oldBillNum} for vendor "${oldVendor}" not found in sheet. Edit aborted to prevent data corruption.`
        );
      }
    }

    // Parse poNumbers array (multi-PO support); fall back to legacy poNumber field.
    // There's no longer a mandatory top-level PO selection — each item carries
    // its own PO (assigned by auto-matching or a manual override on the bill
    // form), so this only matters as a fallback for items that somehow don't.
    let poNumbersArray;
    try {
      const raw = formData.poNumbers || formData.poNumber || '[]';
      if (Array.isArray(raw)) {
        poNumbersArray = raw;
      } else if (String(raw).startsWith('[')) {
        poNumbersArray = JSON.parse(raw);
      } else {
        poNumbersArray = String(raw).split(',').map(s => s.trim()).filter(Boolean);
      }
    } catch (e) {
      poNumbersArray = formData.poNumber ? [String(formData.poNumber)] : [];
    }
    const defaultPoNumber = sanitizeString(poNumbersArray[0] || 'DIRECT', 'poNumber');

    // ─────────────────────────────────────────────────────────────────
    // Parse and validate bill date
    // ─────────────────────────────────────────────────────────────────
    const billDateNative = toSafeDateObject(formData.billDate);
    if (!billDateNative) {
      return buildResponse(
        false,
        null,
        'Invalid bill date. Accepted formats: DD/MM/YYYY or YYYY-MM-DD.'
      );
    }

    // ─────────────────────────────────────────────────────────────────
    // Sanitize header fields
    // ─────────────────────────────────────────────────────────────────
    const contact = sanitizeString(formData.contact, 'contact');
    const remarks = sanitizeString(formData.remarks, 'remarks');

    // ─────────────────────────────────────────────────────────────────
    // Build batch write array
    // ─────────────────────────────────────────────────────────────────
    const itemUnitMap = _getItemUnitInfoMap();
    const unitsMap = _getUnitsMap();

    // Items the user chose "Ledger only" for in the checkStockAdjustmentConflicts
    // warning (see Script_Core.html's billForm submit handler) — these lines are
    // still saved and shown in the Bill Ledger, but excluded from Stock's Billed
    // Qty sum (see _getBilledAndConsumedQtyMaps in module_stock.js) because the
    // goods they represent were already reflected in a later manual stock count.
    let excludeFromStockKeys = [];
    try {
      const rawExclude = formData.excludeFromStockKeys;
      excludeFromStockKeys = typeof rawExclude === 'string'
        ? (rawExclude ? JSON.parse(rawExclude) : [])
        : (Array.isArray(rawExclude) ? rawExclude : []);
    } catch (e) {
      excludeFromStockKeys = [];
    }
    const excludeFromStockSet = new Set(excludeFromStockKeys.map(k => String(k).trim().toLowerCase()));

    const newRows = items.map(function(item) {
      const qty = _toValidNumber(item.qty, 'Qty');
      const price = _toValidNumber(item.price, 'Price');
      // `|| DEFAULT_GST_RATE_PCT` would also override an intentional 0%
      // (GST-exempt) line since 0 is falsy — only fall back when the field
      // was actually left unset.
      const gstRatePct = _validateGstRate(item.gst != null && item.gst !== '' ? item.gst : DEFAULT_GST_RATE_PCT);
      const itemPoNumber = sanitizeString(item.po || item.poNumber || defaultPoNumber, 'poNumber');
      const unit = sanitizeString(item.unit || 'Pcs', 'item.unit');

      const { lineTotal } = _computeLineTotals(qty, price, gstRatePct);

      // An unconvertible unit must never block saving the bill — fall back
      // to treating the as-entered qty/price as already in the Base Unit.
      // The line still saves with whatever unit the user actually typed.
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

      // Attach converted values onto the item so autoExtractFromPoOrBill()
      // (called below with this same `items` array) can use the per-Base-Unit
      // rate instead of the as-entered price, and log this line's actual PO.
      item.baseQty = baseQty;
      item.baseRate = baseRate;
      item.poNumber = itemPoNumber;

      const itemKey = String(item.name || '').trim().toLowerCase() + '|' + String(item.size || '').trim().toLowerCase();
      const affectsStock = excludeFromStockSet.has(itemKey) ? 'N' : 'Y';

      return [
        itemPoNumber,                                      // 1: PO_NUMBER
        billNumber,                                        // 2: BILL_NUMBER
        billDateNative,                                    // 3: BILL_DATE
        vendor,                                            // 4: VENDOR
        contact,                                           // 5: CONTACT
        sanitizeString(item.name, 'item.name'),            // 6: ITEM_NAME
        sanitizeString(item.size || '', 'item.size'),      // 7: SIZE
        sanitizeString(item.narration || '', 'item.narration'), // 8: NARRATION
        qty,                                               // 9: QTY
        unit,                                              // 10: UNIT
        price,                                             // 11: PRICE
        gstRatePct,                                        // 12: GST (rate %)
        lineTotal,                                         // 13: TOTAL
        remarks,                                           // 14: REMARKS
        baseQty,                                           // 15: BASE_QTY
        baseRate,                                          // 16: BASE_RATE
        affectsStock                                       // 17: AFFECTS_STOCK
      ];
    });

    // ─────────────────────────────────────────────────────────────────
    // Append to sheet or update
    // ─────────────────────────────────────────────────────────────────
    if (isEdit) {
      _rewriteBillsWithoutMatchingRows(sheet, startRow, oldVendor, oldBillNum);
      _insertBillRowsAt(sheet, firstMatchRow, newRows);
    } else {
      const appendRow = sheet.getLastRow() + 1;
      sheet.getRange(
        appendRow,
        BILL_COL.PO_NUMBER,
        newRows.length,
        newRows[0].length
      ).setValues(newRows);
    }

    // Auto-extract vendors, items, and rates
    autoExtractFromPoOrBill(formData.vendor, formData.contact, items, {
      date: billDateNative,
      poNumber: defaultPoNumber,
      billNumber: billNumber
    });

    if (typeof recalculateStock === 'function') {
      recalculateStock();
    }

    SpreadsheetApp.flush();

    let msg = isEdit
      ? `Bill #${billNumber} updated successfully.`
      : `Bill #${billNumber} recorded successfully.`;

    const overageWarnings = _computeBillOverageWarnings(items);
    if (overageWarnings.length > 0) {
      msg += ' Warning: ' + overageWarnings.join(' ');
    }

    logAction(isEdit ? 'UPDATE' : 'CREATE', APP_CONFIG.SHEETS.BILL, billNumber, `Items: ${items.length}`, 'SUCCESS');

    return buildResponse(true, { billNumber }, msg);
  } catch (error) {
    Log.error('[saveBill] Error:', error.message);
    logAction('ERROR', 'saveBill', formData.billNumber, error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to save bill: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────────────────
// BILL DELETION
// ─────────────────────────────────────────────────────────────────────────

/**
 * deleteBill(vendor, billNumber)
 *
 * Deletes all rows for a given (vendor, bill number) pair. Vendor is required
 * because bill numbers are only unique per-vendor — matching on bill number
 * alone could delete a different vendor's bill that happens to share it.
 *
 * Operations:
 * 1. Acquire document lock
 * 2. Validate vendor + bill number
 * 3. Remove all rows from Bill sheet matching the (vendor, bill number) pair
 * 4. Flush changes
 *
 * Safety:
 * - Document lock prevents concurrent deletes
 * - Validates vendor + bill number before deletion
 * - Uses efficient batch deletion algorithm
 * - Audit log entry for each delete
 *
 * @param {string} vendor - Vendor of the bill to delete
 * @param {string|number} billNumber - Bill number to delete
 *
 * @returns {Object} API response
 *   - success: true/false
 *   - data: null
 *   - message: Status/error message
 */
function deleteBill(vendor, billNumber) {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(BILL_LOCK_TIMEOUT_MS)) {
    return buildResponse(
      false,
      null,
      'System is busy. Please try again.'
    );
  }

  try {
    // Validate vendor + bill number
    const validBillNumber = _validateBillNumber(billNumber);
    const validVendor = sanitizeString(vendor, 'vendor');

    const sheet = getSheet(APP_CONFIG.SHEETS.BILL);
    if (!sheet) {
      throw new Error('Bill sheet not found.');
    }

    // Delete all rows matching the (vendor, bill number) pair
    _rewriteBillsWithoutMatchingRows(
      sheet,
      APP_CONFIG.BILL_SETTINGS.DATA_START_ROW,
      validVendor,
      validBillNumber
    );

    if (typeof recalculateStock === 'function') {
      recalculateStock();
    }

    SpreadsheetApp.flush();

    const msg = `Bill #${validBillNumber} deleted successfully.`;
    logAction('DELETE', APP_CONFIG.SHEETS.BILL, validBillNumber, msg, 'SUCCESS');

    return buildResponse(true, null, msg);
  } catch (error) {
    Log.error('[deleteBill] Error:', error.message);
    logAction('ERROR', 'deleteBill', billNumber, error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete bill: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * deleteBillsBulk(bills)
 *
 * Deletes all rows for multiple (vendor, bill number) pairs in a single batch
 * operation. Mirrors deleteBill() but rewrites the sheet once instead of once
 * per bill.
 *
 * @param {Array<{vendor: string, billNumber: string}>} bills - Bills to delete
 * @returns {Object} API response
 */
function deleteBillsBulk(bills) {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(BILL_LOCK_TIMEOUT_MS)) {
    return buildResponse(
      false,
      null,
      'System is busy. Please try again.'
    );
  }

  try {
    const pairs = (bills || [])
      .map(b => ({
        vendor: String((b && b.vendor) || '').trim(),
        billNumber: String((b && b.billNumber) || '').trim()
      }))
      .filter(p => p.billNumber);

    if (pairs.length === 0) {
      return buildResponse(true, null, 'No bills selected.');
    }

    const targetSet = new Set(
      pairs.map(p => p.vendor.toLowerCase() + '|' + p.billNumber.toLowerCase())
    );

    const sheet = getSheet(APP_CONFIG.SHEETS.BILL);
    if (!sheet) {
      throw new Error('Bill sheet not found.');
    }

    const { rowsDeleted, deletedKeys } = _rewriteBillsWithoutMatchingRowsBulk(
      sheet,
      APP_CONFIG.BILL_SETTINGS.DATA_START_ROW,
      targetSet
    );

    if (typeof recalculateStock === 'function') {
      recalculateStock();
    }

    SpreadsheetApp.flush();

    const msg = `Deleted ${pairs.length} bill(s) (${rowsDeleted} row(s) removed).`;
    logAction('BULK_DELETE', APP_CONFIG.SHEETS.BILL, 'multiple', msg, 'SUCCESS');

    return buildResponse(true, { deletedKeys }, msg);
  } catch (error) {
    Log.error('[deleteBillsBulk] Error:', error.message);
    logAction('ERROR', 'deleteBillsBulk', 'multiple', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete bills: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────────────────
// HELPER: Display date formatting
// ─────────────────────────────────────────────────────────────────────────

/**
 * Converts an ISO date string (YYYY-MM-DD) to display format (DD/MM/YYYY).
 * Named distinctly from module_po.js's _toDisplayDate (which accepts Date
 * objects and is timezone-aware) to avoid a cross-file declaration collision.
 *
 * @param {string} isoDateString - ISO date (YYYY-MM-DD)
 * @returns {string} Display date (DD/MM/YYYY)
 *
 * @private
 */
function _isoToDisplayDate(isoDateString) {
  if (!isoDateString) return 'N/A';

  const parts = String(isoDateString).split('-');
  if (parts.length !== 3) return String(isoDateString);

  // Convert YYYY-MM-DD to DD/MM/YYYY
  return `${parts[2]}/${parts[1]}/${parts[0]}`;
}

/**
 * Self-healing migration: ensures the Bill sheet contains the "Narration" column.
 * If not present, inserts it at column 8 (before Qty) and labels it "Narration".
 * 
 * @param {Sheet} sheet - The Google Sheet object for Bill Ledger
 * @private
 */
function _ensureBillSheetColumns(sheet) {
  const cache = CacheService.getScriptCache();
  if (cache.get('bill_columns_ok') === '1') return;

  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const hasNarration = headers.some(h => String(h).trim().toLowerCase() === 'narration');

  if (!hasNarration) {
    sheet.insertColumnBefore(BILL_COL.NARRATION);
    sheet.getRange(1, BILL_COL.NARRATION).setValue('Narration');
    SpreadsheetApp.flush();
  }

  // Re-read headers — Narration insert above may have shifted columns
  const refreshedLastCol = sheet.getLastColumn();
  const refreshedHeaders = refreshedLastCol >= 1
    ? sheet.getRange(1, 1, 1, refreshedLastCol).getValues()[0]
    : [];
  const hasBaseQty = refreshedHeaders.some(h => String(h).trim().toLowerCase() === 'base qty');

  if (!hasBaseQty) {
    sheet.getRange(1, BILL_COL.BASE_QTY, 1, 2).setValues([['Base Qty', 'Base Rate']]);
    SpreadsheetApp.flush();
  }

  const hasAffectsStock = refreshedHeaders.some(h => String(h).trim().toLowerCase() === 'affects stock');
  if (!hasAffectsStock) {
    sheet.getRange(1, BILL_COL.AFFECTS_STOCK).setValue('Affects Stock');
    SpreadsheetApp.flush();
  }

  // Cache for 6 hours — columns won't disappear between calls
  cache.put('bill_columns_ok', '1', 21600);
}

/**
 * Checks if a (vendor, bill number) pair already exists in the sheet.
 * The same bill number can legitimately recur across different vendors, so
 * uniqueness is enforced on the compound key, not the bill number alone.
 * Can exclude a specific (vendor, bill number) pair (useful during edits).
 *
 * @param {Sheet} sheet - Bill ledger sheet
 * @param {number} startRow - Start row
 * @param {string} vendor - Vendor name to check
 * @param {string} billNumber - Bill number to check
 * @param {string} [excludeVendor] - Vendor of the pair to exclude from check
 * @param {string} [excludeBillNumber] - Bill number of the pair to exclude from check
 * @returns {boolean} True if duplicate exists
 * @private
 */
function _billExists(sheet, startRow, vendor, billNumber, excludeVendor = null, excludeBillNumber = null) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return false;

  const numCols = BILL_COL.VENDOR - BILL_COL.BILL_NUMBER + 1;
  const values = sheet
    .getRange(startRow, BILL_COL.BILL_NUMBER, lastRow - startRow + 1, numCols)
    .getValues();
  const vendorOff = BILL_COL.VENDOR - BILL_COL.BILL_NUMBER;

  const targetBill = String(billNumber).trim().toLowerCase();
  const targetVendor = String(vendor).trim().toLowerCase();
  const excludeBill = excludeBillNumber ? String(excludeBillNumber).trim().toLowerCase() : null;
  const excludeVend = excludeVendor ? String(excludeVendor).trim().toLowerCase() : null;

  for (let i = 0; i < values.length; i++) {
    const rowBill = String(values[i][0] || '').trim().toLowerCase();
    const rowVendor = String(values[i][vendorOff] || '').trim().toLowerCase();
    if (rowBill === targetBill && rowVendor === targetVendor) {
      if (excludeBill !== null && rowBill === excludeBill && rowVendor === excludeVend) {
        continue;
      }
      return true;
    }
  }
  return false;
}

/**
 * Finds the first row number matching a (vendor, bill number) pair.
 *
 * @param {Sheet} sheet - Bill ledger sheet
 * @param {number} startRow - Start row
 * @param {string} vendor - Vendor name to search for
 * @param {string} billNumber - Bill number to search for
 * @returns {number} Row index (1-based) or -1 if not found
 * @private
 */
function _findFirstBillRow(sheet, startRow, vendor, billNumber) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return -1;

  const numCols = BILL_COL.VENDOR - BILL_COL.BILL_NUMBER + 1;
  const values = sheet
    .getRange(startRow, BILL_COL.BILL_NUMBER, lastRow - startRow + 1, numCols)
    .getValues();
  const vendorOff = BILL_COL.VENDOR - BILL_COL.BILL_NUMBER;

  const targetBill = String(billNumber).trim().toLowerCase();
  const targetVendor = String(vendor).trim().toLowerCase();

  for (let i = 0; i < values.length; i++) {
    if (
      String(values[i][0] || '').trim().toLowerCase() === targetBill &&
      String(values[i][vendorOff] || '').trim().toLowerCase() === targetVendor
    ) {
      return startRow + i;
    }
  }
  return -1;
}

/**
 * Removes all rows matching a specific (vendor, bill number) pair from the sheet.
 * Uses the same rewrite strategy as module_po to ensure safety and speed.
 *
 * @param {Sheet} sheet - Bill ledger sheet
 * @param {number} startRow - Start row
 * @param {string} vendor - Vendor name to match
 * @param {string} billNumber - Bill number to delete
 * @private
 */
function _rewriteBillsWithoutMatchingRows(sheet, startRow, vendor, billNumber) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < startRow) return;

  const numRows = lastRow - startRow + 1;
  const allRows = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
  const billIndex = BILL_COL.BILL_NUMBER - 1;
  const vendorIndex = BILL_COL.VENDOR - 1;
  const targetBill = String(billNumber).trim().toLowerCase();
  const targetVendor = String(vendor).trim().toLowerCase();

  const kept = allRows.filter(function(row) {
    const rowBill = String(row[billIndex] || '').trim().toLowerCase();
    const rowVendor = String(row[vendorIndex] || '').trim().toLowerCase();
    return !(rowBill === targetBill && rowVendor === targetVendor);
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

/**
 * Removes all rows matching any of a set of (vendor, bill number) pairs
 * from the sheet, in a single batch rewrite (bulk-delete counterpart to
 * _rewriteBillsWithoutMatchingRows).
 *
 * @param {Sheet} sheet - Bill ledger sheet
 * @param {number} startRow - Start row
 * @param {Set<string>} targetSet - Set of "vendorLower|billNumberLower" composite keys
 * @returns {{rowsDeleted: number, deletedKeys: string[]}} rowsDeleted count and
 *   the as-stored "vendor|billNumber" keys actually removed (original casing)
 * @private
 */
function _rewriteBillsWithoutMatchingRowsBulk(sheet, startRow, targetSet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < startRow) return { rowsDeleted: 0, deletedKeys: [] };

  const numRows = lastRow - startRow + 1;
  const allRows = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
  const billIndex = BILL_COL.BILL_NUMBER - 1;
  const vendorIndex = BILL_COL.VENDOR - 1;
  const matchedKeys = new Set();

  const kept = allRows.filter(function(row) {
    const billVal = String(row[billIndex] || '').trim();
    const vendorVal = String(row[vendorIndex] || '').trim();
    const lookupKey = vendorVal.toLowerCase() + '|' + billVal.toLowerCase();
    if (targetSet.has(lookupKey)) {
      matchedKeys.add(vendorVal + '|' + billVal);
      return false;
    }
    return true;
  });

  sheet.getRange(startRow, 1, numRows, lastCol).clearContent();

  if (kept.length > 0) {
    sheet.getRange(startRow, 1, kept.length, lastCol).setValues(kept);
  }

  const rowsToDelete = numRows - kept.length;
  if (rowsToDelete > 0) {
    sheet.deleteRows(startRow + kept.length, rowsToDelete);
  }

  return { rowsDeleted: rowsToDelete, deletedKeys: Array.from(matchedKeys) };
}

/**
 * Inserts rows at the specified row index, shifting subsequent rows down.
 * 
 * @param {Sheet} sheet - Bill ledger sheet
 * @param {number} rowIndex - Row index to insert at
 * @param {Array} rows - 2D array of row data
 * @private
 */
function _insertBillRowsAt(sheet, rowIndex, rows) {
  if (!rows || rows.length === 0) return;

  const currentLastRow = sheet.getLastRow();
  if (rowIndex <= currentLastRow) {
    sheet.insertRows(rowIndex, rows.length);
  }

  sheet.getRange(rowIndex, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * Rewrites Bill Ledger rows referencing a renamed/merged item, so historical
 * bills stay attributed to the item's current name + size.
 *
 * Runs under the caller's existing document lock — does not acquire its own.
 *
 * @param {string} oldName
 * @param {string} oldSize
 * @param {string} newName
 * @param {string} newSize
 */
function backfillBillItemRefs(oldName, oldSize, newName, newSize) {
  const sheet = getSheet(APP_CONFIG.SHEETS.BILL);
  if (!sheet) return;

  const startRow = APP_CONFIG.BILL_SETTINGS.DATA_START_ROW;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  const numRows = lastRow - startRow + 1;
  const range = sheet.getRange(startRow, BILL_COL.ITEM_NAME, numRows, 2);
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
