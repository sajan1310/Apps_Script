/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_po.gs — PURCHASE ORDER MODULE
 * 
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Complete CRUD operations for Purchase Orders including:
 * - Retrieval and aggregation (getPOData)
 * - Creation and updates (savePO)
 * - Deletion with cascade (deletePO)
 * - Strict data validation and normalization
 * - Document-level locking for concurrency safety
 * - Comprehensive date parsing and formatting
 * 
 * Key Features:
 * ───────────────────────────────────────────────────────────────────────────
 * - Configurable ID policy (prefix, start number, padding)
 * - Timezone-aware date handling
 * - Formula injection prevention
 * - Atomic operations with rollback support
 * - Efficient row operations (bulk rewrite, batch insert)
 * - Detailed error messages for user feedback
 * 
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ─────────────────────────────────────────────────────────────────────────
// CONSTANTS & CONFIGURATION
// ─────────────────────────────────────────────────────────────────────────

/**
 * Document lock timeout (milliseconds)
 * Prevents deadlocks and concurrent modification conflicts
 * Timeout: 15 seconds (sufficient for most operations)
 */
const PO_LOCK_TIMEOUT_MS = 15000;

/**
 * PO ID Generation Policy
 * Configurable to support different ID formats
 * 
 * Examples:
 * - { prefix: "PO-", start: 1001, pad: 0 } → "PO-1001", "PO-1002"
 * - { prefix: "", start: 1001, pad: 0 } → "1001", "1002"
 * - { prefix: "PO", start: 1, pad: 4 } → "PO0001", "PO0002"
 */
const PO_ID_POLICY = Object.freeze({
  prefix: '',      // Optional prefix (e.g., "PO-", "INV-")
  start: 1001,     // Starting number
  pad: 0           // Left-pad with zeros (0 = no padding)
});

// ─────────────────────────────────────────────────────────────────────────
// TIMEZONE & SYSTEM UTILITIES
// ─────────────────────────────────────────────────────────────────────────

/**
 * Gets the spreadsheet's timezone
 * Falls back to script timezone if unavailable
 * 
 * @returns {string} Timezone string (e.g., "America/New_York")
 */
function _getSpreadsheetTimeZone() {
  try {
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    if (tz) return tz;
  } catch (error) {
    console.warn('[_getSpreadsheetTimeZone] Error getting spreadsheet TZ:', error);
  }

  try {
    return Session.getScriptTimeZone();
  } catch (error) {
    console.warn('[_getSpreadsheetTimeZone] Error getting script TZ:', error);
    return 'UTC';  // Final fallback
  }
}

/**
 * Acquires document-level lock for mutation operations
 * Prevents concurrent modification and data corruption
 * 
 * @returns {Lock} Document lock object
 * @throws {Error} If lock service unavailable
 */
function _getDocumentMutationLock() {
  const lock = LockService.getDocumentLock();
  if (!lock) {
    throw new Error('Document lock service unavailable. Cannot safely modify data.');
  }
  return lock;
}

// ─────────────────────────────────────────────────────────────────────────
// TEXT NORMALIZATION & VALIDATION
// ─────────────────────────────────────────────────────────────────────────

/**
 * Normalizes text for sheet cells
 * 
 * Prevents:
 * - Formula injection (=, @, +, - prefixes)
 * - Excess whitespace
 * - Null/undefined issues
 * 
 * @param {*} value - Value to normalize (any type)
 * @returns {string} Normalized string
 * 
 * @example
 * normalizeTextCell('=SUM(A1:A10)') // Returns: "'=SUM(A1:A10)"
 * normalizeTextCell('  hello  ')     // Returns: "hello"
 * normalizeTextCell(null)             // Returns: ""
 */
function _normalizeTextCell(value) {
  let s = String(value == null ? '' : value).trim();

  // Escape formula injection patterns
  if (/^[=@+\-]/.test(s)) {
    s = "'" + s;
  }

  return s;
}

/**
 * Validates that a text value is non-empty after normalization
 * 
 * @param {*} value - Value to validate
 * @param {string} fieldLabel - Field name for error message
 * @returns {string} Normalized non-empty string
 * @throws {Error} If value is empty after normalization
 * 
 * @example
 * const name = requireNonEmptyText(userInput, 'Item name');
 */
function _requireNonEmptyText(value, fieldLabel) {
  const s = _normalizeTextCell(value);
  if (!s) {
    throw new Error(`${fieldLabel} must not be empty.`);
  }
  return s;
}

/**
 * Validates a PO number
 * @param {*} poNumber - PO number to validate
 * @returns {string} Validated PO number
 * @throws {Error} If invalid
 */
function _validatePoNumber(poNumber) {
  return _requireNonEmptyText(poNumber, 'PO number');
}

/**
 * Converts value to finite number with fallback
 * @param {*} value - Value to convert
 * @param {number} fallback - Value to use if conversion fails
 * @returns {number} Converted number or fallback
 */
function _toFiniteNumber(value, fallback) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

/**
 * Validates that value is a non-negative number
 * @param {*} value - Value to validate
 * @param {string} fieldLabel - Field name for error message
 * @returns {number} Validated non-negative number
 * @throws {Error} If invalid or negative
 */
function _toNonNegativeNumber(value, fieldLabel) {
  const n = Number(value);
  if (!Number.isFinite(n) || n < 0) {
    throw new Error(`${fieldLabel} must be a valid non-negative number.`);
  }
  return n;
}

// ─────────────────────────────────────────────────────────────────────────
// DATE ENGINE
// ─────────────────────────────────────────────────────────────────────────

/**
 * Converts any date input to safe ISO date string (YYYY-MM-DD)
 * 
 * Accepts:
 * - Date objects
 * - ISO strings (YYYY-MM-DD)
 * - Display formats (DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY)
 * - Timestamps (milliseconds)
 * 
 * Returns:
 * - YYYY-MM-DD string (ISO 8601)
 * - null for invalid/unsupported input
 * 
 * Design:
 * - Intentionally does NOT use new Date(string) to avoid ambiguity
 * - Validates date parts (month 1-12, day 1-31)
 * - Handles timezone correctly via Google Sheets timezone
 * 
 * @param {Date|number|string|null} dateVal - Date value to convert
 * @returns {string|null} ISO date string or null if invalid
 * 
 * @example
 * toSafeDateString(new Date('2026-05-23'))  // "2026-05-23"
 * toSafeDateString('2026-05-23')            // "2026-05-23"
 * toSafeDateString('23/05/2026')            // "2026-05-23"
 * toSafeDateString('23-05-2026')            // "2026-05-23"
 * toSafeDateString(1716422400000)           // "2026-05-23" (if correct timestamp)
 * toSafeDateString(null)                    // null
 * toSafeDateString('invalid')               // null
 */
function _toSafeDateString(dateVal) {
  // Early exit for null/empty
  if (dateVal == null || dateVal === '') {
    return null;
  }

  const tz = _getSpreadsheetTimeZone();

  // Handle Date objects
  if (dateVal instanceof Date) {
    // Validate date is valid
    if (isNaN(dateVal.getTime())) {
      return null;
    }

    // Format using timezone
    return Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd');
  }

  const str = String(dateVal).trim();

  // Pattern 1: ISO format (YYYY-MM-DD)
  let m = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const y = Number(m[1]);
    const mo = Number(m[2]);
    const d = Number(m[3]);

    if (_isValidDateParts(y, mo, d)) {
      return `${m[1]}-${m[2]}-${m[3]}`;
    }

    return null;
  }

  // Pattern 2: Display formats (DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY)
  m = str.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);
  if (m) {
    const d = Number(m[1]);
    const mo = Number(m[2]);
    const y = Number(m[3]);

    if (_isValidDateParts(y, mo, d)) {
      return `${String(y)}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    }

    return null;
  }

  // No recognized format
  return null;
}

/**
 * Validates date parts (year, month, day)
 * Checks for valid ranges and handles month/day boundaries correctly
 * 
 * @param {number} year - Full year (e.g., 2026)
 * @param {number} month - Month 1-12
 * @param {number} day - Day 1-31
 * @returns {boolean} True if valid date
 * 
 * @private
 */
function _isValidDateParts(year, month, day) {
  // All must be integers
  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) {
    return false;
  }

  // Basic range checks
  if (month < 1 || month > 12 || day < 1 || day > 31) {
    return false;
  }

  // Create date and validate (catches invalid dates like Feb 30)
  const dt = new Date(year, month - 1, day);
  return (
    dt.getFullYear() === year &&
    dt.getMonth() === month - 1 &&
    dt.getDate() === day
  );
}

/**
 * Converts any date to display format (DD/MM/YYYY)
 * 
 * @param {Date|number|string|null} dateVal - Date value to convert
 * @returns {string} Display date string (DD/MM/YYYY) or original value if unparseable
 * 
 * @example
 * toDisplayDate(new Date('2026-05-23'))  // "23/05/2026"
 * toDisplayDate('2026-05-23')            // "23/05/2026"
 * toDisplayDate('23/05/2026')            // "23/05/2026"
 * toDisplayDate(null)                    // "N/A"
 */
function _toDisplayDate(dateVal) {
  const tz = _getSpreadsheetTimeZone();

  // Handle Date objects
  if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
    return Utilities.formatDate(dateVal, tz, 'dd/MM/yyyy');
  }

  // Try to convert to safe format first
  const safe = _toSafeDateString(dateVal);
  if (!safe) {
    return dateVal ? String(dateVal) : 'N/A';
  }

  // Convert YYYY-MM-DD to DD/MM/YYYY
  const parts = safe.split('-');
  return `${parts[2]}/${parts[1]}/${parts[0]}`;
}

// ─────────────────────────────────────────────────────────────────────────
// ID GENERATION
// ─────────────────────────────────────────────────────────────────────────

/**
 * Escapes special regex characters
 * @param {string} s - String to escape
 * @returns {string} Escaped string safe for regex
 * @private
 */
function _escapeRegex(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Generates the next sequential PO number based on existing data
 * 
 * Algorithm:
 * 1. Fetch all PO numbers in ID column
 * 2. Extract numeric portion matching the configured prefix
 * 3. Find maximum numeric value
 * 4. Return next number with proper formatting
 * 
 * Handles:
 * - Empty sheets (returns start number)
 * - Mixed formats (extracts regex match)
 * - Padding/zero-fill (configurable)
 * - Prefixes (e.g., "PO-", "INV-")
 * 
 * @param {Sheet} sheet - Google Sheet object
 * @param {number} startRow - First data row (1-based)
 * @param {number} idColIndex - Column containing IDs (1-based)
 * @param {Object} policy - ID policy configuration
 *   @param {string} policy.prefix - ID prefix (e.g., "PO-")
 *   @param {number} policy.start - Starting number
 *   @param {number} policy.pad - Zero-padding width
 * @returns {string} Next PO number
 * 
 * @example
 * const policy = { prefix: "PO-", start: 1001, pad: 0 };
 * const nextId = generateNextId(sheet, 5, 2, policy);
 * // Returns: "PO-1042" (if 1041 was last)
 */
function generateNextId(sheet, startRow, idColIndex, policy) {
  const cfg = policy || PO_ID_POLICY;
  const prefix = String(cfg.prefix || '');
  const start = Number(cfg.start || 1001);
  const pad = Number(cfg.pad || 0);

  const lastRow = sheet.getLastRow();

  // Empty sheet
  if (lastRow < startRow) {
    return prefix + String(start).padStart(pad, '0');
  }

  // Fetch all IDs
  const values = sheet
    .getRange(startRow, idColIndex, lastRow - startRow + 1, 1)
    .getValues();

  // Build regex to match format: {prefix}{numeric}
  const re = new RegExp('^' + _escapeRegex(prefix) + '(\\d+)$');
  let maxId = start - 1;

  // Find maximum ID
  for (let i = 0; i < values.length; i++) {
    const val = String(values[i][0] || '').trim();
    if (!val) continue;

    const match = val.match(re);
    if (!match) continue;

    const num = parseInt(match[1], 10);
    if (Number.isFinite(num) && num > maxId) {
      maxId = num;
    }
  }

  // Return next ID with formatting
  return prefix + String(maxId + 1).padStart(pad, '0');
}

// ─────────────────────────────────────────────────────────────────────────
// ITEM & ROW MAPPING
// ─────────────────────────────────────────────────────────────────────────

/**
 * Normalizes and validates items array
 * 
 * Validates:
 * - Items is array with at least one item
 * - Each item has required fields (name, qty, price)
 * - Values are valid (non-negative quantities/prices)
 * 
 * Normalizes:
 * - Text fields (trims, prevents injection)
 * - Numbers (converts to finite values)
 * - Units (defaults to "Pcs")
 * 
 * @param {Array|string} rawItems - Items to normalize
 *   - If string: parsed as JSON
 *   - If array: used directly
 * @returns {Array} Normalized items array
 * @throws {Error} If validation fails
 * 
 * @example
 * const items = normalizeItems([
 *   { name: 'Widget', qty: 10, price: 5.50, unit: 'Pcs' }
 * ]);
 */
function _normalizeItems(rawItems) {
  let items;

  // Parse JSON if string
  if (typeof rawItems === 'string') {
    try {
      items = JSON.parse(rawItems);
    } catch (e) {
      throw new Error('Invalid items data: could not parse JSON.');
    }
  } else {
    items = rawItems || [];
  }

  // Validate is array with items
  if (!Array.isArray(items) || items.length === 0) {
    throw new Error('Cannot save a PO with zero items.');
  }

  // Build unit lookup maps once for the whole batch, not per row
  const itemUnitMap = _getItemUnitInfoMap();
  const unitsMap = _getUnitsMap();

  // Normalize each item
  return items.map((item, idx) => {
    const name = _requireNonEmptyText(item.name, `Item name at row ${idx + 1}`);
    const narration = _normalizeTextCell(item.narration);
    const size = _normalizeTextCell(item.size);
    const qty = _toNonNegativeNumber(item.qty, `Qty at row ${idx + 1}`);
    const unit = _normalizeTextCell(item.unit || 'Pcs');
    const price = _toNonNegativeNumber(item.price, `Price at row ${idx + 1}`);

    // An unconvertible unit (e.g. a brand-new family mismatch with no
    // Weight-per-Base-Unit set yet) must never block saving the PO — fall
    // back to treating the as-entered qty/price as already in the Base
    // Unit. The line still saves with the unit the user actually typed.
    const unitInfo = _lookupItemUnitInfo(itemUnitMap, name, size);
    let baseQty = qty;
    let baseRate = price;
    try {
      baseQty = convertQtyToBaseUnit(qty, unit, unitInfo, unitsMap);
      baseRate = convertRateToBaseUnit(price, unit, unitInfo, unitsMap);
    } catch (e) {
      baseQty = qty;
      baseRate = price;
    }

    return {
      name,
      narration,
      size,
      qty,
      unit,
      price,
      total: qty * price,
      baseQty,
      baseRate
    };
  });
}

/**
 * Builds a complete PO sheet row for a single item
 * 
 * Row format matches PO_COL column indices exactly
 * 
 * @param {Object} header - PO header information
 * @param {Object} item - Item line data
 * @returns {Array} Row array matching sheet columns
 * 
 * @private
 */
function _buildPoSheetRow(header, item) {
  return [
    header.poNumber,                     // Column A (PO number)
    header.poDate,                       // Column B (PO date)
    header.vendor,                       // Column C (Vendor)
    header.contact,                      // Column D (Contact)
    item.name,                           // Column E (Item name)
    item.narration,                      // Column F (Narration)
    item.size,                           // Column G (Size)
    item.qty,                            // Column H (Quantity)
    item.unit,                           // Column I (Unit)
    header.poDescription,                // Column J (PO description)
    header.poRemarks,                    // Column K (PO remarks)
    item.price,                          // Column L (Price)
    item.total,                          // Column M (Line total)
    header.totalQty,                     // Column N (Total qty)
    header.supplierRemarks,              // Column O (Supplier remarks)
    item.baseQty,                        // Column P (Qty converted to item's Base Unit)
    item.baseRate                        // Column Q (Rate converted to item's Base Unit)
  ];
}

/**
 * Creates PO header object from form data and items
 * Converts ISO date string back to native Date for Google Sheets
 * 
 * @param {Object} formData - Form submission data
 * @param {string} poNumber - PO number to assign
 * @param {string} poDateString - ISO date (YYYY-MM-DD) or null
 * @param {Array} items - Normalized items array
 * @returns {Object} PO header object
 * 
 * @private
 */
function _makePoHeader(formData, poNumber, poDateString, items) {
  // Convert ISO date string to native Date object for Google Sheets
  let nativeDate = new Date();

  if (poDateString) {
    const parts = poDateString.split('-');
    // Date constructor: Year, Month (0-indexed), Day
    nativeDate = new Date(parts[0], parts[1] - 1, parts[2]);
  }

  return {
    poNumber: poNumber,
    poDate: nativeDate,  // Native Date object for sheet storage
    vendor: _requireNonEmptyText(formData.vendor, 'Vendor'),
    contact: _normalizeTextCell(formData.contact),
    poDescription: _normalizeTextCell(formData.poDescription),
    poRemarks: _normalizeTextCell(formData.poRemarks),
    supplierRemarks: _normalizeTextCell(formData.supplierRemarks),
    totalQty: items.reduce(function(sum, item) {
      return sum + item.qty;
    }, 0)
  };
}

// ─────────────────────────────────────────────────────────────────────────
// SHEET REPOSITORY HELPERS
// ─────────────────────────────────────────────────────────────────────────

/**
 * Fetches all PO data rows from sheet in single batch call
 * 
 * @param {Sheet} sheet - PO sheet object
 * @param {number} startRow - First data row (1-based)
 * @returns {Array} 2D array of PO data
 * 
 * @private
 */
function _getPoDataRangeValues(sheet, startRow) {
  const lastRow = sheet.getLastRow();

  if (lastRow < startRow) {
    return [];
  }

  // Must extend through BASE_RATE (not just SUPPLIER_REMARKS) — BASE_QTY/
  // BASE_RATE are the last two columns and getPOData() reads them for every
  // item; stopping short silently left them out of range and made every
  // consumer (status/pending-qty, suggestPoAllocations, dashboard's Open POs)
  // fall back to the as-entered qty/price for any item whose PO unit isn't
  // already its base unit.
  const numCols = PO_COL.BASE_RATE - PO_COL.PO_NUMBER + 1;
  return sheet
    .getRange(startRow, PO_COL.PO_NUMBER, lastRow - startRow + 1, numCols)
    .getValues();
}

/**
 * Finds first row in sheet matching a PO number
 * 
 * @param {Sheet} sheet - PO sheet object
 * @param {number} startRow - First row to search (1-based)
 * @param {string|number} poNumber - PO number to find
 * @returns {number} Row number (1-based) or -1 if not found
 * 
 * @private
 */
function _findFirstPoRow(sheet, startRow, poNumber) {
  const lastRow = sheet.getLastRow();

  if (lastRow < startRow) {
    return -1;
  }

  const values = sheet
    .getRange(startRow, PO_COL.PO_NUMBER, lastRow - startRow + 1, 1)
    .getValues();

  const target = String(poNumber).trim();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === target) {
      return startRow + i;
    }
  }

  return -1;
}

/**
 * Finds existing PO date in sheet for a given PO number
 * Used to preserve date during edits if new date not provided
 * 
 * @param {Sheet} sheet - PO sheet object
 * @param {number} startRow - First row to search (1-based)
 * @param {string|number} poNumber - PO number to find
 * @returns {Date|null} Date value or null if not found
 * 
 * @private
 */
function _findExistingPoDate(sheet, startRow, poNumber) {
  const lastRow = sheet.getLastRow();

  if (lastRow < startRow) {
    return null;
  }

  const values = sheet
    .getRange(startRow, PO_COL.PO_NUMBER, lastRow - startRow + 1, 2)
    .getValues();

  const target = String(poNumber).trim();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === target) {
      return values[i][1];
    }
  }

  return null;
}

/**
 * Efficiently rewrites sheet data region excluding matching rows
 * 
 * Algorithm:
 * 1. Read entire data region
 * 2. Filter out rows where matchCol equals target
 * 3. Clear original region
 * 4. Write back filtered rows
 * 5. Delete excess rows
 * 
 * More deterministic than repeated deleteRows() calls
 * 
 * @param {Sheet} sheet - Sheet to modify
 * @param {number} startRow - First data row (1-based)
 * @param {number} matchCol - Column to match on (1-based)
 * @param {string|number} target - Value to match and exclude
 * 
 * @private
 */
function _rewriteWithoutMatchingRows(sheet, startRow, matchCol, target) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < startRow) {
    return;
  }

  const numRows = lastRow - startRow + 1;
  const allRows = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
  const matchIndex = matchCol - 1;
  const targetStr = String(target).trim();

  // Filter out matching rows
  const kept = allRows.filter(function(row) {
    return String(row[matchIndex] || '').trim() !== targetStr;
  });

  // Clear entire region
  sheet.getRange(startRow, 1, numRows, lastCol).clearContent();

  // Write back kept rows
  if (kept.length > 0) {
    sheet.getRange(startRow, 1, kept.length, lastCol).setValues(kept);
  }

  // Delete excess rows
  const rowsToDelete = numRows - kept.length;
  if (rowsToDelete > 0) {
    sheet.deleteRows(startRow + kept.length, rowsToDelete);
  }
}

/**
 * Like _rewriteWithoutMatchingRows, but removes every row whose matchCol
 * value is present in `targetSet` (trimmed string match), in a single pass.
 *
 * @param {Sheet} sheet - Sheet to modify
 * @param {number} startRow - First data row (1-based)
 * @param {number} matchCol - Column to match on (1-based)
 * @param {Set<string>} targetSet - Set of values to exclude
 * @returns {number} Number of rows removed
 * @private
 */
function _rewriteWithoutMatchingRowsBulk(sheet, startRow, matchCol, targetSet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < startRow) {
    return { rowsDeleted: 0, deletedIds: [] };
  }

  const numRows = lastRow - startRow + 1;
  const allRows = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
  const matchIndex = matchCol - 1;
  const matchedIds = new Set();

  const kept = allRows.filter(function(row) {
    const id = String(row[matchIndex] || '').trim();
    if (targetSet.has(id)) {
      matchedIds.add(id);
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

  return { rowsDeleted: rowsToDelete, deletedIds: Array.from(matchedIds) };
}

/**
 * Inserts rows at specified position, shifting existing rows down
 * 
 * @param {Sheet} sheet - Sheet to modify
 * @param {number} rowIndex - Row to insert at (1-based)
 * @param {Array} rows - 2D array of rows to insert
 * 
 * @private
 */
function _insertRowsAt(sheet, rowIndex, rows) {
  if (!rows || rows.length === 0) {
    return;
  }

  const currentLastRow = sheet.getLastRow();

  // Insert blank rows to make space
  if (rowIndex <= currentLastRow) {
    sheet.insertRows(rowIndex, rows.length);
  }

  // Write data to inserted rows
  sheet.getRange(rowIndex, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * Appends rows to end of sheet
 * 
 * @param {Sheet} sheet - Sheet to modify
 * @param {Array} rows - 2D array of rows to append
 * 
 * @private
 */
function _appendRows(sheet, rows) {
  if (!rows || rows.length === 0) {
    return;
  }

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

// ─────────────────────────────────────────────────────────────────────────
// PO STATUS (derived from Bill Ledger receipts)
// ─────────────────────────────────────────────────────────────────────────

/**
 * Attaches a display status plus per-line received/pending qty to a PO
 * object, derived from how much of each line has actually been billed.
 * Reuses module_bill.js's own billed-qty aggregation (keyed the same way
 * _getOpenPoSummary in module_dashboard.js keys it) so "received" stays in
 * sync with however billing already defines it — a PO only reads
 * "Completed" here when every line is actually fully billed, not just when
 * the totals happen to net out.
 *
 * Mutates `po` in place: sets po.status and, per item, item.receivedQty /
 * item.pendingQty — both expressed in the item's own entered unit (not the
 * base unit billedMap is keyed in), so the ledger can show them directly.
 *
 * @param {Object} po - PO object with poNumber + items[] (as built by getPOData)
 * @param {Object} billedMap - from _aggregateBilledBaseQtyByPo()
 * @private
 */
function _attachPoStatus(po, billedMap) {
  const items = po.items || [];
  let anyBilled = false;
  let allLinesComplete = items.length > 0;

  items.forEach(function(item) {
    const key = _buildPoLineKey(po.poNumber, item.name, item.size, item.narration);
    const billedBaseQty = billedMap[key] || 0;
    const orderedBaseQty = item.baseQty || item.qty || 0;

    // Convert the billed base qty back to the unit this line was entered in,
    // using the same qty/baseQty ratio the line itself was converted with.
    const ratio = item.baseQty > 0 ? item.qty / item.baseQty : 1;
    item.receivedQty = billedBaseQty * ratio;
    item.pendingQty = Math.max(0, item.qty - item.receivedQty);

    if (billedBaseQty > 0.0001) anyBilled = true;
    if (orderedBaseQty - billedBaseQty > 0.0001) allLinesComplete = false;
  });

  po.status = anyBilled && allLinesComplete ? 'Completed'
    : anyBilled ? 'Partially Received'
    : 'PO Issued';
}

// ─────────────────────────────────────────────────────────────────────────
// getPOData
// ─────────────────────────────────────────────────────────────────────────

/**
 * getPOData()
 * 
 * Retrieves all Purchase Orders from sheet and aggregates by PO number
 * 
 * Data Transformation:
 * 1. Fetch all rows from PO sheet
 * 2. Group rows by PO number
 * 3. Aggregate items under each PO
 * 4. Calculate totals (grand total, total qty)
 * 5. Sort by PO number (descending)
 * 
 * Performance:
 * - Single batch API call for all data
 * - O(n) aggregation algorithm
 * - Returns only what frontend needs
 * 
 * @returns {Object} API response
 *   - success: true/false
 *   - data: Array of PO objects (sorted by number descending)
 *   - message: Status message
 * 
 * PO Object Structure:
 * {
 *   poNumber: "1001",
 *   poDate: "23/05/2026",
 *   poDateRaw: "2026-05-23",
 *   vendor: "Hero MotoCorp",
 *   contact: "Ludhiana",
 *   poDescription: "Monthly Restock",
 *   poRemarks: "Urgent",
 *   supplierRemarks: "Ref-123",
 *   items: [
 *     { name, narration, size, qty, unit, price, total, receivedQty, pendingQty }
 *   ],
 *   grandTotal: 12500,
 *   totalQty: 100,
 *   status: "PO Issued" | "Partially Received" | "Completed"
 * }
 *
 * @param {Object} [preloadedBilledMap] - Internal only: a map already built by
 *   _aggregateBilledBaseQtyByPo(), so a caller that needs both this PO list
 *   AND the billed map itself (see suggestPoAllocations) can share one Bill
 *   Ledger read instead of aggregating it twice. Public client calls always
 *   omit this and getPOData computes its own, unchanged from before.
 */
function getPOData(preloadedBilledMap) {
  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.PO);
    if (!sheet) {
      return buildResponse(false, null, 'PO sheet not found.');
    }

    const startRow = APP_CONFIG.PO_SETTINGS.DATA_START_ROW;
    const data = _getPoDataRangeValues(sheet, startRow);

    if (data.length === 0) {
      return buildResponse(true, []);
    }

    // Calculate column offsets relative to PO_NUMBER column
    const base = PO_COL.PO_NUMBER;
    const OFF = {
      poNum:      PO_COL.PO_NUMBER - base,
      poDate:     PO_COL.PO_DATE - base,
      vendor:     PO_COL.VENDOR - base,
      contact:    PO_COL.CONTACT - base,
      desc:       PO_COL.PO_DESCRIPTION - base,
      remarks:    PO_COL.PO_REMARKS - base,
      totalQty:   PO_COL.TOTAL_QTY - base,
      supRemarks: PO_COL.SUPPLIER_REMARKS - base,
      itemName:   PO_COL.ITEM_NAME - base,
      narration:  PO_COL.NARRATION - base,
      size:       PO_COL.SIZE - base,
      qty:        PO_COL.QTY - base,
      unit:       PO_COL.UNIT - base,
      price:      PO_COL.PRICE - base,
      baseQty:    PO_COL.BASE_QTY - base,
      baseRate:   PO_COL.BASE_RATE - base
    };

    // Aggregate rows into PO objects
    const poMap = {};

    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const poNum = String(r[OFF.poNum] || '').trim();

      if (!poNum) continue;

      // Create PO entry if first time seeing this PO number
      if (!poMap[poNum]) {
        const rawDate = r[OFF.poDate];

        poMap[poNum] = {
          poNumber: poNum,
          poDate: _toDisplayDate(rawDate),
          poDateRaw: _toSafeDateString(rawDate) || '',
          vendor: String(r[OFF.vendor] || ''),
          contact: String(r[OFF.contact] || ''),
          poDescription: String(r[OFF.desc] || ''),
          poRemarks: String(r[OFF.remarks] || ''),
          supplierRemarks: String(r[OFF.supRemarks] || ''),
          items: [],
          grandTotal: 0,
          totalQty: 0
        };
      }

      // Add item to PO
      const qty = _toFiniteNumber(r[OFF.qty], 0);
      const price = _toFiniteNumber(r[OFF.price], 0);
      const total = qty * price;

      poMap[poNum].items.push({
        name: String(r[OFF.itemName] || ''),
        narration: String(r[OFF.narration] || ''),
        size: String(r[OFF.size] || ''),
        qty: qty,
        unit: String(r[OFF.unit] || 'Pcs'),
        price: price,
        baseQty: _toFiniteNumber(r[OFF.baseQty], qty),
        baseRate: _toFiniteNumber(r[OFF.baseRate], price)
      });

      // Update totals
      poMap[poNum].grandTotal += total;
      poMap[poNum].totalQty += qty;
    }

    // Attach a derived status (PO Issued / Partially Received / Completed)
    // to every PO based on what's actually been billed so far. Fails open —
    // if the Bill ledger can't be read for any reason, every PO simply shows
    // as 'PO Issued' rather than blocking the PO Ledger from loading.
    let billedMap = preloadedBilledMap;
    if (!billedMap) {
      billedMap = {};
      try {
        billedMap = _aggregateBilledBaseQtyByPo();
      } catch (e) {
        console.warn('[getPOData] Could not aggregate billed qty for status:', e.message);
      }
    }
    Object.values(poMap).forEach(function(po) {
      _attachPoStatus(po, billedMap);
    });

    // Convert to array and sort by PO number (descending)
    const sorted = Object.values(poMap).sort(function(a, b) {
      const numA = parseInt(String(a.poNumber).replace(/\D/g, ''), 10) || 0;
      const numB = parseInt(String(b.poNumber).replace(/\D/g, ''), 10) || 0;
      return numB - numA;
    });

    return buildResponse(true, sorted);
  } catch (error) {
    console.error('[getPOData] Error:', error.message);
    logAction('ERROR', 'getPOData', 'module_po', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to load PO data: ' + error.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────
// PO AUTO-MATCHING (for bills entered without a known PO)
// ─────────────────────────────────────────────────────────────────────────

/**
 * Max difference (in base-unit currency) allowed between a bill item's price
 * and a PO line's price for them to be considered the same deal.
 */
const PO_PRICE_MATCH_EPSILON = 0.01;

/**
 * suggestPoAllocations(vendor, items)
 *
 * Given a vendor and a set of "DIRECT" bill line items (qty/unit/price,
 * not yet linked to any PO), suggests which open PO line(s) each item most
 * likely belongs to. Matches by item name + size + price (compared in the
 * item's base unit, so a PO ordered in "Gross" can match a bill received in
 * "Pcs"), allocating oldest-PO-first, and splits a single bill line across
 * multiple PO lines if one PO doesn't have enough open quantity left.
 *
 * This never writes anything — it only proposes allocations for the UI to
 * show and the user to confirm or override before save.
 *
 * @param {string} vendor
 * @param {Array} items - [{ rowIndex, name, size, narration, qty, unit, price }]
 * @param {string} [billDate] - The bill's own date (any format toSafeDateString
 *   understands). PO lines dated after this are excluded — a bill can't be
 *   fulfilling a PO that didn't exist yet, so allowing that match would just
 *   produce a confusing, backdated-looking allocation.
 * @returns {Object} API response; data = [{ rowIndex, allocations: [{poNumber, qty, rateConflict?}], unmatchedQty }]
 *   rateConflict, when present, is { poRate, poUnit, billRate, billUnit } —
 *   the bill's rate still wins (qty is allocated regardless), this is only
 *   a notice for the UI to surface and let the user override if they want
 *   the PO's quoted rate kept instead.
 */
function suggestPoAllocations(vendor, items, billDate) {
  try {
    const vendorName = String(vendor || '').trim();
    if (!vendorName || !Array.isArray(items) || items.length === 0) {
      return buildResponse(true, []);
    }

    // Computed once and shared with getPOData below (which would otherwise
    // aggregate the entire Bill Ledger a second time just to derive status)
    // instead of each doing its own full Bill-sheet read.
    const billedMap = _aggregateBilledBaseQtyByPo();
    const poResponse = getPOData(billedMap);
    if (!poResponse.success) {
      return buildResponse(true, []); // fail open — never block bill entry
    }

    let vendorPOs = (poResponse.data || []).filter(function(po) {
      return String(po.vendor || '').trim().toLowerCase() === vendorName.toLowerCase();
    });
    if (vendorPOs.length === 0) {
      return buildResponse(true, []);
    }

    const billDateObj = toSafeDateObject(billDate);
    if (billDateObj) {
      vendorPOs = vendorPOs.filter(function(po) {
        if (!po.poDateRaw) return true; // no PO date on record — don't exclude on a guess
        return new Date(po.poDateRaw).getTime() <= billDateObj.getTime();
      });
    }
    if (vendorPOs.length === 0) {
      return buildResponse(true, []);
    }

    // Flatten to open PO lines (remaining base qty > 0), oldest PO first.
    const candidates = [];
    vendorPOs.forEach(function(po) {
      (po.items || []).forEach(function(poItem) {
        const key = _buildPoLineKey(po.poNumber, poItem.name, poItem.size, poItem.narration);
        const billedBaseQty = billedMap[key] || 0;
        const remainingBaseQty = poItem.baseQty - billedBaseQty;
        if (remainingBaseQty > 0.0001) {
          candidates.push({
            poNumber: po.poNumber,
            poDateRaw: po.poDateRaw,
            name: String(poItem.name),
            size: String(poItem.size),
            price: poItem.price,
            unit: String(poItem.unit || 'Pcs'),
            baseRate: poItem.baseRate,
            remainingBaseQty: remainingBaseQty
          });
        }
      });
    });

    candidates.sort(function(a, b) {
      const da = a.poDateRaw ? new Date(a.poDateRaw).getTime() : 0;
      const db = b.poDateRaw ? new Date(b.poDateRaw).getTime() : 0;
      if (da !== db) return da - db;
      return String(a.poNumber).localeCompare(String(b.poNumber), undefined, { numeric: true });
    });

    const itemUnitMap = _getItemUnitInfoMap();
    const unitsMap = _getUnitsMap();

    const results = items.map(function(item) {
      const rowIndex = item.rowIndex;
      const name = String(item.name || '').trim();
      const size = String(item.size || '').trim();
      const qty = Number(item.qty) || 0;
      const price = Number(item.price) || 0;
      const unit = String(item.unit || 'Pcs');

      if (!name || qty <= 0) {
        return { rowIndex: rowIndex, allocations: [], unmatchedQty: qty };
      }

      const unitInfo = _lookupItemUnitInfo(itemUnitMap, name, size);
      let baseQty = qty;
      let baseRate = price;
      try {
        baseQty = convertQtyToBaseUnit(qty, unit, unitInfo, unitsMap);
        baseRate = convertRateToBaseUnit(price, unit, unitInfo, unitsMap);
      } catch (e) {
        // Unconvertible unit (e.g. no Weight-per-Base-Unit set) — fall back
        // to comparing the as-entered qty/price directly, same as saveBill().
        baseQty = qty;
        baseRate = price;
      }

      // qty-per-1-baseQty, used to convert allocated base qty back to the
      // unit this bill line was actually entered in.
      const ratio = baseQty > 0 ? qty / baseQty : 1;

      const sameNameSize = candidates.filter(function(c) {
        return c.name.toLowerCase() === name.toLowerCase() &&
          c.size.toLowerCase() === size.toLowerCase() &&
          c.remainingBaseQty > 0.0001;
      });

      // Price is a preference, not a hard requirement: if the bill gives a
      // price and it disambiguates between PO lines at different rates
      // (e.g. the same item ordered at ₹23 on one PO and ₹25 on another),
      // prefer the matching-price lines. If price is absent, or doesn't
      // match any open line (a renegotiated rate), fall back to every open
      // line for that name+size, oldest PO first — never give up to DIRECT
      // just because the price didn't line up exactly.
      const priceMatched = price > 0
        ? sameNameSize.filter(function(c) {
            return Math.abs(c.baseRate - baseRate) <= PO_PRICE_MATCH_EPSILON;
          })
        : [];
      const matching = priceMatched.length > 0 ? priceMatched : sameNameSize;

      let baseQtyLeft = baseQty;
      const allocations = [];
      for (let i = 0; i < matching.length && baseQtyLeft > 0.0001; i++) {
        const c = matching[i];
        const take = Math.min(c.remainingBaseQty, baseQtyLeft);
        if (take <= 0) continue;

        // Bill rate has dominion over PO rate (the PO price is often just a
        // proposed/quoted rate, not what was actually paid) — but flag the
        // disagreement so the user can see it and explicitly choose to keep
        // the PO's rate instead, rather than this silently overwriting it.
        const allocation = { poNumber: c.poNumber, qty: parseFloat((take * ratio).toFixed(4)) };
        if (price > 0 && Math.abs(c.baseRate - baseRate) > PO_PRICE_MATCH_EPSILON) {
          allocation.rateConflict = { poRate: c.price, poUnit: c.unit, billRate: price, billUnit: unit };
        }
        allocations.push(allocation);

        // Mutate the shared candidate so other bill rows in this same batch
        // don't also get allocated against capacity this row just claimed.
        c.remainingBaseQty -= take;
        baseQtyLeft -= take;
      }

      return {
        rowIndex: rowIndex,
        allocations: allocations,
        unmatchedQty: parseFloat((baseQtyLeft * ratio).toFixed(4))
      };
    });

    return buildResponse(true, results);
  } catch (error) {
    console.error('[suggestPoAllocations] Error:', error.message);
    return buildResponse(true, []); // fail open — never block bill entry on a suggestion error
  }
}

// ─────────────────────────────────────────────────────────────────────────
// savePO
// ─────────────────────────────────────────────────────────────────────────

/**
 * savePO(formData)
 * 
 * Creates new PO or updates existing PO
 * 
 * Operations:
 * - Validate and normalize all input data
 * - Acquire document lock (concurrency safety)
 * - Generate PO number (if new) or validate (if edit)
 * - Parse and validate date
 * - Normalize items array
 * - Build sheet rows
 * - For edits: remove old rows, insert updated rows
 * - For new: append rows
 * - Flush changes to spreadsheet
 * 
 * Date Handling:
 * - If poDate provided: parse and validate
 * - If edit and no poDate: preserve existing date
 * - If new and no poDate: use today's date
 * 
 * @param {Object} formData - Form submission data
 *   @param {string} formData.poNumber - PO number (auto-generated if new)
 *   @param {string} formData.existingPoNumber - For edits: original PO number
 *   @param {string} formData.poDate - Date (DD/MM/YYYY or YYYY-MM-DD)
 *   @param {string} formData.vendor - Vendor name
 *   @param {string} formData.contact - Contact location
 *   @param {string} formData.poDescription - PO description
 *   @param {string} formData.poRemarks - PO remarks
 *   @param {string} formData.supplierRemarks - Supplier reference
 *   @param {Array|string} formData.items - Items array or JSON string
 * 
 * @returns {Object} API response
 *   - success: true/false
 *   - data: { poNumber } if successful
 *   - message: Status/error message
 */
function savePO(formData) {
  const lock = _getDocumentMutationLock();

  // Try to acquire lock with timeout
  if (!lock.tryLock(PO_LOCK_TIMEOUT_MS)) {
    return buildResponse(
      false,
      null,
      'System is busy. Please try again in a moment.'
    );
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.PO);
    if (!sheet) {
      throw new Error(`PO sheet "${APP_CONFIG.SHEETS.PO}" not found.`);
    }

    const startRow = APP_CONFIG.PO_SETTINGS.DATA_START_ROW;

    // Normalize and validate items
    const items = _normalizeItems(formData.items);

    // Determine if this is an edit or create
    const oldPoNum = String(formData.existingPoNumber || '').trim();
    const isEdit = !!oldPoNum;

    let poNumToSave = String(formData.poNumber || oldPoNum || '').trim();

    // Validate or generate PO number
    if (isEdit) {
      poNumToSave = _validatePoNumber(poNumToSave);
    } else {
      poNumToSave = generateNextId(
        sheet,
        startRow,
        APP_CONFIG.PO_SETTINGS.ID_COL_INDEX,
        PO_ID_POLICY
      );
    }

    // Parse and validate date
    let poDateString = null;

    if (formData.poDate && String(formData.poDate).trim() !== '') {
      poDateString = _toSafeDateString(formData.poDate);

      if (!poDateString) {
        return buildResponse(
          false,
          null,
          `Could not understand the date format: "${formData.poDate}". Please use DD/MM/YYYY or YYYY-MM-DD.`
        );
      }
    } else if (isEdit) {
      // Edit without new date: preserve existing date
      console.warn(
        `[savePO] Edit for PO #${oldPoNum} received empty poDate. Preserving existing date.`
      );

      const existingSheetDate = _findExistingPoDate(sheet, startRow, oldPoNum);
      poDateString = _toSafeDateString(existingSheetDate) || _toSafeDateString(new Date());
    } else {
      // New PO without date: use today
      poDateString = _toSafeDateString(new Date());
    }

    // Build PO header
    const header = _makePoHeader(formData, poNumToSave, poDateString, items);

    // Build sheet rows
    const newRows = items.map(function(item) {
      return _buildPoSheetRow(header, item);
    });

    // Execute create or update
    if (isEdit) {
      const firstMatchRow = _findFirstPoRow(sheet, startRow, oldPoNum);

      if (firstMatchRow === -1) {
        throw new Error(
          `Original PO #${oldPoNum} not found in sheet. Edit aborted to prevent data corruption.`
        );
      }

      // Remove old rows and insert updated rows at same position
      _rewriteWithoutMatchingRows(sheet, startRow, PO_COL.PO_NUMBER, oldPoNum);
      _insertRowsAt(sheet, firstMatchRow, newRows);
    } else {
      // Append new rows
      _appendRows(sheet, newRows);
    }

    // Auto-extract vendors, items, and rates
    autoExtractFromPoOrBill(formData.vendor, formData.contact, items, {
      date: header.poDate,
      poNumber: poNumToSave
    });

    // Flush changes
    SpreadsheetApp.flush();

    const msg = isEdit
      ? `PO updated successfully to #${poNumToSave}.`
      : `PO #${poNumToSave} created successfully.`;

    logAction('CREATE', APP_CONFIG.SHEETS.PO, poNumToSave, `Items: ${items.length}`, 'SUCCESS');

    return buildResponse(true, { poNumber: poNumToSave }, msg);
  } catch (error) {
    console.error('[savePO] Error:', error.message);
    logAction('ERROR', 'savePO', formData.poNumber, error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to save PO: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────────────────
// deletePO
// ─────────────────────────────────────────────────────────────────────────

/**
 * deletePO(poNumber)
 * 
 * Deletes PO and all associated bills (cascade delete)
 * 
 * Operations:
 * 1. Acquire document lock
 * 2. Validate PO number
 * 3. Remove all rows from PO sheet matching PO number
 * 4. Remove all rows from Bill sheet matching PO number
 * 5. Flush changes
 * 
 * Safety:
 * - Document lock prevents concurrent deletes
 * - Validates PO number before deletion
 * - Atomicity: both sheets updated before return
 * - Audit log entry for each delete
 * 
 * @param {string|number} poNumber - PO number to delete
 * 
 * @returns {Object} API response
 *   - success: true/false
 *   - data: null
 *   - message: Status/error message
 */
function deletePO(poNumber) {
  const lock = _getDocumentMutationLock();

  if (!lock.tryLock(PO_LOCK_TIMEOUT_MS)) {
    return buildResponse(
      false,
      null,
      'System is busy. Please try again.'
    );
  }

  try {
    // Validate PO number
    const validPoNumber = _validatePoNumber(poNumber);

    // Delete from PO sheet only. Bills are an independent financial record of
    // what was actually received — a bill row that references this PO number
    // keeps that reference even after the PO itself is gone; it must never be
    // deleted or altered just because its PO was deleted.
    const poSheet = getSheet(APP_CONFIG.SHEETS.PO);
    if (!poSheet) {
      throw new Error('PO sheet not found.');
    }

    _rewriteWithoutMatchingRows(
      poSheet,
      APP_CONFIG.PO_SETTINGS.DATA_START_ROW,
      PO_COL.PO_NUMBER,
      validPoNumber
    );

    // Flush changes
    SpreadsheetApp.flush();

    const msg = `PO #${validPoNumber} deleted.`;
    logAction('DELETE', APP_CONFIG.SHEETS.PO, validPoNumber, msg, 'SUCCESS');

    return buildResponse(true, null, msg);
  } catch (error) {
    console.error('[deletePO] Error:', error.message);

    const errorMsg = 'Failed to delete PO: ' + error.message;
    logAction('ERROR', 'deletePO', poNumber, errorMsg, 'ERROR');
    return buildResponse(false, null, errorMsg);
  } finally {
    lock.releaseLock();
  }
}

/**
 * deletePOsBulk(poNumbers)
 *
 * Deletes multiple POs from the PO sheet only. Bills referencing these PO
 * numbers are left untouched — they remain a standalone record of what was
 * actually received regardless of whether the PO still exists.
 *
 * @param {Array<string|number>} poNumbers - PO numbers to delete
 * @returns {Object} API response
 */
function deletePOsBulk(poNumbers) {
  const lock = _getDocumentMutationLock();

  if (!lock.tryLock(PO_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const targetSet = new Set(
      (poNumbers || []).map(p => String(p).trim()).filter(Boolean)
    );

    if (targetSet.size === 0) {
      return buildResponse(true, null, 'No purchase orders selected.');
    }

    const poSheet = getSheet(APP_CONFIG.SHEETS.PO);
    if (!poSheet) {
      throw new Error('PO sheet not found.');
    }

    // PO deletion only touches the PO sheet. Bills are an independent record
    // of what was actually received and must never be removed or altered
    // just because the PO they reference was deleted.
    const { rowsDeleted: poRowsDeleted, deletedIds } = _rewriteWithoutMatchingRowsBulk(
      poSheet,
      APP_CONFIG.PO_SETTINGS.DATA_START_ROW,
      PO_COL.PO_NUMBER,
      targetSet
    );

    SpreadsheetApp.flush();

    const msg = `Deleted ${targetSet.size} PO(s): ${poRowsDeleted} PO row(s) removed.`;
    logAction('BULK_DELETE', APP_CONFIG.SHEETS.PO, 'multiple', msg, 'SUCCESS');

    return buildResponse(true, { deletedIds }, msg);
  } catch (error) {
    console.error('[deletePOsBulk] Error:', error.message);
    logAction('ERROR', 'deletePOsBulk', 'multiple', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete POs: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Rewrites PO Tracker rows referencing a renamed/merged item, so historical
 * POs stay attributed to the item's current name + size.
 *
 * Narration sits between ITEM_NAME and SIZE in this sheet, so the 3-column
 * block is read/written together and only the name/size cells are mutated —
 * this leaves narration untouched.
 *
 * Runs under the caller's existing document lock — does not acquire its own.
 *
 * @param {string} oldName
 * @param {string} oldSize
 * @param {string} newName
 * @param {string} newSize
 */
function backfillPOItemRefs(oldName, oldSize, newName, newSize) {
  const sheet = getSheet(APP_CONFIG.SHEETS.PO);
  if (!sheet) return;

  const startRow = APP_CONFIG.PO_SETTINGS.DATA_START_ROW;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  const numRows = lastRow - startRow + 1;
  const range = sheet.getRange(startRow, PO_COL.ITEM_NAME, numRows, 3);
  const values = range.getValues();

  const tOldName = String(oldName || '').trim().toLowerCase();
  const tOldSize = String(oldSize || '').trim().toLowerCase();

  let changed = false;
  for (let i = 0; i < values.length; i++) {
    const rowName = String(values[i][0] || '').trim().toLowerCase();
    const rowSize = String(values[i][2] || '').trim().toLowerCase();
    if (rowName === tOldName && rowSize === tOldSize) {
      values[i][0] = newName;
      values[i][2] = newSize;
      changed = true;
    }
  }

  if (changed) {
    range.setValues(values);
  }
}
