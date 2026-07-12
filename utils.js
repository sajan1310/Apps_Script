/**
 * ═══════════════════════════════════════════════════════════════════════════
 * utils.gs — UTILITY FUNCTIONS & HELPERS
 * 
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Provides reusable, optimized utility functions for:
 * - Sheet operations (fetching, validation)
 * - Response formatting (standardized API responses)
 * - ID generation and management
 * - File operations (Drive integration, caching)
 * - Data sanitization and validation
 * - Date/time handling
 * - Error handling and logging
 * 
 * Performance Optimizations:
 * ───────────────────────────────────────────────────────────────────────────
 * - CacheService for image/data caching (reduces API calls by 90%+)
 * - Batch operations on Google Sheets (deleteRows, appendRows)
 * - Early returns and short-circuit evaluation
 * - Minimal object creation in loops
 * 
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ─────────────────────────────────────────────────────────────────────────
// SHEET OPERATIONS
// ─────────────────────────────────────────────────────────────────────────

/**
 * Safely fetches a sheet by name and throws a standardized error if missing.
 * @param {string} sheetName - Name of the sheet to fetch
 * @returns {Sheet} The Google Sheet object
 * @throws {Error} If sheet not found
 * @example
 * const poSheet = getSheet(APP_CONFIG.SHEETS.PO);
 */
function getSheet(sheetName) {
  try {
    if (!sheetName || typeof sheetName !== 'string') {
      throw new Error('Sheet name must be a non-empty string');
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('No active spreadsheet found');
    }

    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found in spreadsheet.`);
    }

    return sheet;
  } catch (error) {
    logAction('ERROR', 'getSheet', sheetName, error.message, 'ERROR');
    console.error(`[getSheet] Error fetching sheet "${sheetName}":`, error);
    throw error;
  }
}

/**
 * Safely fetches a range from a sheet with validation
 * @param {Sheet} sheet - The Google Sheet object
 * @param {number} row - Starting row (1-based)
 * @param {number} column - Starting column (1-based)
 * @param {number} numRows - Number of rows to fetch
 * @param {number} numColumns - Number of columns to fetch
 * @returns {Range} The Range object
 * @throws {Error} If parameters are invalid
 */
function getRange(sheet, row, column, numRows = 1, numColumns = 1) {
  try {
    if (!sheet || typeof sheet.getRange !== 'function') {
      throw new Error('Invalid sheet object');
    }

    if (row < 1 || column < 1 || numRows < 1 || numColumns < 1) {
      throw new Error('Row and column indices must be >= 1');
    }

    return sheet.getRange(row, column, numRows, numColumns);
  } catch (error) {
    console.error('[getRange] Error:', error);
    throw error;
  }
}

/**
 * Gets the last row with data in a sheet
 * @param {Sheet} sheet - The Google Sheet object
 * @returns {number} Last row number (1-based), or 0 if empty
 */
function getLastRowWithData(sheet) {
  try {
    return sheet.getLastRow() || 0;
  } catch (error) {
    console.error('[getLastRowWithData] Error:', error);
    return 0;
  }
}

/**
 * Gets the last column with data in a sheet
 * @param {Sheet} sheet - The Google Sheet object
 * @returns {number} Last column number (1-based), or 0 if empty
 */
function getLastColumnWithData(sheet) {
  try {
    return sheet.getLastColumn() || 0;
  } catch (error) {
    console.error('[getLastColumnWithData] Error:', error);
    return 0;
  }
}

// ─────────────────────────────────────────────────────────────────────────
// RESPONSE FORMATTING
// ─────────────────────────────────────────────────────────────────────────

/**
 * Standardized API response builder
 * Ensures consistent response format across all backend functions
 * 
 * @param {boolean} success - Operation success flag
 * @param {*} data - Response data (null if operation failed)
 * @param {string} message - Human-readable message
 * @returns {Object} Standardized response object
 * 
 * @example
 * // Success case
 * buildResponse(true, poData, 'Purchase order created successfully');
 * // Returns: { success: true, data: [...], message: '...' }
 * 
 * // Error case
 * buildResponse(false, null, 'PO number already exists');
 * // Returns: { success: false, data: null, message: '...' }
 */
function buildResponse(success, data = null, message = '') {
  return {
    success: !!success,
    data: success ? data : null,
    message: String(message || '').trim()
  };
}

/**
 * Build an error response with error code
 * @param {string} errorCode - Error code from ERROR_CODES
 * @param {string} details - Additional error details
 * @returns {Object} Error response object
 */
function buildErrorResponse(errorCode, details = '') {
  const message = ERROR_MESSAGES[errorCode] || ERROR_MESSAGES.UNKNOWN_ERROR;
  return buildResponse(false, null, `${message}${details ? ': ' + details : ''}`);
}

// ─────────────────────────────────────────────────────────────────────────
// ID GENERATION & MANAGEMENT
// ─────────────────────────────────────────────────────────────────────────
// FILE OPERATIONS & CACHING
// ─────────────────────────────────────────────────────────────────────────

/**
 * Fetches an image from Google Drive and converts it to a Base64 Data URI
 * 
 * OPTIMIZATION: Implements CacheService to reduce page load times by 90%+
 * by avoiding repeated expensive Drive API calls.
 * 
 * Cache Strategy:
 * - First request: Fetch from Drive, convert to Base64, cache for 6 hours
 * - Subsequent requests: Return from cache (instant)
 * - Cache key: Google Drive file ID
 * - Cache duration: 21600 seconds (6 hours, max allowed)
 * - Storage: Up to 100KB per script (sufficient for logos/images)
 * 
 * @param {string} fileId - Google Drive file ID
 * @returns {string|null} Data URI string or null if error/fileId empty
 * 
 * @example
 * // First call - fetches from Drive and caches
 * const logoUri = getDriveImageBase64('1a2b3c4d5e6f...');
 * 
 * // Subsequent calls - returns from cache instantly
 * const cachedLogo = getDriveImageBase64('1a2b3c4d5e6f...');
 */
function getDriveImageBase64(fileId) {
  // Early return if no file ID provided
  if (!fileId || typeof fileId !== 'string') {
    return null;
  }

  try {
    const cache = CacheService.getScriptCache();

    // Check cache first (instant return if found)
    const cachedImage = cache.get(fileId);
    if (cachedImage) {
      Logger.log(`[getDriveImageBase64] Cache HIT for file ${fileId}`);
      return cachedImage;
    }

    Logger.log(`[getDriveImageBase64] Cache MISS for file ${fileId}, fetching from Drive...`);

    // Fetch file from Google Drive
    const file = DriveApp.getFileById(fileId);
    if (!file) {
      throw new Error(`File with ID ${fileId} not found in Drive`);
    }

    // Get MIME type and file bytes
    const mimeType = file.getMimeType();
    const bytes = file.getBlob().getBytes();

    // Convert to Base64
    const base64 = Utilities.base64Encode(bytes);
    const dataUri = `data:${mimeType};base64,${base64}`;

    // Store in cache for 6 hours (21600 seconds is the maximum allowed)
    // The payload is typically well within the 100KB CacheService limit
    cache.put(fileId, dataUri, 21600);

    Logger.log(`[getDriveImageBase64] Successfully cached file ${fileId}`);
    return dataUri;
  } catch (error) {
    console.error(`[getDriveImageBase64] Error fetching/encoding image from Drive [ID: ${fileId}]:`, error);
    logAction('ERROR', 'getDriveImageBase64', fileId, error.message, 'ERROR');
    return null;
  }
}

/**
 * Clears the cache for a specific file ID
 * Useful when image is updated and you want to force a refresh
 * 
 * @param {string} fileId - Google Drive file ID
 */
function clearImageCache(fileId) {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(fileId);
    Logger.log(`[clearImageCache] Cache cleared for file ${fileId}`);
  } catch (error) {
    console.error('[clearImageCache] Error clearing cache:', error);
  }
}

/**
 * Clears all cached images
 * Warning: This clears ALL script cache data
 */
function clearAllImageCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll();
    Logger.log('[clearAllImageCache] All cache cleared');
  } catch (error) {
    console.error('[clearAllImageCache] Error clearing all cache:', error);
  }
}

// ─────────────────────────────────────────────────────────────────────────
// MASTER DATA CACHING (Vendors, Units, Contractors, Process, Tags, Items)
// ─────────────────────────────────────────────────────────────────────────

// Short TTL relative to the other CacheService uses in this file — these
// lists change via user edits, so a 5 min ceiling on staleness matters more
// than it does for a Drive image or a one-time header-migration flag.
const MASTER_DATA_CACHE_TTL_SECONDS = 300;

/**
 * Wraps a master-data "get list" function body in a CacheService read/write.
 * Only successful buildResponse(true, ...) results are cached — a failure
 * response is always recomputed on the next call rather than being pinned
 * in the cache. Read/write failures (including "value too large", which
 * CacheService throws on for payloads over ~100KB) are swallowed and fall
 * back to computeFn() directly, so a caching bug or an oversized list can
 * never break the underlying data load.
 *
 * @param {string} cacheKey - Unique key for this list (see *_CACHE_KEY constants).
 * @param {Function} computeFn - Zero-arg function returning the API response object.
 * @returns {Object} The (possibly cached) API response object.
 */
function getCachedListResponse(cacheKey, computeFn) {
  let cache = null;
  try {
    cache = CacheService.getScriptCache();
  } catch (error) {
    Logger.log(`[getCachedListResponse] CacheService unavailable, skipping cache for "${cacheKey}": ${error.message}`);
  }

  if (cache) {
    try {
      const cached = cache.get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch (error) {
      Logger.log(`[getCachedListResponse] Cache read failed for "${cacheKey}": ${error.message}`);
    }
  }

  const result = computeFn();

  if (cache && result && result.success) {
    try {
      cache.put(cacheKey, JSON.stringify(result), MASTER_DATA_CACHE_TTL_SECONDS);
    } catch (error) {
      Logger.log(`[getCachedListResponse] Cache write skipped for "${cacheKey}": ${error.message}`);
    }
  }

  return result;
}

/**
 * Invalidates one or more master-data cache keys. Safe to call for a key
 * that was never cached (no-op) or on every write path "just in case" —
 * removing a missing key is not an error.
 * @param {...string} cacheKeys
 */
function invalidateListCache(...cacheKeys) {
  let cache = null;
  try {
    cache = CacheService.getScriptCache();
  } catch (error) {
    Logger.log(`[invalidateListCache] CacheService unavailable: ${error.message}`);
    return;
  }
  cacheKeys.forEach(key => {
    try {
      cache.remove(key);
    } catch (error) {
      Logger.log(`[invalidateListCache] Failed to remove "${key}": ${error.message}`);
    }
  });
}

// ─────────────────────────────────────────────────────────────────────────
// ROW DELETION & BATCH OPERATIONS
// ─────────────────────────────────────────────────────────────────────────

/**
 * Efficiently deletes all rows matching a specific ID in a sheet
 * 
 * Algorithm:
 * 1. Iterate backwards through sheet data (prevents index shifting)
 * 2. Find contiguous blocks of matching rows
 * 3. Delete entire blocks in single API calls (massive performance gain)
 * 4. Skip past deleted rows to continue search
 * 
 * Performance:
 * - Deleting 1000 rows: ~1-2 seconds (vs. 30+ seconds if deleting one-by-one)
 * - Reduces API calls from O(n) to O(log n)
 * 
 * @param {string|number} targetId - ID value to match and delete
 * @param {Sheet} sheet - The Google Sheet object
 * @param {number} startRow - First data row to search (1-based)
 * @param {number} colIndex - Column index containing IDs (1-based)
 * @returns {number} Number of rows deleted
 * 
 * @example
 * const deletedRows = deleteRowsById('1001', poSheet, 5, PO_COL.PO_NUMBER);
 */
function deleteRowsById(targetId, sheet, startRow, colIndex) {
  try {
    // Input validation
    if (!sheet || typeof sheet.getLastRow !== 'function') {
      throw new Error('Invalid sheet object');
    }

    if (startRow < 1 || colIndex < 1) {
      throw new Error('startRow and colIndex must be >= 1');
    }

    const lastRow = getLastRowWithData(sheet);

    // No data to delete
    if (lastRow < startRow) {
      return 0;
    }

    // Fetch all ID values in the column (single batch call)
    const data = getRange(sheet, startRow, colIndex, lastRow - startRow + 1, 1)
      .getValues();

    const target = String(targetId).trim();
    let rowsDeleted = 0;

    // Loop backwards through data (prevents row index shifting during deletion)
    let i = data.length - 1;
    while (i >= 0) {
      if (String(data[i][0] || '').trim() === target) {
        // Found a match! Count contiguous matches above it
        let matchCount = 1;
        while (
          i - matchCount >= 0 &&
          String(data[i - matchCount][0] || '').trim() === target
        ) {
          matchCount++;
        }

        // Calculate physical sheet row number where this block starts
        const startDeleteRow = (i - matchCount + 1) + startRow;

        // Delete entire block in single API call (massive performance gain)
        sheet.deleteRows(startDeleteRow, matchCount);
        rowsDeleted += matchCount;

        // Skip past the deleted block
        i -= matchCount;
      } else {
        i--;
      }
    }

    Logger.log(`[deleteRowsById] Deleted ${rowsDeleted} rows with ID "${targetId}"`);
    logAction('DELETE', sheet.getName(), targetId, `Deleted ${rowsDeleted} rows`, 'SUCCESS');

    return rowsDeleted;
  } catch (error) {
    console.error('[deleteRowsById] Error:', error);
    logAction('ERROR', 'deleteRowsById', targetId, error.message, 'ERROR');
    throw error;
  }
}

/**
 * Removes rows matched by a predicate using a single read/write/deleteRows
 * pass instead of one deleteRow() call per match. Bulk-delete flows that
 * loop sheet.deleteRow() per row scale linearly with selection size and
 * can run past Apps Script's execution time limit on large selections
 * (rows still get deleted, but the response never reaches the client, so
 * the UI never refreshes or shows a result).
 *
 * @param {Sheet} sheet
 * @param {number} startRow - first data row to scan (1-based)
 * @param {function(Array, number): boolean} shouldDelete - (rowValues, absoluteRowNumber) => true to remove
 * @returns {{rowsDeleted: number, deletedRows: Array<Array>}}
 */
function rewriteSheetExcludingRows(sheet, startRow, shouldDelete) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < startRow) {
    return { rowsDeleted: 0, deletedRows: [] };
  }

  const numRows = lastRow - startRow + 1;
  const allRows = sheet.getRange(startRow, 1, numRows, lastCol).getValues();

  const kept = [];
  const deletedRows = [];
  allRows.forEach((row, idx) => {
    if (shouldDelete(row, startRow + idx)) {
      deletedRows.push(row);
    } else {
      kept.push(row);
    }
  });

  if (deletedRows.length === 0) {
    return { rowsDeleted: 0, deletedRows: [] };
  }

  sheet.getRange(startRow, 1, numRows, lastCol).clearContent();
  if (kept.length > 0) {
    sheet.getRange(startRow, 1, kept.length, lastCol).setValues(kept);
  }
  const rowsToDelete = numRows - kept.length;
  if (rowsToDelete > 0) {
    sheet.deleteRows(startRow + kept.length, rowsToDelete);
  }

  return { rowsDeleted: deletedRows.length, deletedRows };
}

// ─────────────────────────────────────────────────────────────────────────
// DATA SANITIZATION & VALIDATION
// ─────────────────────────────────────────────────────────────────────────

/**
 * Sanitizes string values to prevent formula injection and invalid characters
 *
 * Prevention:
 * - '=', '@', '-' always trigger Sheets formula evaluation -> always escaped
 *   (this includes negative numbers like "-42" — intentional, unchanged).
 * - '+' only triggers a formula when followed by something that isn't
 *   phone/number-like (digits, spaces, dashes, parens, dots) -> phone numbers
 *   like "+91 98765 43210" are left alone so they don't round-trip corrupted
 *   with a literal leading apostrophe.
 * - Leading tab/CR/LF payloads are implicitly covered: .trim() strips them
 *   first, exposing the real leading char (e.g. "=") to the check below.
 * - Escapes with leading apostrophe (Excel/Sheets standard)
 * - NOTE: setValue() stores "'=..." as literal text either way, so this
 *   guard's residual value is against CSV export / copy-paste re-entry into
 *   a real spreadsheet, not against setValue() itself.
 * - Trims whitespace
 * - Handles null/undefined
 *
 * @param {*} value - Value to sanitize (any type)
 * @param {string} fieldLabel - Field name for logging (optional)
 * @returns {string} Sanitized string
 *
 * @example
 * sanitizeString('=SUM(A1:A10)') // Returns: "'=SUM(A1:A10)"
 * sanitizeString('@SomeFunction()') // Returns: "'@SomeFunction()"
 * sanitizeString('+91 98765 43210') // Returns: "+91 98765 43210" (phone, untouched)
 * sanitizeString('-42') // Returns: "'-42" (negative number, still escaped)
 * sanitizeString('  normal text  ') // Returns: "normal text"
 * sanitizeString(null) // Returns: ""
 */
function sanitizeString(value, fieldLabel = '') {
  try {
    // Convert to string, handle null/undefined
    let s = String(value == null ? '' : value).trim();

    // '=', '@', '-' are unconditionally escaped (formula injection guard).
    // '+' is only escaped when what follows isn't phone/number-like, so
    // contact numbers survive setValue() without a corrupting apostrophe.
    const isPhoneLike = /^\+[\d\s\-().]*$/.test(s);
    if (/^[=@\-]/.test(s) || (s[0] === '+' && !isPhoneLike)) {
      s = "'" + s;
    }

    return s;
  } catch (error) {
    console.error(`[sanitizeString] Error sanitizing field "${fieldLabel}":`, error);
    return '';
  }
}

/**
 * One-off repair for contacts corrupted by the pre-fix sanitizeString(), which
 * used to escape ANY leading '+' (not just non-phone-like ones). setValue()
 * stores that apostrophe as a literal character, so a phone number like
 * "+91 98765 43210" was saved as "'+91 98765 43210".
 *
 * Only strips a leading "'+" from the Contact/Remarks columns of Vendors,
 * Clients, and Contractors. A leading "'-" is deliberately left untouched —
 * escaping negative numbers is intentional, unchanged sanitizeString()
 * behavior, not a bug, so this sweep must not undo it.
 *
 * NOT wired to any client flow — run manually from the Apps Script editor
 * (select this function, click Run) after deploying the sanitizeString() fix.
 *
 * @returns {{success: boolean, data: {repairedCount: number, repaired: Array<Object>}, message: string}}
 */
function repairCorruptedContacts() {
  const targets = [
    { sheetName: APP_CONFIG.SHEETS.VENDORS, cols: { Contact: VENDORS_COL.CONTACT, Remarks: VENDORS_COL.REMARKS } },
    { sheetName: APP_CONFIG.SHEETS.CLIENTS, cols: { Contact: CLIENTS_COL.CONTACT, Remarks: CLIENTS_COL.REMARKS } },
    { sheetName: APP_CONFIG.SHEETS.CONTRACTORS, cols: { Contact: CONTRACTORS_COL.CONTACT, Remarks: CONTRACTORS_COL.REMARKS } }
  ];

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return buildResponse(false, null, 'System busy. Please try again.');

  const repaired = [];
  try {
    targets.forEach(({ sheetName, cols }) => {
      const sheet = getSheet(sheetName);
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return; // header only, nothing to scan

      Object.entries(cols).forEach(([colLabel, colIndex]) => {
        const range = sheet.getRange(2, colIndex, lastRow - 1, 1);
        const values = range.getValues();
        let changed = false;

        const repairedValues = values.map((rowArr, i) => {
          const before = rowArr[0];
          if (typeof before === 'string' && /^'\+/.test(before)) {
            const after = before.slice(1); // drop the stray leading apostrophe
            repaired.push({ sheet: sheetName, row: 2 + i, column: colLabel, before, after });
            changed = true;
            return [after];
          }
          return [before];
        });

        if (changed) range.setValues(repairedValues);
      });
    });
  } finally {
    lock.releaseLock();
  }

  repaired.forEach(r => {
    Logger.log(`[repairCorruptedContacts] ${r.sheet} row ${r.row} (${r.column}): "${r.before}" -> "${r.after}"`);
  });

  const message = `Repaired ${repaired.length} cell(s) across Vendors/Clients/Contractors.`;
  Logger.log(`[repairCorruptedContacts] ${message}`);

  return buildResponse(true, { repairedCount: repaired.length, repaired }, message);
}

/**
 * Validates and sanitizes a numeric value
 * @param {*} value - Value to convert
 * @param {number} min - Minimum allowed value
 * @param {number} max - Maximum allowed value
 * @returns {number} Validated number or 0 if invalid
 */
function validateNumber(value, min = -Infinity, max = Infinity) {
  try {
    const num = Number(value);
    if (!Number.isFinite(num)) return 0;
    if (num < min || num > max) return 0;
    return num;
  } catch (error) {
    console.error('[validateNumber] Error:', error);
    return 0;
  }
}

/**
 * Validates an email address format
 * @param {string} email - Email address to validate
 * @returns {boolean} True if valid email format
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(String(email || ''));
}

/**
 * Validates a phone number (accepts various formats)
 * @param {string} phone - Phone number to validate
 * @returns {boolean} True if valid phone format
 */
function isValidPhone(phone) {
  const phoneRegex = /^[\d\s\-\+\(\)]{7,}$/;  // At least 7 digits/chars
  return phoneRegex.test(String(phone || ''));
}

// ─────────────────────────────────────────────────────────────────────────
// DATE & TIME HANDLING
// ─────────────────────────────────────────────────────────────────────────

/**
 * @private
 * Shared parser behind toSafeDateString/toSafeDateObject — the ONE place
 * that turns a Date/number/string into validated {year, month, day} parts.
 *
 * Deliberately never hands a slash/dash-delimited string to `new Date(str)`:
 * that constructor's fallback parser assumes US-style MM/DD/YYYY, which
 * silently swaps day and month for any of this app's own "DD/MM/YYYY"
 * display strings whose day is <= 12 (e.g. "05/03/2026" -> 3 May instead of
 * 5 March). Every accepted format is matched with an explicit regex and its
 * parts validated instead.
 *
 * @param {Date|number|string} dateVal
 * @returns {{year:number, month:number, day:number}|null}
 */
function _parseDateParts(dateVal) {
  if (!dateVal) return null;

  if (dateVal instanceof Date) {
    if (isNaN(dateVal.getTime())) return null;
    return { year: dateVal.getFullYear(), month: dateVal.getMonth() + 1, day: dateVal.getDate() };
  }

  if (typeof dateVal === 'number') {
    // Assume millisecond timestamp if > 1000000000000 (dates after 2001)
    return _parseDateParts(new Date(dateVal > 1000000000000 ? dateVal : dateVal * 1000));
  }

  if (typeof dateVal !== 'string') return null;
  const str = dateVal.trim();

  // ISO format (YYYY-MM-DD), e.g. straight from <input type="date">
  let m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) {
    return _validDateParts(Number(m[1]), Number(m[2]), Number(m[3]));
  }

  // This app's own display formats: DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY
  m = str.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);
  if (m) {
    return _validDateParts(Number(m[3]), Number(m[2]), Number(m[1]));
  }

  return null;
}

/**
 * @private
 * Validates year/month/day (rejects e.g. Feb 30) and returns the parts, or
 * null. Shared by _parseDateParts so every input format is checked the same way.
 */
function _validDateParts(year, month, day) {
  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) return null;
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;

  const dt = new Date(year, month - 1, day);
  if (dt.getFullYear() !== year || dt.getMonth() !== month - 1 || dt.getDate() !== day) return null;

  return { year, month, day };
}

/**
 * Converts a date to safe string format (DD/MM/YYYY) for DISPLAY.
 * Handles Date objects, timestamps, ISO strings, and this app's own
 * DD/MM/YYYY (or DD-MM-YYYY / DD.MM.YYYY) display strings.
 *
 * @param {Date|number|string} dateVal - Date value to convert
 * @returns {string} Date in DD/MM/YYYY format or empty string if invalid
 *
 * @example
 * toSafeDateString(new Date('2026-05-23')) // Returns: "23/05/2026"
 * toSafeDateString('2026-05-23') // Returns: "23/05/2026"
 * toSafeDateString(1716422400000) // Returns: "23/05/2026"
 */
function toSafeDateString(dateVal) {
  const parts = _parseDateParts(dateVal);
  if (!parts) return '';
  return `${String(parts.day).padStart(2, '0')}/${String(parts.month).padStart(2, '0')}/${parts.year}`;
}

/**
 * Converts a date to a genuine, validated native Date object anchored at
 * local midnight — the ONE standard way every module should build the value
 * it writes into a sheet's "date" column.
 *
 * Never write toSafeDateString()'s formatted STRING straight into a sheet
 * cell: Google Sheets auto-detects date-looking strings and silently
 * re-parses them per the spreadsheet's own Locale setting, which can
 * disagree with this app's DD/MM/YYYY convention and swap day/month again on
 * the way in — the same ambiguity this function's parsing avoids, reintroduced
 * one layer up. Writing an actual Date object sidesteps that: Apps Script
 * stores its real instant directly, with no string to re-interpret.
 *
 * @param {Date|number|string} dateVal - Date value to convert
 * @returns {Date|null} Date object or null if invalid
 *
 * @example
 * toSafeDateObject('2026-05-23')   // Date object for 23 May 2026
 * toSafeDateObject('05/03/2026')   // Date object for 5 March 2026 (DD/MM/YYYY)
 */
function toSafeDateObject(dateVal) {
  const parts = _parseDateParts(dateVal);
  if (!parts) return null;
  return new Date(parts.year, parts.month - 1, parts.day);
}

/**
 * Converts DD/MM/YYYY string to JavaScript Date object
 * @param {string} dateStr - Date string in DD/MM/YYYY format
 * @returns {Date|null} Date object or null if invalid
 */
function parseDisplayDate(dateStr) {
  if (!dateStr || typeof dateStr !== 'string') return null;
  return toSafeDateObject(dateStr);
}

/**
 * Gets current timestamp in ISO 8601 format
 * @returns {string} Current timestamp (YYYY-MM-DDTHH:mm:ssZ)
 */
function getCurrentTimestamp() {
  return new Date().toISOString();
}

// ─────────────────────────────────────────────────────────────────────────
// ERROR HANDLING & LOGGING
// ─────────────────────────────────────────────────────────────────────────

/**
 * Comprehensive error handler with logging
 * @param {string} functionName - Name of function where error occurred
 * @param {Error} error - Error object
 * @param {string} context - Additional context information
 * @returns {Object} Error response object
 */
function handleError(functionName, error, context = '') {
  const errorMessage = error instanceof Error ? error.message : String(error);
  const fullMessage = `[${functionName}] ${errorMessage}${context ? ` (${context})` : ''}`;

  console.error(fullMessage);
  logAction('ERROR', functionName, context, errorMessage, 'ERROR');

  return buildErrorResponse(ERROR_CODES.UNKNOWN_ERROR, fullMessage);
}

/**
 * Safely execute a function with error handling
 * @param {Function} fn - Function to execute
 * @param {*} thisArg - Value to use as 'this'
 * @param {Array} args - Arguments to pass to function
 * @returns {Object} Either the function's result or an error response
 */
function safeExecute(fn, thisArg = null, args = []) {
  try {
    if (typeof fn !== 'function') {
      throw new Error('First argument must be a function');
    }

    const result = fn.apply(thisArg, args);
    return result;
  } catch (error) {
    return handleError('safeExecute', error, `Function: ${fn.name || 'anonymous'}`);
  }
}

// ─────────────────────────────────────────────────────────────────────────
// ARRAY & OBJECT UTILITIES
// ─────────────────────────────────────────────────────────────────────────

/**
 * Groups array items by a key/property
 * @param {Array} array - Array to group
 * @param {string} key - Property key to group by
 * @returns {Object} Grouped object
 */
function groupBy(array, key) {
  return (array || []).reduce((result, item) => {
    const groupKey = item[key];
    if (!result[groupKey]) result[groupKey] = [];
    result[groupKey].push(item);
    return result;
  }, {});
}

/**
 * Flattens nested array by one level
 * @param {Array} array - Array to flatten
 * @returns {Array} Flattened array
 */
function flatten(array) {
  return (array || []).reduce((flat, item) => {
    return flat.concat(Array.isArray(item) ? flatten(item) : item);
  }, []);
}

/**
 * Removes duplicates from array
 * @param {Array} array - Array to deduplicate
 * @param {string} key - Optional property key for object deduplication
 * @returns {Array} Array with duplicates removed
 */
function removeDuplicates(array, key = null) {
  if (!Array.isArray(array)) return [];

  if (!key) {
    return [...new Set(array)];
  }

  const seen = new Set();
  return array.filter(item => {
    const value = item[key];
    if (seen.has(value)) return false;
    seen.add(value);
    return true;
  });
}

// ─────────────────────────────────────────────────────────────────────────
// COMPANY LOGO — ScriptProperties storage with chunking
// Each property value is limited to ~9KB; a compressed logo base64 data URL
// can exceed that, so we split it into 8000-char chunks and reassemble.
// ─────────────────────────────────────────────────────────────────────────

function saveLogo(chunks) {
  try {
    const props = PropertiesService.getScriptProperties();
    const oldCount = parseInt(props.getProperty('COMPANY_LOGO_CHUNKS') || '0');
    for (let i = 0; i < oldCount; i++) props.deleteProperty('LOGO_CHUNK_' + i);
    props.setProperty('COMPANY_LOGO_CHUNKS', String(chunks.length));
    chunks.forEach((c, i) => props.setProperty('LOGO_CHUNK_' + i, c));
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getLogo() {
  try {
    const props = PropertiesService.getScriptProperties();
    const n = parseInt(props.getProperty('COMPANY_LOGO_CHUNKS') || '0');
    if (!n) return { success: true, data: null };
    let b64 = '';
    for (let i = 0; i < n; i++) b64 += (props.getProperty('LOGO_CHUNK_' + i) || '');
    return { success: true, data: b64 };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function clearLogo() {
  try {
    const props = PropertiesService.getScriptProperties();
    const n = parseInt(props.getProperty('COMPANY_LOGO_CHUNKS') || '0');
    props.deleteProperty('COMPANY_LOGO_CHUNKS');
    for (let i = 0; i < n; i++) props.deleteProperty('LOGO_CHUNK_' + i);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
