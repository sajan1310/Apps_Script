/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_items.gs — ITEMS MASTER MODULE
 * 
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Complete CRUD operations for Items Master including:
 * - Retrieval with vendor rates (getItemsData)
 * - Creation and updates (saveItem)
 * - Deletion (deleteItem)
 * - Vendor management (storing in grid, not notes)
 * - Duplicate detection and prevention
 * - Document-level locking for concurrency safety
 * 
 * Sheet Layout:
 * ───────────────────────────────────────────────────────────────────────────
 * Row 1:     Headers (Item Name, Size, Remarks, Narration, Specification, Vendor pairs...)
 * Rows 2+:   Data rows
 * 
 * Columns:
 * - A (1):   Item Name (unique + size = composite key)
 * - B (2):   Size (unique + name = composite key)
 * - C (3):   Remarks
 * - D (4):   Narration
 * - E (5):   Specification
 * - F+ (6+): Vendor pairs: [Vendor Name | Vendor Rate] repeating
 *            Example: F=Vendor1, G=Rate1, H=Vendor2, I=Rate2, etc.
 * 
 * Why Vendor Pairs in Grid vs Cell Notes:
 * ───────────────────────────────────────────────────────────────────────────
 * Cell notes are:
 * - NOT indexed or searchable
 * - NOT exported to CSV/Excel
 * - NOT accessible via formulas
 * - NOT visible to pivot tables
 * - Lost on most import/export flows
 * 
 * Grid cells are:
 * - Indexed and searchable
 * - Exported with data
 * - Accessible via formulas
 * - Visible to pivot tables
 * - Preserved in all export flows
 * 
 * Therefore: Vendor data lives in the grid as adjacent column pairs
 * 
 * Key Features:
 * ───────────────────────────────────────────────────────────────────────────
 * - Composite key: (name, size) — same item name allowed with different sizes
 * - Efficient duplicate detection: single range read detects duplicates
 * - Vendor pair storage: Name | Rate in adjacent columns
 * - Rate validation: non-numeric rates are skipped with warning
 * - No implicit header regeneration: safe to call saveItem repeatedly
 * - Set up once with initItemsSheet()
 * 
 * Dependencies:
 * ───────────────────────────────────────────────────────────────────────────
 * - config.gs (APP_CONFIG, ITEMS_COL)
 * - utils.gs (getSheet, buildResponse, sanitizeString)
 * 
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ─────────────────────────────────────────────────────────────────────────
// CONSTANTS & CONFIGURATION
// ─────────────────────────────────────────────────────────────────────────

/**
 * Items lock timeout (milliseconds)
 * Prevents deadlocks and concurrent modification conflicts
 * Timeout: 15 seconds (sufficient for most operations)
 */
const ITEMS_LOCK_TIMEOUT_MS = 15000;

/**
 * Number of columns per vendor pair (name + rate)
 * Used to calculate vendor column positions
 */
const VENDOR_PAIR_WIDTH = 2;

/**
 * Default number of vendor pair columns to pre-create during sheet initialization
 */
const DEFAULT_VENDOR_SLOTS = 10;

/**
 * Minimum vendor rate to store (prevents empty rates from being saved)
 */
const MIN_VENDOR_RATE = 0.01;

// ─────────────────────────────────────────────────────────────────────────
// VALIDATION HELPERS
// ─────────────────────────────────────────────────────────────────────────

/**
 * Validates that an item name is non-empty and safe
 * 
 * @param {*} name - Item name to validate
 * @returns {string} Validated item name
 * @throws {Error} If invalid
 * 
 * @example
 * const name = validateItemName('Brake Pads');  // "Brake Pads"
 * const name = validateItemName('');            // throws
 */
function _validateItemName(name) {
  const s = String(name || '').trim();
  if (!s) {
    throw new Error('Item name must not be empty.');
  }

  // Reuse injection check from utils
  return sanitizeString(s, 'itemName');
}

/**
 * Validates vendor rate is numeric and within acceptable range
 * 
 * @param {*} rate - Rate to validate
 * @returns {number} Validated rate
 * @throws {Error} If invalid
 */
function _validateVendorRate(rate) {
  const r = Number(rate);

  if (!Number.isFinite(r) || r < 0) {
    throw new Error('Vendor rate must be a non-negative number.');
  }

  return r;
}

/**
 * Validates/normalizes an item's Base Unit. Blank defaults to 'Pcs' (matches
 * the system's historical implicit assumption).
 *
 * @param {*} value - Raw base unit input
 * @returns {string} Validated base unit name
 */
function _validateBaseUnit(value) {
  const s = String(value || '').trim();
  return s ? sanitizeString(s, 'baseUnit') : 'Pcs';
}

/**
 * Validates/normalizes an item's Purchase Unit. Blank defaults to the
 * item's Base Unit (i.e. "vendor quotes in whatever we stock in").
 *
 * @param {*} value - Raw purchase unit input
 * @param {string} baseUnit - Already-validated base unit, used as the default
 * @returns {string} Validated purchase unit name
 */
function _validatePurchaseUnit(value, baseUnit) {
  const s = String(value || '').trim();
  return s ? sanitizeString(s, 'purchaseUnit') : baseUnit;
}

/**
 * Validates an optional Weight per Base Unit (grams). Only meaningful when
 * Purchase Unit's family differs from Base Unit's family (e.g. bought by Kg,
 * used by Pcs) — blank/0 means "not applicable."
 *
 * @param {*} value - Raw weight input
 * @returns {number} Validated non-negative weight in grams (0 if not set)
 * @throws {Error} If a non-empty value is not a valid non-negative number
 */
function _validateWeightPerBaseUnit(value) {
  if (value === undefined || value === null || value === '') {
    return 0;
  }
  const n = Number(value);
  if (!Number.isFinite(n) || n < 0) {
    throw new Error('Weight per Base Unit must be a non-negative number.');
  }
  return n;
}

/**
 * Validates an optional initial stock value supplied when creating a new item.
 * Blank/undefined defaults to 0. Negative values are allowed on purpose — an
 * over-consumed item (e.g. production issued more than was on hand) can leave
 * stock negative, and the user should be able to see and later correct that
 * rather than being blocked from saving.
 *
 * @param {*} value - Raw initial stock input
 * @returns {number} Validated initial stock
 * @throws {Error} If a non-empty value is not a valid number
 */
function _validateInitialStock(value) {
  if (value === undefined || value === null || value === '') {
    return 0;
  }

  const n = Number(value);
  if (!Number.isFinite(n)) {
    throw new Error('Initial stock must be a valid number.');
  }

  return n;
}

// ─────────────────────────────────────────────────────────────────────────
// COLUMN OFFSET CALCULATIONS
// ─────────────────────────────────────────────────────────────────────────

/**
 * Calculates the zero-based column offset for a vendor's NAME column
 * within a data row, given the vendor's 0-based position in vendors array
 * 
 * Formula: (VENDORS_START - 1) + vendorIndex * VENDOR_PAIR_WIDTH
 * 
 * @param {number} vendorIndex - 0-based vendor position
 * @returns {number} Zero-based column offset in data row
 * 
 * @example
 * vendorNameOffset(0)  // First vendor name column
 * vendorNameOffset(1)  // Second vendor name column
 * 
 * @private
 */
function _vendorNameOffset(vendorIndex) {
  return (ITEMS_COL.VENDOR_DATA - 1) + vendorIndex * VENDOR_PAIR_WIDTH;
}

/**
 * Calculates the zero-based column offset for a vendor's RATE column
 * 
 * @param {number} vendorIndex - 0-based vendor position
 * @returns {number} Zero-based column offset in data row
 * 
 * @private
 */
function _vendorRateOffset(vendorIndex) {
  return _vendorNameOffset(vendorIndex) + 1;
}

/**
 * Parses vendor pairs (Name | Rate) out of a raw Items Master row, using the
 * same stop/skip rules as getItemsData (stop on first empty pair, skip
 * non-numeric rates, require rate >= MIN_VENDOR_RATE).
 *
 * @param {Array} row - Full row values as returned by getValues()[0]
 * @returns {Array<{vendor: string, rate: number}>}
 *
 * @private
 */
function _parseVendorPairsFromRow(row) {
  const vendors = [];
  const vendorStart = ITEMS_COL.VENDOR_DATA - 1;

  for (let v = 0; vendorStart + v * VENDOR_PAIR_WIDTH + 1 < row.length; v++) {
    const nameOffset = vendorStart + v * VENDOR_PAIR_WIDTH;
    const rateOffset = nameOffset + 1;

    const vendorName = String(row[nameOffset] || '').trim();
    const rateRaw = row[rateOffset];
    const rate = Number(rateRaw);

    if (!vendorName && !rateRaw) break;
    if (rateRaw !== '' && rateRaw !== null && isNaN(rate)) continue;

    if (rate >= MIN_VENDOR_RATE && vendorName) {
      vendors.push({ vendor: vendorName, rate: rate });
    }
  }

  return vendors;
}

// ─────────────────────────────────────────────────────────────────────────
// ITEM LOOKUP & DUPLICATE DETECTION
// ─────────────────────────────────────────────────────────────────────────

/**
 * Finds item row(s) matching name + size combination (case-insensitive)
 * 
 * Single-pass algorithm:
 * 1. Read all name+size columns
 * 2. Scan for matches (case-insensitive)
 * 3. Return first match row AND duplicate flag
 * 4. Optionally exclude a specific row from duplicate detection (for edits)
 * 
 * Design:
 * - Composite key (name + size) allows same item name with different sizes
 * - Callers never need two separate reads
 * - Efficient for both new and edit operations
 * 
 * @param {Sheet} sheet - Items sheet object
 * @param {string} name - Item name to find
 * @param {string} size - Item size to find
 * @param {number} [excludeRow] - Optional 1-based row to exclude from duplicate detection
 *
 * Note: Name + Size is the sole identity key. Narration is descriptive
 * metadata only and is never used to distinguish items.
 * 
 * @returns {Object} Result object
 *   @returns {number} matchRow - 1-based row number of first match, or -1 if not found
 *   @returns {boolean} hasDuplicate - True if duplicate exists (excluding excludeRow if provided)
 * 
 * @example
 * // Find existing item
 * const result = findItemRow(sheet, 'Brake Pads', 'Standard');
 * if (result.matchRow !== -1) Log.info('Found at row', result.matchRow);
 * 
 * // Check for duplicates when renaming (exclude original row)
 * const result = findItemRow(sheet, newName, newSize, existingRow);
 * if (result.hasDuplicate) throw new Error('Duplicate name+size');
 * 
 * @private
 */
function _findItemRow(sheet, name, size, excludeRow) {
  const lastRow = sheet.getLastRow();

  // No data rows
  if (lastRow < 2) {
    return { matchRow: -1, hasDuplicate: false };
  }

  // Fetch name + size columns in single batch read
  const data = sheet.getRange(
    2,
    ITEMS_COL.ITEM_NAME,
    lastRow - 1,
    ITEMS_COL.SIZE - ITEMS_COL.ITEM_NAME + 1
  ).getValues();

  const targetName = String(name || '').trim().toLowerCase();
  const targetSize = String(size || '').trim().toLowerCase();

  let matchRow = -1;
  let hasDuplicate = false;

  // Scan for matches
  for (let i = 0; i < data.length; i++) {
    const rowName = String(data[i][ITEMS_COL.ITEM_NAME - 1] || '').trim().toLowerCase();
    const rowSize = String(data[i][ITEMS_COL.SIZE - 1] || '').trim().toLowerCase();

    if (rowName === targetName && rowSize === targetSize) {
      const sheetRow = i + 2;  // Convert to 1-based sheet row

      // First match found
      if (matchRow === -1) {
        matchRow = sheetRow;
      }

      // Check for duplicates (excluding specified row if provided)
      if (excludeRow !== undefined && sheetRow !== excludeRow) {
        hasDuplicate = true;
      } else if (excludeRow === undefined && sheetRow !== matchRow) {
        hasDuplicate = true;
      }
    }
  }

  return { matchRow, hasDuplicate };
}

/**
 * Checks whether any row in the Items sheet still has the given name + size
 * (case-insensitive). Used to decide whether a Stock entry should be
 * renamed/removed when an Items Master record changes.
 *
 * @param {Sheet} sheet - Items sheet object
 * @param {string} name - Item name to check
 * @param {string} size - Item size to check
 * @returns {boolean} True if at least one row matches
 *
 * @private
 */
function _nameSizeStillExists(sheet, name, size) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const data = sheet.getRange(2, ITEMS_COL.ITEM_NAME, lastRow - 1, 2).getValues();
  const targetName = String(name || '').trim().toLowerCase();
  const targetSize = String(size || '').trim().toLowerCase();

  for (let i = 0; i < data.length; i++) {
    const rowName = String(data[i][0] || '').trim().toLowerCase();
    const rowSize = String(data[i][1] || '').trim().toLowerCase();
    if (rowName === targetName && rowSize === targetSize) return true;
  }
  return false;
}

// ─────────────────────────────────────────────────────────────────────────
// SHEET INITIALIZATION
// ─────────────────────────────────────────────────────────────────────────

/**
 * Writes the header row for the Items sheet
 * Called during initialization only, not during saves
 * 
 * Headers include:
 * - Fixed fields (Name, Size, Remarks, Narration, Specification)
 * - Vendor pair headers (Vendor 1, Rate 1, Vendor 2, Rate 2, etc.)
 * 
 * @param {Sheet} sheet - Items sheet object
 * @param {number} maxVendorCount - Number of vendor pair columns to create
 * 
 * @private
 */
function _writeItemHeaders(sheet, maxVendorCount) {
  try {
    const fixedHeaders = [
      'Item Name',
      'Size',
      'Remarks',
      'Narration',
      'Specification',
      'Base Unit',
      'Purchase Unit',
      'Weight per Base Unit (g)'
    ];

    // Build vendor pair headers
    const vendorHeaders = [];
    for (let i = 0; i < maxVendorCount; i++) {
      vendorHeaders.push(`Vendor ${i + 1}`, `Rate ${i + 1}`);
    }

    // Combine all headers
    const allHeaders = fixedHeaders.concat(vendorHeaders);

    // Write to sheet and format
    sheet
      .getRange(1, 1, 1, allHeaders.length)
      .setValues([allHeaders])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    Logger.log(
      `[_writeItemHeaders] Created headers for ${maxVendorCount} vendor slots`
    );
  } catch (error) {
    Log.error('[_writeItemHeaders] Error:', error.message);
    throw error;
  }
}

/**
 * Builds a "name|size" -> {baseUnit, purchaseUnit, weightPerBaseUnit} lookup
 * map from a single batch read of Items Master. Used by module_po.js /
 * module_bill.js to convert a line item's entered qty/price into the item's
 * Base Unit before writing PO/Bill rows, without re-reading the Items sheet
 * once per line.
 *
 * Unknown (name, size) pairs are not included — callers should default to
 * {baseUnit: 'Pcs', purchaseUnit: 'Pcs', weightPerBaseUnit: 0} for those
 * (e.g. a brand-new item being registered via this very PO/Bill line).
 *
 * @returns {Object} { [nameLower|sizeLower]: {baseUnit, purchaseUnit, weightPerBaseUnit} }
 */
function _getItemUnitInfoMap() {
  const map = {};
  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
    const lastRow = sheet ? sheet.getLastRow() : 0;
    if (!sheet || lastRow < 2) return map;

    const data = sheet.getRange(2, ITEMS_COL.ITEM_NAME, lastRow - 1, ITEMS_COL.WEIGHT_PER_BASE_UNIT - ITEMS_COL.ITEM_NAME + 1).getValues();
    data.forEach(row => {
      const name = String(row[ITEMS_COL.ITEM_NAME - 1] || '').trim();
      if (!name) return;
      const size = String(row[ITEMS_COL.SIZE - 1] || '').trim();
      const baseUnit = String(row[ITEMS_COL.BASE_UNIT - 1] || '').trim() || 'Pcs';
      const purchaseUnit = String(row[ITEMS_COL.PURCHASE_UNIT - 1] || '').trim() || baseUnit;
      const weightPerBaseUnit = Number(row[ITEMS_COL.WEIGHT_PER_BASE_UNIT - 1]) || 0;

      map[name.toLowerCase() + '|' + size.toLowerCase()] = { baseUnit, purchaseUnit, weightPerBaseUnit };
    });
  } catch (e) {
    // Items sheet not found/accessible — callers fall back to Pcs defaults.
  }
  return map;
}

/**
 * Builds a "name|size" -> narration lookup map from a single batch read of
 * Items Master. Items Master is the source of truth for narration, so any
 * consumer that keeps its own snapshot of it (Production's
 * COMPONENTS_CONSUMED / CUSTOM_COMPONENTS JSON) resolves through this rather
 * than trusting a value copied in at save time — see
 * backfillProductionNarrationFromItems and saveProduction.
 *
 * Only non-blank narrations are included. A blank one carries no information
 * and must never overwrite a note somebody typed by hand into a consumer, so
 * omitting it here lets every caller treat "absent from the map" as
 * "keep what you have" without a second check.
 *
 * @returns {Object} { [nameLower|sizeLower]: narration }
 */
function _getItemNarrationMap() {
  const map = {};
  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
    const lastRow = sheet ? sheet.getLastRow() : 0;
    if (!sheet || lastRow < 2) return map;

    const data = sheet.getRange(2, ITEMS_COL.ITEM_NAME, lastRow - 1, ITEMS_COL.NARRATION).getValues();
    data.forEach(row => {
      const name = String(row[ITEMS_COL.ITEM_NAME - 1] || '').trim();
      if (!name) return;
      const narration = String(row[ITEMS_COL.NARRATION - 1] || '').trim();
      if (!narration) return;
      const size = String(row[ITEMS_COL.SIZE - 1] || '').trim();
      map[name.toLowerCase() + '|' + size.toLowerCase()] = narration;
    });
  } catch (e) {
    // Items sheet not found/accessible — callers keep their stored narration.
  }
  return map;
}

/**
 * Looks up unit info for a single (name, size) pair from a map built by
 * _getItemUnitInfoMap(). Falls back to {baseUnit: 'Pcs', purchaseUnit: 'Pcs',
 * weightPerBaseUnit: 0} for items not found (new/unregistered items).
 *
 * @param {Object} itemUnitMap - Result of _getItemUnitInfoMap()
 * @param {string} name
 * @param {string} size
 * @returns {{baseUnit: string, purchaseUnit: string, weightPerBaseUnit: number}}
 */
function _lookupItemUnitInfo(itemUnitMap, name, size) {
  const key = String(name || '').trim().toLowerCase() + '|' + String(size || '').trim().toLowerCase();
  return itemUnitMap[key] || { baseUnit: 'Pcs', purchaseUnit: 'Pcs', weightPerBaseUnit: 0 };
}

/**
 * Converts a vendor rate quoted in a SOURCE item's Purchase Unit into the
 * equivalent rate quoted in a TARGET item's Purchase Unit, via the target's
 * Base Unit as the common reference (used when merging two Item Master rows
 * — see mergeItemEdit/_mergeItemIdentities). Without this, copying a
 * source's raw rate number straight into the target row silently
 * misinterprets it whenever the two items were set up with different
 * Purchase/Base Units (e.g. a rate quoted per Dozen copied onto a row whose
 * Purchase Unit is Pcs — getItemsData would then read it back as 12x too
 * high). If sourceInfo.purchaseUnit === targetInfo.purchaseUnit, this is a
 * no-op (returns `rate` unchanged), matching prior behavior exactly for the
 * common case.
 *
 * @param {number} rate - Vendor rate as quoted, per sourceInfo.purchaseUnit
 * @param {{baseUnit: string, purchaseUnit: string, weightPerBaseUnit: number}} sourceInfo
 * @param {{baseUnit: string, purchaseUnit: string, weightPerBaseUnit: number}} targetInfo
 * @param {Object} unitsMap - _getUnitsMap() result
 * @returns {number} Rate quoted per targetInfo.purchaseUnit
 * @throws {Error} If the units can't be reconciled (propagates from convertQtyToBaseUnit/convertRateToBaseUnit)
 */
function _convertVendorRateAcrossItems(rate, sourceInfo, targetInfo, unitsMap) {
  if (sourceInfo.purchaseUnit === targetInfo.purchaseUnit) return rate;
  const ratePerTargetBaseUnit = convertRateToBaseUnit(rate, sourceInfo.purchaseUnit, targetInfo, unitsMap);
  const baseUnitsPerTargetPurchaseUnit = convertQtyToBaseUnit(1, targetInfo.purchaseUnit, targetInfo, unitsMap);
  return ratePerTargetBaseUnit * baseUnitsPerTargetPurchaseUnit;
}

// ─────────────────────────────────────────────────────────────────────────
// getItemsData
// ─────────────────────────────────────────────────────────────────────────

/**
 * getItemsData()
 * 
 * Retrieves all items from sheet with their vendor rates
 * 
 * Data Transformation:
 * 1. Fetch all rows from Items sheet (starting row 2)
 * 2. For each row:
 *    - Extract fixed fields (name, size, remarks, etc.)
 *    - Parse vendor pairs from adjacent columns (Name | Rate)
 *    - Skip vendors with empty name or missing rate
 *    - Log warnings for non-numeric rates (don't silently drop)
 * 3. Return structured array of items
 * 
 * Vendor Pair Handling:
 * - Reads from column F onwards (adjacent pairs: Name, Rate, Name, Rate, ...)
 * - Stops reading when both name and rate are empty
 * - Skips vendors with non-numeric rates (with warning log)
 * - Requires rate > 0 to include vendor
 * 
 * Performance:
 * - Single batch read of all data
 * - O(n*m) where n=items, m=vendor pairs per item
 * - Returns only non-empty items
 * 
 * @returns {Object} API response
 *   - success: true/false
 *   - data: Array of item objects (sorted by name)
 *   - message: Status message
 * 
 * Item Object Structure:
 * {
 *   name: "Brake Pads",
 *   size: "Standard",
 *   remarks: "General use",
 *   narration: "Front brake pad set",
 *   specification: "OEM compatible",
 *   vendors: [
 *     { vendor: "Avon Cycles", rate: 120.50 },
 *     { vendor: "DT Swiss", rate: 118.00 }
 *   ]
 * }
 */
function getItemsData() {
  return getCachedListResponse(MASTER_DATA_CACHE_KEYS.ITEMS, () => _getItemsDataUncached());
}

/**
 * Maps one raw Items sheet row into the record shape the client expects.
 * Shared by _getItemsDataUncached's bulk read and saveItem's single fresh-
 * row read-back (used to patch just the saved item into the client's
 * already-loaded table in place instead of a full list reload). Returns
 * null for a blank/empty row (mirrors the bulk loop's `continue`).
 * @private
 */
function _mapItemRow(row, unitsMap, rowIdxForLog) {
  const name = String(row[ITEMS_COL.ITEM_NAME - 1] || '').trim();
  if (!name) return null;

  const baseUnit = String(row[ITEMS_COL.BASE_UNIT - 1] || '').trim() || 'Pcs';
  const purchaseUnit = String(row[ITEMS_COL.PURCHASE_UNIT - 1] || '').trim() || baseUnit;
  const weightPerBaseUnit = Number(row[ITEMS_COL.WEIGHT_PER_BASE_UNIT - 1]) || 0;

  const itemRecord = {
    name: name,
    size: String(row[ITEMS_COL.SIZE - 1] || '').trim(),
    remarks: String(row[ITEMS_COL.REMARKS - 1] || '').trim(),
    narration: String(row[ITEMS_COL.NARRATION - 1] || '').trim(),
    specification: String(row[ITEMS_COL.SPECIFICATION - 1] || '').trim(),
    baseUnit: baseUnit,
    purchaseUnit: purchaseUnit,
    weightPerBaseUnit: weightPerBaseUnit,
    vendors: []
  };

  // Read vendor pairs: F = name, G = rate, H = name, I = rate, etc.
  const vendorStart = ITEMS_COL.VENDOR_DATA - 1;  // Convert to 0-based

  for (let v = 0; vendorStart + v * VENDOR_PAIR_WIDTH + 1 < row.length; v++) {
    const nameOffset = vendorStart + v * VENDOR_PAIR_WIDTH;
    const rateOffset = nameOffset + 1;

    const vendorName = String(row[nameOffset] || '').trim();
    const rateRaw = row[rateOffset];
    const rate = Number(rateRaw);

    // Stop if both name and rate are empty
    if (!vendorName && !rateRaw) {
      break;
    }

    // Warn if rate is non-numeric but not empty
    if (rateRaw !== '' && rateRaw !== null && isNaN(rate)) {
      Log.warn(
        `[_mapItemRow] Non-numeric rate "${rateRaw}" for vendor "${vendorName}" ` +
        `on item "${name}" (sheet row ${rowIdxForLog}) — skipped.`
      );
      continue;
    }

    // Include vendor only if rate is positive and name exists
    if (rate >= MIN_VENDOR_RATE && vendorName) {
      let ratePerBaseUnit = rate;
      try {
        ratePerBaseUnit = convertRateToBaseUnit(rate, purchaseUnit, itemRecord, unitsMap);
      } catch (e) {
        Log.warn(
          `[_mapItemRow] Could not convert rate for vendor "${vendorName}" on item ` +
          `"${name}" (sheet row ${rowIdxForLog}): ${e.message}`
        );
      }

      itemRecord.vendors.push({
        vendor: vendorName,
        rate: rate,
        ratePerBaseUnit: ratePerBaseUnit
      });
    }
  }

  return itemRecord;
}

function _getItemsDataUncached() {
  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
    if (!sheet) {
      return buildResponse(false, null, 'Items sheet not found.');
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    // No data rows (only headers or empty)
    if (lastRow < 2 || lastCol < ITEMS_COL.ITEM_NAME) {
      return buildResponse(true, []);
    }

    // Fetch all data in single batch read
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const items = [];
    const unitsMap = _getUnitsMap();

    for (let i = 0; i < data.length; i++) {
      const itemRecord = _mapItemRow(data[i], unitsMap, i + 2);
      if (itemRecord) items.push(itemRecord);
    }

    // Sort by name for consistent ordering
    items.sort((a, b) => a.name.localeCompare(b.name));

    return buildResponse(true, items);
  } catch (error) {
    Log.error('[getItemsData] Error:', error.message);
    logAction('ERROR', 'getItemsData', 'module_items', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to load items: ' + error.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────
// saveItem
// ─────────────────────────────────────────────────────────────────────────

/**
 * saveItem(formData)
 * 
 * Creates a new item or updates an existing item
 * 
 * Operations (Create):
 * 1. Validate all input data
 * 2. Check for duplicate (name + size)
 * 3. Append new row to sheet
 * 4. Write fixed fields (A-E)
 * 5. Write vendor pairs (F onwards)
 * 6. Flush changes
 * 
 * Operations (Edit):
 * 1. Validate all input data
 * 2. Locate existing row by original (name, size)
 * 3. If renaming or resizing: check for collision with other row
 * 4. Clear vendor columns from existing row
 * 5. Update fixed fields
 * 6. Write vendor pairs
 * 7. Flush changes
 * 
 * Duplicate Detection:
 * - Composite key: (name, size)
 * - Same item name allowed with different sizes
 * - Single read detects duplicates (no two-pass overhead)
 * - On edit: allows keeping original (name, size) unchanged
 * 
 * Vendor Storage:
 * - Stored as adjacent column pairs (Name | Rate)
 * - NOT in cell notes (notes lost on export/import)
 * - Each vendor occupies 2 columns
 * - First vendor pair starts at column F (6)
 * 
 * Headers:
 * - NOT regenerated on each save (safe for repeated calls)
 * - Initialize once with initItemsSheet()
 * 
 * @param {Object} formData - Form submission data
 *   @param {string} formData.itemName - Item name
 *   @param {string} formData.itemSize - Item size/variant
 *   @param {string} formData.itemRemarks - Optional remarks
 *   @param {string} formData.itemNarration - Optional description
 *   @param {string} formData.itemSpec - Optional specification
 *   @param {Array|string} formData.vendors - Vendors array or JSON string
 *     Each vendor: { vendor, rate }
 *   @param {string} [formData.originalName] - For edits: original item name
 *   @param {string} [formData.originalSize] - For edits: original item size
 * 
 * @returns {Object} API response
 *   - success: true/false
 *   - data: { name, size } if successful
 *   - message: Status/error message
 */
function saveItem(formData) {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(ITEMS_LOCK_TIMEOUT_MS)) {
    return buildResponse(
      false,
      null,
      'System is busy. Please try again in a moment.'
    );
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
    if (!sheet) {
      throw new Error(`Sheet "${APP_CONFIG.SHEETS.ITEMS}" not found.`);
    }

    // ─────────────────────────────────────────────────────────────────
    // Parse and validate vendors
    // ─────────────────────────────────────────────────────────────────
    let vendors;
    try {
      vendors = typeof formData.vendors === 'string'
        ? JSON.parse(formData.vendors)
        : (formData.vendors || []);
    } catch (error) {
      return buildResponse(
        false,
        null,
        'Invalid vendors data: could not parse JSON.'
      );
    }

    if (!Array.isArray(vendors)) {
      vendors = [];
    }

    // ─────────────────────────────────────────────────────────────────
    // Validate and sanitize fixed fields
    // ─────────────────────────────────────────────────────────────────
    const newName = _validateItemName(formData.itemName);
    const newSize = sanitizeString(formData.itemSize || '', 'itemSize');
    const remarks = sanitizeString(formData.itemRemarks || '', 'itemRemarks');
    const narration = sanitizeString(formData.itemNarration || '', 'itemNarration');
    const spec = sanitizeString(formData.itemSpec || '', 'itemSpec');
    const initialStock = _validateInitialStock(formData.itemInitialStock);
    const baseUnit = _validateBaseUnit(formData.itemBaseUnit);
    const purchaseUnit = _validatePurchaseUnit(formData.itemPurchaseUnit, baseUnit);
    const weightPerBaseUnit = _validateWeightPerBaseUnit(formData.itemWeightPerBaseUnit);

    // ─────────────────────────────────────────────────────────────────
    // Determine create vs edit
    // ─────────────────────────────────────────────────────────────────
    const isEdit = !!formData.originalName;
    const originalName = isEdit ? _validateItemName(formData.originalName) : newName;
    const originalSize = isEdit
      ? sanitizeString(String(formData.originalSize || ''), 'originalSize')
      : newSize;

    // ─────────────────────────────────────────────────────────────────
    // Find existing item and check for duplicates. Original + new
    // (name,size) are looked up from ONE read of the sheet's Name+Size
    // columns — they're two different lookups against the exact same
    // snapshot, so calling _findItemRow() a second time for the new combo
    // would just re-read the identical range the first call already fetched.
    // ─────────────────────────────────────────────────────────────────
    const _itemsLastRow = sheet.getLastRow();
    const _nameSizeData = _itemsLastRow >= 2
      ? sheet.getRange(2, ITEMS_COL.ITEM_NAME, _itemsLastRow - 1, ITEMS_COL.SIZE - ITEMS_COL.ITEM_NAME + 1).getValues()
      : [];
    const _findRowInSnapshot = (name, size) => {
      const targetName = String(name || '').trim().toLowerCase();
      const targetSize = String(size || '').trim().toLowerCase();
      for (let i = 0; i < _nameSizeData.length; i++) {
        const rowName = String(_nameSizeData[i][ITEMS_COL.ITEM_NAME - 1] || '').trim().toLowerCase();
        const rowSize = String(_nameSizeData[i][ITEMS_COL.SIZE - 1] || '').trim().toLowerCase();
        if (rowName === targetName && rowSize === targetSize) return i + 2;
      }
      return -1;
    };

    const existingRow = _findRowInSnapshot(originalName, originalSize);

    // For rename/resize edits, check new combo doesn't collide
    if (newName !== originalName || newSize !== originalSize) {
      const dupeRow = _findRowInSnapshot(newName, newSize);
      if (dupeRow !== -1 && dupeRow !== existingRow) {
        const dupeMsg = `Duplicate: "${newName}" with size "${newSize}" already exists.`;

        // On edit (not new-item creation), offer the caller enough info to
        // present a "merge into existing item" confirmation instead of just
        // blocking the save.
        if (isEdit) {
          const targetVendors = _parseVendorPairsFromRow(
            sheet.getRange(dupeRow, 1, 1, sheet.getLastColumn()).getValues()[0]
          );

          let targetStock = 0;
          try {
            const stockSheet = getSheet(APP_CONFIG.SHEETS.STOCK);
            const stockRow = _findStockRow(stockSheet, newName, newSize);
            if (stockRow !== -1) {
              targetStock = Number(
                stockSheet.getRange(stockRow, STOCK_COL.CURRENT_STOCK).getValue()
              ) || 0;
            }
          } catch (e) {
            // Stock sheet not found/accessible — report stock as unknown (0).
          }

          // NOTE: buildResponse() always nulls `data` when success is false
          // (utils.js) — bypass it here so the client can actually see the
          // merge info; the {success, data, message} shape is kept identical.
          return {
            success: false,
            data: {
              mergeable: true,
              targetName: newName,
              targetSize: newSize,
              targetStock: targetStock,
              targetVendorCount: targetVendors.length
            },
            message: dupeMsg
          };
        }

        return buildResponse(false, null, dupeMsg);
      }
    } else if (!isEdit) {
      // New item — check for any collision
      if (existingRow !== -1) {
        return buildResponse(
          false,
          null,
          `Duplicate: "${newName}" with size "${newSize}" already exists.`
        );
      }
    }

    // On edit, item MUST exist
    if (isEdit && existingRow === -1) {
      throw new Error(
        `Item "${originalName}" (size: "${originalSize}") not found. ` +
        `Edit aborted to prevent data inconsistency.`
      );
    }

    // ─────────────────────────────────────────────────────────────────
    // Determine target row and clear vendor columns (edit only)
    // ─────────────────────────────────────────────────────────────────
    const targetRow = isEdit ? existingRow : sheet.getLastRow() + 1;
    const currentLastCol = sheet.getLastColumn();

    if (isEdit && currentLastCol >= ITEMS_COL.VENDOR_DATA) {
      // Clear vendor columns before writing new ones
      sheet.getRange(
        targetRow,
        ITEMS_COL.VENDOR_DATA,
        1,
        currentLastCol - ITEMS_COL.VENDOR_DATA + 1
      ).clearContent();
    }

    // ─────────────────────────────────────────────────────────────────
    // Write fixed fields
    // ─────────────────────────────────────────────────────────────────
    const fixedData = [newName, newSize, remarks, narration, spec, baseUnit, purchaseUnit, weightPerBaseUnit];
    sheet.getRange(
      targetRow,
      ITEMS_COL.ITEM_NAME,
      1,
      fixedData.length
    ).setValues([fixedData]);

    // ─────────────────────────────────────────────────────────────────
    // Write vendor pairs (Name | Rate) as grid cells
    // ─────────────────────────────────────────────────────────────────
    if (vendors.length > 0) {
      const vendorRow = [];

      vendors.forEach(v => {
        const vendorName = sanitizeString(String(v.vendor || ''), 'vendor.name');
        const rate = _validateVendorRate(v.rate || 0);
        vendorRow.push(vendorName, rate);
      });

      sheet.getRange(
        targetRow,
        ITEMS_COL.VENDOR_DATA,
        1,
        vendorRow.length
      ).setValues([vendorRow]);
    }

    SpreadsheetApp.flush();
    invalidateListCache(MASTER_DATA_CACHE_KEYS.ITEMS);

    // ─────────────────────────────────────────────────────────────────
    // Keep Stock sheet in sync (Item Name + Size as common key)
    // ─────────────────────────────────────────────────────────────────
    if (!isEdit) {
      syncStockForItem('ensure', { name: newName, size: newSize, initialStock });
    } else if (newName !== originalName || newSize !== originalSize) {
      if (_nameSizeStillExists(sheet, originalName, originalSize)) {
        syncStockForItem('ensure', { name: newName, size: newSize, initialStock });
      } else {
        syncStockForItem('rename', {
          oldName: originalName,
          oldSize: originalSize,
          newName,
          newSize
        });
        _propagateItemIdentityChange(originalName, originalSize, newName, newSize);
      }
    } else {
      // Plain edit (identity unchanged) — backfills a missing Stock row for
      // orphaned items (e.g. flagged "No Metadata"). No-ops if a Stock row
      // already exists, so it never clobbers real stock figures.
      syncStockForItem('ensure', { name: newName, size: newSize, initialStock });
    }

    const itemLabel = `"${newName}"${newSize ? ` (${newSize})` : ''}`;
    const msg = isEdit ? `Item ${itemLabel} updated successfully.` : `Item ${itemLabel} added successfully.`;
    logAction(
      isEdit ? 'UPDATE' : 'CREATE',
      APP_CONFIG.SHEETS.ITEMS,
      newName,
      `Size: ${newSize}, Vendors: ${vendors.length}`,
      'SUCCESS'
    );

    // Read this item's own just-written row back (cheap — one row, not the
    // whole sheet) and map it through the same logic getItemsData uses, so
    // the client can patch this one item into its already-loaded list in
    // place instead of a full reload.
    const freshRowValues = sheet.getRange(targetRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const freshItem = _mapItemRow(freshRowValues, _getUnitsMap(), targetRow);

    return buildResponse(true, { name: newName, size: newSize, item: freshItem }, msg);
  } catch (error) {
    Log.error('[saveItem] Error:', error.message);
    logAction('ERROR', 'saveItem', formData.itemName, error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to save item: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────────────────
// deleteItem
// ─────────────────────────────────────────────────────────────────────────

/**
 * @private
 * Returns the subset of {name, size} keys ("nameLower|sizeLower") that are
 * currently referenced elsewhere and therefore unsafe to delete:
 *  - a Product's BOM component row (BOM_COL.ITEM_NAME/SIZE)
 *  - an ITEM-sourced Process Components recipe row
 *    (PROCESS_COMPONENTS_COL.ITEM_NAME/SIZE where SOURCE_TYPE is 'ITEM' — a
 *    POOL-sourced row references an upstream Process's output, not an Items
 *    Master row, so it's excluded)
 * Deleting a still-referenced item would leave those recipes showing a
 * name/size Items Master no longer has any record of (no rate/unit
 * resolvable), and a new Production lot logged off that recipe would try to
 * consume/debit Stock under a (name, size) key that no longer exists. Same
 * "blocked if referenced" pattern as module_process.js's
 * _getProcessIdsInUse.
 * @param {Array<{name:string,size:string}>} items
 * @returns {Set<string>} "nameLower|sizeLower" keys found in use
 */
function _getItemKeysInUse(items) {
  const requested = new Set((items || [])
    .map(it => String(it.name || '').trim().toLowerCase() + '|' + String(it.size || '').trim().toLowerCase())
    .filter(key => key.split('|')[0]));
  const inUse = new Set();
  if (requested.size === 0) return inUse;

  let bomSheet;
  try {
    bomSheet = getSheet(APP_CONFIG.SHEETS.BOM);
  } catch (e) {
    bomSheet = null;
  }
  if (bomSheet) {
    const lastRow = bomSheet.getLastRow();
    if (lastRow >= 2) {
      const rows = bomSheet.getRange(2, 1, lastRow - 1, BOM_COL.SIZE).getValues();
      rows.forEach(row => {
        const key = String(row[BOM_COL.ITEM_NAME - 1] || '').trim().toLowerCase() + '|' + String(row[BOM_COL.SIZE - 1] || '').trim().toLowerCase();
        if (requested.has(key)) inUse.add(key);
      });
    }
  }

  let compSheet;
  try {
    compSheet = getSheet(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);
  } catch (e) {
    compSheet = null;
  }
  if (compSheet) {
    const lastRow = compSheet.getLastRow();
    if (lastRow >= 2) {
      const rows = compSheet.getRange(2, 1, lastRow - 1, PROCESS_COMPONENTS_COL.SOURCE_TYPE).getValues();
      rows.forEach(row => {
        const sourceType = String(row[PROCESS_COMPONENTS_COL.SOURCE_TYPE - 1] || '').trim().toUpperCase();
        if (sourceType === COMPONENT_SOURCE_TYPES.POOL) return;
        const key = String(row[PROCESS_COMPONENTS_COL.ITEM_NAME - 1] || '').trim().toLowerCase() + '|' + String(row[PROCESS_COMPONENTS_COL.SIZE - 1] || '').trim().toLowerCase();
        if (requested.has(key)) inUse.add(key);
      });
    }
  }

  return inUse;
}

/**
 * deleteItem(name, size)
 * 
 * Deletes the row for a given item (name + size combination)
 * 
 * Operations:
 * 1. Validate parameters
 * 2. Acquire document lock
 * 3. Find matching row by (name, size)
 * 4. Delete the row
 * 5. Flush changes
 * 
 * Safety:
 * - Document lock prevents concurrent deletes
 * - Validates parameters before deletion
 * - Requires exact (name, size) match
 * - Returns clear error if item not found
 * - Audit log entry for each delete
 * 
 * @param {string} name - Item name to delete
 * @param {string} size - Item size to delete
 * 
 * @returns {Object} API response
 *   - success: true/false
 *   - data: null
 *   - message: Status/error message
 */
function deleteItem(name, size) {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(ITEMS_LOCK_TIMEOUT_MS)) {
    return buildResponse(
      false,
      null,
      'System is busy. Please try again.'
    );
  }

  try {
    // Validate parameters
    const validName = _validateItemName(name);
    const validSize = sanitizeString(size || '', 'size');

    const sheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
    if (!sheet) {
      throw new Error(`Sheet "${APP_CONFIG.SHEETS.ITEMS}" not found.`);
    }

    // Find matching row
    const { matchRow } = _findItemRow(sheet, validName, validSize);

    if (matchRow === -1) {
      return buildResponse(
        false,
        null,
        `Item "${validName}" (size: "${validSize}") not found.`
      );
    }

    const inUseKey = validName.toLowerCase() + '|' + validSize.toLowerCase();
    if (_getItemKeysInUse([{ name: validName, size: validSize }]).has(inUseKey)) {
      return buildResponse(
        false,
        null,
        `Cannot delete "${validName}"${validSize ? ' (' + validSize + ')' : ''}: referenced by a Product's BOM or a Process recipe.`
      );
    }

    // Delete the row
    sheet.deleteRow(matchRow);
    SpreadsheetApp.flush();
    invalidateListCache(MASTER_DATA_CACHE_KEYS.ITEMS);

    // Remove the matching Stock entry if no other item shares this name + size
    if (!_nameSizeStillExists(sheet, validName, validSize)) {
      syncStockForItem('remove', { name: validName, size: validSize });
    }

    const msg = `Item "${validName}" deleted successfully.`;
    logAction('DELETE', APP_CONFIG.SHEETS.ITEMS, validName, msg, 'SUCCESS');

    return buildResponse(true, null, msg);
  } catch (error) {
    Log.error('[deleteItem] Error:', error.message);
    logAction('ERROR', 'deleteItem', name, error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete item: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────────────────
// IDENTITY-CHANGE PROPAGATION (rename + merge)
// ─────────────────────────────────────────────────────────────────────────

/**
 * Repoints historical Bill Ledger, PO Tracker, BOM, Process Components,
 * Return Ledger, Wastage Log, Issued Stock Log, and Production
 * Components-Consumed rows from an item's old name+size to its new
 * name+size. Called after a rename (saveItem) or a merge (mergeItemEdit) so
 * history doesn't silently drift out of sync with Item Master / Stock — see
 * module_stock.js's _getBilledAndConsumedQtyMaps(), which keys every one of
 * these sheets' historical rows to the Stock sheet by lowercase name+size;
 * an un-backfilled sheet's rows silently stop matching the (renamed) Stock
 * row on the next recalculateStock() and their qty drops out of Current
 * Stock math entirely.
 *
 * Each sheet's backfill is isolated in its own try/catch so a failure in
 * one sheet doesn't prevent the others from being updated.
 *
 * Runs under the caller's existing document lock — does not acquire its own.
 *
 * @param {string} oldName
 * @param {string} oldSize
 * @param {string} newName
 * @param {string} newSize
 *
 * @private
 */
function _propagateItemIdentityChange(oldName, oldSize, newName, newSize) {
  const backfills = [
    { label: 'Bill Ledger', fn: backfillBillItemRefs },
    { label: 'PO Tracker', fn: backfillPOItemRefs },
    { label: 'BOM', fn: backfillBOMItemRefs },
    { label: 'Process Components', fn: backfillProcessComponentItemRefs },
    { label: 'Return Ledger', fn: backfillReturnItemRefs },
    { label: 'Wastage Log', fn: backfillWastageItemRefs },
    { label: 'Issued Stock Log', fn: backfillIssueItemRefs },
    { label: 'Production (Components Consumed)', fn: backfillProductionConsumedItemRefs }
  ];

  backfills.forEach(({ label, fn }) => {
    if (typeof fn !== 'function') return;
    try {
      fn(oldName, oldSize, newName, newSize);
    } catch (error) {
      Log.error(`[_propagateItemIdentityChange] ${label} backfill failed:`, error.message);
      logAction(
        'ERROR',
        '_propagateItemIdentityChange',
        `${oldName}|${oldSize} -> ${newName}|${newSize}`,
        `${label}: ${error.message}`,
        'ERROR'
      );
    }
  });
}

// ─────────────────────────────────────────────────────────────────────────
// ITEM IDENTITY DRIFT DIAGNOSTIC (verify rename/merge cascades)
// ─────────────────────────────────────────────────────────────────────────

/**
 * Builds a lowercase "name|size" -> true lookup of every (Item Name, Size)
 * combination currently in Items Master. Shared by getItemIdentityDriftReport
 * so every sheet-scanner below checks against the exact same snapshot.
 * @returns {Set<string>}
 * @private
 */
function _getValidItemIdentityKeys() {
  const keys = new Set();
  const sheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return keys;

  const data = sheet.getRange(2, ITEMS_COL.ITEM_NAME, lastRow - 1, ITEMS_COL.SIZE - ITEMS_COL.ITEM_NAME + 1).getValues();
  data.forEach(row => {
    const name = String(row[0] || '').trim().toLowerCase();
    if (!name) return;
    const size = String(row[1] || '').trim().toLowerCase();
    keys.add(name + '|' + size);
  });
  return keys;
}

/**
 * getItemIdentityDriftReport()
 *
 * Read-only diagnostic — never blocks a save, never auto-fixes anything.
 * Verifies that every (Item Name, Size) reference outside Items Master
 * still resolves to a row that actually exists there. Same philosophy as
 * module_bom.js's getBomProcessComponentsDrift().
 *
 * Exists because a rename/merge cascade (see _propagateItemIdentityChange)
 * only knows how to repoint the sheets it's explicitly wired to backfill —
 * a sheet added later without a matching backfill, or historical rows saved
 * before a given backfill existed, would otherwise drift silently out of
 * sync with Items Master and nobody would notice until Stock math went
 * wrong. This scans every sheet that stores an Items Master (name, size)
 * reference and flags any that don't currently resolve to a real item —
 * whether that's leftover damage from before a fix existed, or a brand new
 * gap in a cascade that hasn't been added yet.
 *
 * POOL-sourced rows (Process Components / Production Components Consumed)
 * are skipped — their "item name" is an upstream process's Output Item
 * Name (Warehouse Pool), not an Items Master reference.
 *
 * @returns {Object} API response; data = [{sheet, context, itemName, size}],
 *   one entry per reference that doesn't resolve to a current Items Master
 *   row.
 */
function getItemIdentityDriftReport() {
  try {
    const validKeys = _getValidItemIdentityKeys();
    const findings = [];

    const check = (sheetLabel, context, itemName, size) => {
      const name = String(itemName || '').trim();
      if (!name) return;
      const key = name.toLowerCase() + '|' + String(size || '').trim().toLowerCase();
      if (validKeys.has(key)) return;
      findings.push({ sheet: sheetLabel, context: context, itemName: name, size: String(size || '').trim() });
    };

    // Bill Ledger
    try {
      const sheet = getSheet(APP_CONFIG.SHEETS.BILL);
      const startRow = APP_CONFIG.BILL_SETTINGS.DATA_START_ROW;
      const lastRow = sheet.getLastRow();
      if (lastRow >= startRow) {
        const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, BILL_COL.SIZE).getValues();
        data.forEach(row => {
          check('Bill Ledger', `Bill #${row[BILL_COL.BILL_NUMBER - 1] || ''}`, row[BILL_COL.ITEM_NAME - 1], row[BILL_COL.SIZE - 1]);
        });
      }
    } catch (e) { /* sheet not found — nothing to check */ }

    // PO Tracker
    try {
      const sheet = getSheet(APP_CONFIG.SHEETS.PO);
      const startRow = APP_CONFIG.PO_SETTINGS.DATA_START_ROW;
      const lastRow = sheet.getLastRow();
      if (lastRow >= startRow) {
        const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, PO_COL.SIZE).getValues();
        data.forEach(row => {
          check('PO Tracker', `PO #${row[PO_COL.PO_NUMBER - 1] || ''}`, row[PO_COL.ITEM_NAME - 1], row[PO_COL.SIZE - 1]);
        });
      }
    } catch (e) { /* sheet not found — nothing to check */ }

    // BOM
    try {
      const sheet = getSheet(APP_CONFIG.SHEETS.BOM);
      const startRow = APP_CONFIG.BOM_SETTINGS.DATA_START_ROW;
      const lastRow = sheet.getLastRow();
      if (lastRow >= startRow) {
        const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, BOM_COL.SIZE).getValues();
        data.forEach(row => {
          check('BOM', `${row[BOM_COL.PRODUCT_NAME - 1] || ''} (${row[BOM_COL.PRODUCT_ID - 1] || ''})`, row[BOM_COL.ITEM_NAME - 1], row[BOM_COL.SIZE - 1]);
        });
      }
    } catch (e) { /* sheet not found — nothing to check */ }

    // Process Components (ITEM-sourced only — POOL rows reference an
    // upstream process's Output Item Name, not Items Master)
    try {
      const sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const data = sheet.getRange(2, 1, lastRow - 1, PROCESS_COMPONENTS_COL.SOURCE_TYPE).getValues();
        data.forEach(row => {
          const sourceType = String(row[PROCESS_COMPONENTS_COL.SOURCE_TYPE - 1] || '').trim().toUpperCase();
          if (sourceType === COMPONENT_SOURCE_TYPES.POOL) return;
          check('Process Components', `Process ${row[PROCESS_COMPONENTS_COL.PROCESS_ID - 1] || ''}`, row[PROCESS_COMPONENTS_COL.ITEM_NAME - 1], row[PROCESS_COMPONENTS_COL.SIZE - 1]);
        });
      }
    } catch (e) { /* sheet not found — nothing to check */ }

    // Return Ledger
    try {
      const sheet = getSheet(APP_CONFIG.SHEETS.RETURN);
      const startRow = APP_CONFIG.RETURN_SETTINGS.DATA_START_ROW;
      const lastRow = sheet.getLastRow();
      if (lastRow >= startRow) {
        const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, RETURN_COL.SIZE).getValues();
        data.forEach(row => {
          check('Return Ledger', `Return #${row[RETURN_COL.RETURN_NUMBER - 1] || ''}`, row[RETURN_COL.ITEM_NAME - 1], row[RETURN_COL.SIZE - 1]);
        });
      }
    } catch (e) { /* sheet not found — nothing to check */ }

    // Wastage Log
    try {
      const sheet = getSheet(APP_CONFIG.SHEETS.WASTAGE);
      const startRow = APP_CONFIG.WASTAGE_SETTINGS.DATA_START_ROW;
      const lastRow = sheet.getLastRow();
      if (lastRow >= startRow) {
        const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, WASTAGE_COL.SIZE).getValues();
        data.forEach(row => {
          check('Wastage Log', `Wastage ${row[WASTAGE_COL.WASTAGE_ID - 1] || ''}`, row[WASTAGE_COL.ITEM_NAME - 1], row[WASTAGE_COL.SIZE - 1]);
        });
      }
    } catch (e) { /* sheet not found — nothing to check */ }

    // Issued Stock Log
    try {
      const sheet = getSheet(APP_CONFIG.SHEETS.ISSUE);
      const startRow = APP_CONFIG.ISSUE_SETTINGS.DATA_START_ROW;
      const lastRow = sheet.getLastRow();
      if (lastRow >= startRow) {
        const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, ISSUE_COL.SIZE).getValues();
        data.forEach(row => {
          check('Issued Stock Log', `Issue ${row[ISSUE_COL.ISSUE_ID - 1] || ''}`, row[ISSUE_COL.ITEM_NAME - 1], row[ISSUE_COL.SIZE - 1]);
        });
      }
    } catch (e) { /* sheet not found — nothing to check */ }

    // Production — Components Consumed (ITEM-sourced only) and the legacy
    // Custom Components display snapshot
    try {
      const sheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
      const startRow = APP_CONFIG.PRODUCTION_SETTINGS.DATA_START_ROW;
      const lastRow = sheet.getLastRow();
      if (lastRow >= startRow && sheet.getLastColumn() >= PRODUCTION_COL.COMPONENTS_CONSUMED) {
        const numRows = lastRow - startRow + 1;
        const consumedData = sheet.getRange(startRow, PRODUCTION_COL.COMPONENTS_CONSUMED, numRows, 1).getValues();
        const lotNumberData = sheet.getRange(startRow, PRODUCTION_COL.LOT_NUMBER, numRows, 1).getValues();

        for (let i = 0; i < consumedData.length; i++) {
          const raw = String(consumedData[i][0] || '').trim();
          if (!raw) continue;
          let components;
          try {
            components = JSON.parse(raw);
          } catch (e) {
            continue;
          }
          if (!Array.isArray(components)) continue;

          const context = `Lot ${lotNumberData[i][0] || '(row ' + (startRow + i) + ')'}`;
          components.forEach(comp => {
            const sourceType = String(comp.sourceType || '').trim().toUpperCase();
            if (sourceType === COMPONENT_SOURCE_TYPES.POOL) return;
            check('Production (Components Consumed)', context, comp.itemName, comp.size);
          });
        }
      }
    } catch (e) { /* sheet not found — nothing to check */ }

    const counts = findings.reduce((acc, f) => { acc[f.sheet] = (acc[f.sheet] || 0) + 1; return acc; }, {});
    const message = findings.length === 0
      ? 'No drift found — every Item Name/Size reference resolves to a current Items Master row.'
      : `${findings.length} stale reference(s) found: ` +
        Object.keys(counts).map(s => `${counts[s]} in ${s}`).join(', ');

    return buildResponse(true, findings, message);
  } catch (error) {
    Log.error('[getItemIdentityDriftReport] Error:', error.message);
    return buildResponse(false, null, 'Failed to check item reference integrity: ' + error.message);
  }
}

/**
 * fixItemIdentityDriftReference(staleName, staleSize, targetName, targetSize)
 *
 * Repairs one distinct stale reference surfaced by getItemIdentityDriftReport
 * by repointing EVERY row across EVERY referencing sheet (Bill, PO, BOM,
 * Process Components, Return, Wastage, Issue, Production) from the stale
 * (name, size) to a real, currently-existing Items Master (name, size) that
 * the caller identifies — the system cannot infer this on its own, since
 * Items Master keeps no rename history, so the target must be a deliberate
 * human choice from the "Check Reference Integrity" dialog.
 *
 * Reuses the exact same backfill functions the rename/merge cascade already
 * relies on (_propagateItemIdentityChange) — a stale reference and a normal
 * rename are the same repair, just with the "before" identity sourced from
 * the drift report instead of a live saveItem() call.
 *
 * @param {string} staleName - The orphaned name currently on the referencing rows
 * @param {string} staleSize - The orphaned size currently on the referencing rows
 * @param {string} targetName - A name that currently exists in Items Master
 * @param {string} targetSize - A size that currently exists in Items Master (paired with targetName)
 * @returns {Object} API response
 */
function fixItemIdentityDriftReference(staleName, staleSize, targetName, targetSize) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(ITEMS_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again in a moment.');
  }

  try {
    const validStaleName = sanitizeString(staleName || '', 'staleName');
    const validStaleSize = sanitizeString(staleSize || '', 'staleSize');
    if (!validStaleName) {
      return buildResponse(false, null, 'A stale item name is required.');
    }

    const validTargetName = _validateItemName(targetName);
    const validTargetSize = sanitizeString(targetSize || '', 'targetSize');

    const itemsSheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
    if (!itemsSheet) {
      throw new Error(`Sheet "${APP_CONFIG.SHEETS.ITEMS}" not found.`);
    }
    const { matchRow } = _findItemRow(itemsSheet, validTargetName, validTargetSize);
    if (matchRow === -1) {
      return buildResponse(
        false,
        null,
        `Target item "${validTargetName}"${validTargetSize ? ' (size: "' + validTargetSize + '")' : ''} was not found in Items Master. Pick an item that currently exists.`
      );
    }

    _propagateItemIdentityChange(validStaleName, validStaleSize, validTargetName, validTargetSize);

    if (typeof recalculateStock === 'function') {
      recalculateStock();
    }

    const msg = `Repointed all references from "${validStaleName}"${validStaleSize ? ' (' + validStaleSize + ')' : ''} to "${validTargetName}"${validTargetSize ? ' (' + validTargetSize + ')' : ''}.`;
    logAction(
      'UPDATE',
      'Item Reference Integrity Fix',
      `${validStaleName}|${validStaleSize} -> ${validTargetName}|${validTargetSize}`,
      msg,
      'SUCCESS'
    );
    return buildResponse(true, null, msg);
  } catch (error) {
    Log.error('[fixItemIdentityDriftReference] Error:', error.message);
    logAction('ERROR', 'fixItemIdentityDriftReference', `${staleName}|${staleSize}`, error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to fix reference: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────────────────
// mergeItemEdit
// ─────────────────────────────────────────────────────────────────────────

/**
 * mergeItemEdit(formData)
 *
 * Completes an item edit that saveItem rejected as a duplicate
 * (response.data.mergeable === true), by merging the edited item into the
 * existing item it collided with rather than discarding the edit.
 *
 * Source = the item identified by formData.originalName/originalSize
 *          (the row the user was editing).
 * Target = the item identified by formData.itemName/itemSize
 *          (the existing row the edit collided with).
 *
 * Operations:
 * 1. Re-validate and re-resolve source and target rows fresh from the sheet
 *    (never trust client-cached stock/vendor figures).
 * 2. Union vendors by name; on a rate collision, target's existing rate
 *    wins (the user wasn't shown/editing the target's vendor list).
 *    Target's remarks/narration/specification are left untouched.
 * 3. Delete the source row from Items Master.
 * 4. Combine stock: target's Initial/Current Stock += source's, source's
 *    Stock row is removed (syncStockForItem 'merge' action).
 * 5. Backfill Bill Ledger / PO Tracker / BOM rows that referenced the
 *    source's old identity to the target's identity.
 *
 * @param {Object} formData - Same shape as saveItem's formData, with
 *   originalName/originalSize identifying the source row.
 * @returns {Object} API response
 */
function mergeItemEdit(formData) {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(ITEMS_LOCK_TIMEOUT_MS)) {
    return buildResponse(
      false,
      null,
      'System is busy. Please try again in a moment.'
    );
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
    if (!sheet) {
      throw new Error(`Sheet "${APP_CONFIG.SHEETS.ITEMS}" not found.`);
    }

    if (!formData.originalName) {
      throw new Error('Merge requires an existing item to edit.');
    }

    let vendors;
    try {
      vendors = typeof formData.vendors === 'string'
        ? JSON.parse(formData.vendors)
        : (formData.vendors || []);
    } catch (error) {
      return buildResponse(false, null, 'Invalid vendors data: could not parse JSON.');
    }
    if (!Array.isArray(vendors)) {
      vendors = [];
    }

    const newName = _validateItemName(formData.itemName);
    const newSize = sanitizeString(formData.itemSize || '', 'itemSize');

    const originalName = _validateItemName(formData.originalName);
    const originalSize = sanitizeString(String(formData.originalSize || ''), 'originalSize');

    const { matchRow: sourceRow } = _findItemRow(sheet, originalName, originalSize);
    if (sourceRow === -1) {
      throw new Error(
        `Source item "${originalName}" (size: "${originalSize}") not found — ` +
        `it may have already been modified. Please refresh and try again.`
      );
    }

    const { matchRow: targetRow } = _findItemRow(sheet, newName, newSize);
    if (targetRow === -1) {
      throw new Error(
        `Target item "${newName}" (size: "${newSize}") no longer exists — please refresh and try again.`
      );
    }
    if (targetRow === sourceRow) {
      throw new Error('Source and target are the same item — nothing to merge.');
    }

    const sourceRowValues = sheet.getRange(sourceRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const targetRowValues = sheet.getRange(targetRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const _rowUnitInfo = rowVals => {
      const baseUnit = String(rowVals[ITEMS_COL.BASE_UNIT - 1] || '').trim() || 'Pcs';
      return {
        baseUnit,
        purchaseUnit: String(rowVals[ITEMS_COL.PURCHASE_UNIT - 1] || '').trim() || baseUnit,
        weightPerBaseUnit: Number(rowVals[ITEMS_COL.WEIGHT_PER_BASE_UNIT - 1]) || 0
      };
    };
    const sourceUnitInfo = _rowUnitInfo(sourceRowValues);
    const targetUnitInfo = _rowUnitInfo(targetRowValues);
    const unitsMap = typeof _getUnitsMap === 'function' ? _getUnitsMap() : {};

    // Edited (source) vendors, validated the same way saveItem validates them,
    // then reconciled from the source item's own Purchase Unit into the
    // target's — a raw rate copied across mismatched units would otherwise
    // be silently misread on the target row (see _convertVendorRateAcrossItems).
    const sourceVendors = vendors
      .map(v => {
        const rate = _validateVendorRate(v.rate || 0);
        let convertedRate = rate;
        try {
          convertedRate = _convertVendorRateAcrossItems(rate, sourceUnitInfo, targetUnitInfo, unitsMap);
        } catch (e) {
          // Unconvertible (e.g. no Weight-per-Base-Unit set, or module_units.js
          // not loaded in this context) — fall back to the as-entered rate
          // rather than blocking the merge.
        }
        return {
          vendor: sanitizeString(String(v.vendor || ''), 'vendor.name'),
          rate: convertedRate
        };
      })
      .filter(v => v.vendor && v.rate >= MIN_VENDOR_RATE);

    const targetVendors = _parseVendorPairsFromRow(targetRowValues);

    // Union by vendor name; target's existing rate wins on collision.
    const mergedVendorsByName = new Map();
    targetVendors.forEach(v => mergedVendorsByName.set(v.vendor.toLowerCase(), v));
    sourceVendors.forEach(v => {
      const key = v.vendor.toLowerCase();
      if (!mergedVendorsByName.has(key)) {
        mergedVendorsByName.set(key, v);
      }
    });
    const mergedVendors = Array.from(mergedVendorsByName.values());

    const currentLastCol = sheet.getLastColumn();
    if (currentLastCol >= ITEMS_COL.VENDOR_DATA) {
      sheet.getRange(
        targetRow,
        ITEMS_COL.VENDOR_DATA,
        1,
        currentLastCol - ITEMS_COL.VENDOR_DATA + 1
      ).clearContent();
    }

    if (mergedVendors.length > 0) {
      const vendorRow = [];
      mergedVendors.forEach(v => vendorRow.push(v.vendor, v.rate));
      sheet.getRange(targetRow, ITEMS_COL.VENDOR_DATA, 1, vendorRow.length).setValues([vendorRow]);
    }

    // Delete the source row. Target row was already written above, so a
    // shifted row index from this deletion (when sourceRow < targetRow)
    // doesn't affect anything below.
    sheet.deleteRow(sourceRow);

    SpreadsheetApp.flush();
    invalidateListCache(MASTER_DATA_CACHE_KEYS.ITEMS);

    // Combine stock figures and drop the source's Stock row.
    syncStockForItem('merge', {
      oldName: originalName,
      oldSize: originalSize,
      newName: newName,
      newSize: newSize
    });

    // Repoint historical Bill/PO/BOM rows from the source's old identity
    // to the target's.
    _propagateItemIdentityChange(originalName, originalSize, newName, newSize);

    const msg = `Merged "${originalName}" (size: "${originalSize}") into "${newName}" (size: "${newSize}").`;
    logAction(
      'MERGE',
      APP_CONFIG.SHEETS.ITEMS,
      `${originalName}|${originalSize} -> ${newName}|${newSize}`,
      msg,
      'SUCCESS'
    );

    return buildResponse(true, { name: newName, size: newSize }, msg);
  } catch (error) {
    Log.error('[mergeItemEdit] Error:', error.message);
    logAction('ERROR', 'mergeItemEdit', formData.itemName, error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to merge items: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────────────────
// autoMergeDuplicateItems
// ─────────────────────────────────────────────────────────────────────────

/**
 * Separator used when concatenating Remarks/Narration/Specification text
 * from rows collapsed by autoMergeDuplicateItems().
 */
const AUTO_MERGE_METADATA_SEPARATOR = ' | ';

/**
 * Appends `addition` onto `base` with AUTO_MERGE_METADATA_SEPARATOR,
 * skipping blanks and exact duplicate text (case-insensitive) so re-running
 * the merge doesn't keep growing the same string.
 *
 * @private
 */
function _concatMetadataField(base, addition) {
  const baseTrim = String(base || '').trim();
  const addTrim = String(addition || '').trim();
  if (!addTrim) return baseTrim;
  if (!baseTrim) return addTrim;

  const parts = baseTrim.split(AUTO_MERGE_METADATA_SEPARATOR).map(p => p.trim().toLowerCase());
  if (parts.includes(addTrim.toLowerCase())) return baseTrim;
  return baseTrim + AUTO_MERGE_METADATA_SEPARATOR + addTrim;
}

/**
 * autoMergeDuplicateItems()
 *
 * Scans the whole Items Master sheet for rows sharing the same Name + Size
 * (case-insensitive) — the system's sole identity key, see _findItemRow —
 * and collapses each group down to one surviving row.
 *
 * Unlike mergeItemEdit() (which merges a row being renamed into a
 * *different* name+size it collided with, requiring Stock/PO/Bill/BOM
 * backfill), every row in a group here already shares the same identity, so
 * those other sheets already point at all of them correctly — only the
 * Items Master rows themselves need collapsing. No stock figures are
 * touched and no identity backfill runs.
 *
 * For each duplicate group:
 * - Survivor = the lowest-numbered row.
 * - Remarks/Narration/Specification: concatenated onto the survivor with
 *   " | ", skipping blanks and exact duplicate text.
 * - Vendors: unioned by name; survivor's rate wins on a collision.
 * - The extra row(s) are deleted.
 *
 * All deletions/updates run in a single pass ordered by descending row
 * number, so deleting a row never shifts a not-yet-processed row's index.
 *
 * Intended to run unattended (see installAutoMergeDuplicatesTrigger()), so
 * it never blocks on a confirmation the way saveItem()/mergeItemEdit() do
 * for user-driven edits.
 *
 * @returns {Object} API response with { groupsMerged, rowsRemoved }
 */
function autoMergeDuplicateItems() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(ITEMS_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
    if (!sheet) throw new Error(`Sheet "${APP_CONFIG.SHEETS.ITEMS}" not found.`);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return buildResponse(true, { groupsMerged: 0, rowsRemoved: 0 }, 'Item Master is empty. Nothing to merge.');
    }

    const lastCol = Math.max(sheet.getLastColumn(), ITEMS_COL.VENDOR_DATA);
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // Group sheet rows (1-based) by normalized name|size key, in row order.
    const groups = new Map();
    for (let i = 0; i < data.length; i++) {
      const name = String(data[i][ITEMS_COL.ITEM_NAME - 1] || '').trim();
      if (!name) continue;
      const size = String(data[i][ITEMS_COL.SIZE - 1] || '').trim();
      const key = name.toLowerCase() + '|' + size.toLowerCase();
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(i + 2);
    }

    const duplicateGroups = Array.from(groups.values()).filter(rows => rows.length > 1);

    if (duplicateGroups.length === 0) {
      return buildResponse(true, { groupsMerged: 0, rowsRemoved: 0 }, 'No duplicate Name+Size rows found.');
    }

    // Compute every group's merged values from the in-memory snapshot above
    // (no sheet reads inside the loop), then queue mutations to run in one
    // descending-row-number pass.
    const updates = [];
    const deletions = [];

    duplicateGroups.forEach(rows => {
      const survivorRow = Math.min(...rows);
      const duplicateRows = rows.filter(r => r !== survivorRow);

      const survivorVals = data[survivorRow - 2];
      let remarks = String(survivorVals[ITEMS_COL.REMARKS - 1] || '');
      let narration = String(survivorVals[ITEMS_COL.NARRATION - 1] || '');
      let spec = String(survivorVals[ITEMS_COL.SPECIFICATION - 1] || '');
      const vendorMap = new Map();
      _parseVendorPairsFromRow(survivorVals).forEach(v => vendorMap.set(v.vendor.toLowerCase(), v));

      duplicateRows.forEach(dupRow => {
        const dupVals = data[dupRow - 2];
        remarks = _concatMetadataField(remarks, dupVals[ITEMS_COL.REMARKS - 1]);
        narration = _concatMetadataField(narration, dupVals[ITEMS_COL.NARRATION - 1]);
        spec = _concatMetadataField(spec, dupVals[ITEMS_COL.SPECIFICATION - 1]);
        _parseVendorPairsFromRow(dupVals).forEach(v => {
          const k = v.vendor.toLowerCase();
          if (!vendorMap.has(k)) vendorMap.set(k, v);
        });
        deletions.push(dupRow);
      });

      updates.push({ row: survivorRow, remarks, narration, spec, vendors: Array.from(vendorMap.values()) });
    });

    const ops = [
      ...updates.map(u => ({ type: 'update', row: u.row, payload: u })),
      ...deletions.map(row => ({ type: 'delete', row }))
    ].sort((a, b) => b.row - a.row);

    ops.forEach(op => {
      if (op.type === 'delete') {
        sheet.deleteRow(op.row);
        return;
      }

      const { row, remarks, narration, spec, vendors } = op.payload;
      sheet.getRange(row, ITEMS_COL.REMARKS, 1, 3).setValues([[remarks, narration, spec]]);

      sheet.getRange(row, ITEMS_COL.VENDOR_DATA, 1, lastCol - ITEMS_COL.VENDOR_DATA + 1).clearContent();
      if (vendors.length > 0) {
        const vendorRow = [];
        vendors.forEach(v => vendorRow.push(v.vendor, v.rate));
        sheet.getRange(row, ITEMS_COL.VENDOR_DATA, 1, vendorRow.length).setValues([vendorRow]);
      }
    });

    SpreadsheetApp.flush();
    invalidateListCache(MASTER_DATA_CACHE_KEYS.ITEMS);

    const groupsMerged = duplicateGroups.length;
    const rowsRemoved = deletions.length;
    const msg = `Auto-merged ${groupsMerged} duplicate item group(s), removing ${rowsRemoved} row(s).`;
    logAction('MERGE', APP_CONFIG.SHEETS.ITEMS, 'AUTO_MERGE_DUPLICATES', msg, 'SUCCESS');
    return buildResponse(true, { groupsMerged, rowsRemoved }, msg);
  } catch (error) {
    Log.error('[autoMergeDuplicateItems] Error:', error.message);
    logAction('ERROR', 'autoMergeDuplicateItems', 'AUTO_MERGE_DUPLICATES', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to auto-merge duplicate items: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────────────────
// autoFixTruncatedDuplicateItems
// ─────────────────────────────────────────────────────────────────────────

/**
 * Minimum length (chars) a truncated name candidate must have before
 * autoFixTruncatedDuplicateItems() will act on it. Guards against
 * coincidental short prefixes (e.g. "BIG-BAG" legitimately prefixing many
 * unrelated item names) being mistaken for a cut-off paste/typo.
 */
const MIN_TRUNCATED_NAME_LENGTH = 10;

/**
 * Merges the item identified by (sourceName, sourceSize) into the item
 * identified by (targetName, targetSize), regardless of how similar or
 * different their Name text is. Used both by autoFixTruncatedDuplicateItems
 * (where the pairing was detected automatically) and mergeSelectedItems
 * (where a human picked the pairing by hand because the names don't match
 * any automatic rule — e.g. two rows that encode the same physical item
 * with different/inconsistent dimension text).
 *
 * Concatenates Remarks/Narration/Specification onto the target (rather
 * than discarding the source's, the way mergeItemEdit's user-driven flow
 * does), unions vendors by name (target's rate wins on collision), combines
 * Stock figures, deletes the source row, and backfills Bill/PO/BOM
 * references from the source's identity to the target's.
 *
 * Must be called with the Items lock already held by the caller.
 *
 * @returns {boolean} true if the merge ran, false if either row could no
 *   longer be found (already merged/edited away) or source === target.
 * @private
 */
function _mergeItemIdentities(sheet, sourceName, sourceSize, targetName, targetSize) {
  const { matchRow: sourceRow } = _findItemRow(sheet, sourceName, sourceSize);
  const { matchRow: targetRow } = _findItemRow(sheet, targetName, targetSize);
  if (sourceRow === -1 || targetRow === -1 || sourceRow === targetRow) return false;

  const lastCol = Math.max(sheet.getLastColumn(), ITEMS_COL.VENDOR_DATA);
  const sourceVals = sheet.getRange(sourceRow, 1, 1, lastCol).getValues()[0];
  const targetVals = sheet.getRange(targetRow, 1, 1, lastCol).getValues()[0];

  const remarks = _concatMetadataField(targetVals[ITEMS_COL.REMARKS - 1], sourceVals[ITEMS_COL.REMARKS - 1]);
  const narration = _concatMetadataField(targetVals[ITEMS_COL.NARRATION - 1], sourceVals[ITEMS_COL.NARRATION - 1]);
  const spec = _concatMetadataField(targetVals[ITEMS_COL.SPECIFICATION - 1], sourceVals[ITEMS_COL.SPECIFICATION - 1]);

  // Reconcile source vendor rates into the target's Purchase Unit before
  // merging — see _convertVendorRateAcrossItems (mergeItemEdit's identical
  // handling). Falls back to the raw rate if the units can't be reconciled.
  const _rowUnitInfo = rowVals => {
    const baseUnit = String(rowVals[ITEMS_COL.BASE_UNIT - 1] || '').trim() || 'Pcs';
    return {
      baseUnit,
      purchaseUnit: String(rowVals[ITEMS_COL.PURCHASE_UNIT - 1] || '').trim() || baseUnit,
      weightPerBaseUnit: Number(rowVals[ITEMS_COL.WEIGHT_PER_BASE_UNIT - 1]) || 0
    };
  };
  const sourceUnitInfo = _rowUnitInfo(sourceVals);
  const targetUnitInfo = _rowUnitInfo(targetVals);
  const unitsMap = typeof _getUnitsMap === 'function' ? _getUnitsMap() : {};

  const vendorMap = new Map();
  _parseVendorPairsFromRow(targetVals).forEach(v => vendorMap.set(v.vendor.toLowerCase(), v));
  _parseVendorPairsFromRow(sourceVals).forEach(v => {
    const k = v.vendor.toLowerCase();
    if (!vendorMap.has(k)) {
      let convertedRate = v.rate;
      try {
        convertedRate = _convertVendorRateAcrossItems(v.rate, sourceUnitInfo, targetUnitInfo, unitsMap);
      } catch (e) {
        // Unconvertible — fall back to the as-entered rate rather than
        // blocking the merge.
      }
      vendorMap.set(k, { vendor: v.vendor, rate: convertedRate });
    }
  });

  sheet.getRange(targetRow, ITEMS_COL.REMARKS, 1, 3).setValues([[remarks, narration, spec]]);
  sheet.getRange(targetRow, ITEMS_COL.VENDOR_DATA, 1, lastCol - ITEMS_COL.VENDOR_DATA + 1).clearContent();
  const mergedVendors = Array.from(vendorMap.values());
  if (mergedVendors.length > 0) {
    const vendorRow = [];
    mergedVendors.forEach(v => vendorRow.push(v.vendor, v.rate));
    sheet.getRange(targetRow, ITEMS_COL.VENDOR_DATA, 1, vendorRow.length).setValues([vendorRow]);
  }

  sheet.deleteRow(sourceRow);
  SpreadsheetApp.flush();
  invalidateListCache(MASTER_DATA_CACHE_KEYS.ITEMS);

  syncStockForItem('merge', { oldName: sourceName, oldSize: sourceSize, newName: targetName, newSize: targetSize });
  _propagateItemIdentityChange(sourceName, sourceSize, targetName, targetSize);

  return true;
}

/**
 * mergeSelectedItems(items)
 *
 * Manual override for duplicates the automatic passes (autoMergeDuplicateItems,
 * autoFixTruncatedDuplicateItems) won't touch by design — e.g. two rows a
 * human can tell are the same physical item, but whose Name text doesn't
 * match exactly and isn't a clean prefix of one another (different/typo'd
 * dimension codes, etc). Driven by checkbox selection in the Item Master
 * table: the user checks the item to KEEP first, then the duplicate to
 * remove, and the UI sends them in that order.
 *
 * @param {Array<{name: string, size: string}>} items - Exactly 2 items;
 *   items[0] is the survivor (target), items[1] is merged in and deleted
 *   (source).
 * @returns {Object} API response
 */
function mergeSelectedItems(items) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(ITEMS_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    if (!Array.isArray(items) || items.length !== 2) {
      throw new Error('Select exactly 2 items to merge.');
    }

    const sheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
    if (!sheet) throw new Error(`Sheet "${APP_CONFIG.SHEETS.ITEMS}" not found.`);

    const target = items[0];
    const source = items[1];
    const targetName = _validateItemName(target.name);
    const targetSize = sanitizeString(target.size || '', 'itemSize');
    const sourceName = _validateItemName(source.name);
    const sourceSize = sanitizeString(source.size || '', 'itemSize');

    const ok = _mergeItemIdentities(sheet, sourceName, sourceSize, targetName, targetSize);
    if (!ok) {
      throw new Error('One of the selected items could not be found — it may have already been edited or merged. Please refresh and try again.');
    }

    const msg = `Merged "${sourceName}" (size: "${sourceSize}") into "${targetName}" (size: "${targetSize}").`;
    logAction('MERGE', APP_CONFIG.SHEETS.ITEMS, `${sourceName}|${sourceSize} -> ${targetName}|${targetSize}`, msg, 'SUCCESS');
    return buildResponse(true, { from: sourceName, to: targetName, size: targetSize }, msg);
  } catch (error) {
    Log.error('[mergeSelectedItems] Error:', error.message);
    logAction('ERROR', 'mergeSelectedItems', JSON.stringify(items), error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to merge selected items: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * autoFixTruncatedDuplicateItems()
 *
 * Catches a narrower, different failure mode than autoMergeDuplicateItems():
 * two rows with the SAME Size whose Names are not equal, but one Name is an
 * exact, case-insensitive prefix of the other — the signature of a name
 * that got cut off mid-paste/mid-type (e.g. "...-(50*7" vs the full
 * "...-(50*71)-(3"). Because the Name actually differs, this is a real
 * identity change, not just a cosmetic dup — so unlike
 * autoMergeDuplicateItems, this DOES rename + merge stock + backfill
 * Bill/PO/BOM references (the same machinery saveItem's rename path and
 * mergeItemEdit use).
 *
 * Deliberately conservative:
 * - Only acts when the short name is >= MIN_TRUNCATED_NAME_LENGTH chars.
 * - Only acts when the short name is a prefix of EXACTLY ONE other name in
 *   its Size group. If it's a prefix of more than one (ambiguous — could
 *   be two genuinely different fuller item names), it's skipped and left
 *   for manual review rather than guessed at.
 *
 * Each fix re-resolves rows from the live sheet (via _findItemRow) rather
 * than working off a single stale snapshot, since each fix changes row
 * numbers and Stock/Bill/PO/BOM state that the next fix's lookups depend on.
 *
 * @returns {Object} API response with { fixed: [{ from, to, size }], skippedAmbiguous }
 */
function autoFixTruncatedDuplicateItems() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(ITEMS_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
    if (!sheet) throw new Error(`Sheet "${APP_CONFIG.SHEETS.ITEMS}" not found.`);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return buildResponse(true, { fixed: [], skippedAmbiguous: 0 }, 'Item Master is empty. Nothing to fix.');
    }

    const data = sheet.getRange(2, ITEMS_COL.ITEM_NAME, lastRow - 1, ITEMS_COL.SIZE - ITEMS_COL.ITEM_NAME + 1).getValues();

    // Group names (with original casing preserved) by normalized Size.
    const bySize = new Map();
    data.forEach(row => {
      const name = String(row[ITEMS_COL.ITEM_NAME - 1] || '').trim();
      if (!name) return;
      const size = String(row[ITEMS_COL.SIZE - 1] || '').trim();
      const sizeKey = size.toLowerCase();
      if (!bySize.has(sizeKey)) bySize.set(sizeKey, { size, names: new Set() });
      bySize.get(sizeKey).names.add(name);
    });

    // Detect unambiguous prefix pairs: { shortName, longName, size }.
    const plan = [];
    let skippedAmbiguous = 0;

    bySize.forEach(({ size, names }) => {
      const list = Array.from(names);
      // Lowercased ONCE per name here, reused for every comparison below —
      // this runs on an hourly trigger over the whole Item Master, and the
      // candidate x other comparison is inherently O(k^2) per Size group,
      // but re-deriving other.toLowerCase() from scratch inside the inner
      // loop meant doing that same string transform up to k times per name
      // (O(k^2) toLowerCase calls) instead of once each (O(k)).
      const lowerList = list.map(n => n.toLowerCase());

      list.forEach((candidate, i) => {
        if (candidate.length < MIN_TRUNCATED_NAME_LENGTH) return;
        const candidateLower = lowerList[i];

        const matches = [];
        list.forEach((other, j) => {
          if (j === i) return;
          const otherLower = lowerList[j];
          if (otherLower.length > candidateLower.length && otherLower.startsWith(candidateLower)) {
            matches.push(other);
          }
        });

        if (matches.length === 1) {
          plan.push({ shortName: candidate, longName: matches[0], size });
        } else if (matches.length > 1) {
          skippedAmbiguous++;
        }
      });
    });

    const fixed = [];

    plan.forEach(({ shortName, longName, size }) => {
      const ok = _mergeItemIdentities(sheet, shortName, size, longName, size);
      if (ok) fixed.push({ from: shortName, to: longName, size });
    });

    const msg = fixed.length > 0
      ? `Auto-fixed ${fixed.length} truncated duplicate name(s). ${skippedAmbiguous} ambiguous case(s) left for manual review.`
      : `No truncated duplicate names found. ${skippedAmbiguous} ambiguous case(s) left for manual review.`;
    logAction('MERGE', APP_CONFIG.SHEETS.ITEMS, 'AUTO_FIX_TRUNCATED_NAMES', msg + ' ' + JSON.stringify(fixed), 'SUCCESS');
    return buildResponse(true, { fixed, skippedAmbiguous }, msg);
  } catch (error) {
    Log.error('[autoFixTruncatedDuplicateItems] Error:', error.message);
    logAction('ERROR', 'autoFixTruncatedDuplicateItems', 'AUTO_FIX_TRUNCATED_NAMES', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to auto-fix truncated duplicate items: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Runs both unattended item-cleanup passes in the right order: fix
 * truncated-name duplicates first (since fixing a truncation can itself
 * produce an exact Name+Size match), then collapse any remaining exact
 * Name+Size duplicates. Bound to the scheduled trigger — see
 * installAutoMergeDuplicatesTrigger().
 *
 * @returns {Object} API response combining both passes' data
 */
function runScheduledItemCleanup() {
  const fixResult = autoFixTruncatedDuplicateItems();
  const mergeResult = autoMergeDuplicateItems();
  return buildResponse(
    fixResult.success && mergeResult.success,
    { fix: fixResult.data, merge: mergeResult.data },
    `${fixResult.message} ${mergeResult.message}`
  );
}

/**
 * One-time setup: installs an hourly time-driven trigger that runs
 * runScheduledItemCleanup() (truncated-name fix + exact-duplicate merge)
 * unattended. Run manually once from the Apps Script editor (mirrors
 * initItemsSheet()'s "call once" pattern). Safe to
 * call again — clears any existing trigger for the current handler AND
 * the handler this function used to install ('autoMergeDuplicateItems',
 * before runScheduledItemCleanup wrapped it) before creating a fresh one,
 * so an earlier install can't leave two triggers fighting over the same
 * document lock every hour.
 */
function installAutoMergeDuplicatesTrigger() {
  const staleHandlers = ['runScheduledItemCleanup', 'autoMergeDuplicateItems'];
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (staleHandlers.includes(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('runScheduledItemCleanup')
    .timeBased()
    .everyHours(1)
    .create();
}

// ─────────────────────────────────────────────────────────────────────────
// SHEET INITIALIZATION UTILITY
// ─────────────────────────────────────────────────────────────────────────

/**
 * initItemsSheet(vendorSlots)
 * 
 * Initializes the Items sheet with the correct header row
 * 
 * Call Once:
 * - Called manually from Apps Script editor at setup time
 * - NOT called by saveItem (safe for repeated calls)
 * - Will not duplicate headers if already present
 * - Overwrites existing row 1 with new headers
 * 
 * Operations:
 * 1. Fetch Items sheet
 * 2. Write header row with fixed columns + vendor pairs
 * 3. Format headers (bold, background color)
 * 4. Flush changes
 * 5. Log success message
 * 
 * @param {number} [vendorSlots=10] - Number of vendor pair columns to pre-create
 *   - Each slot takes 2 columns (Vendor Name | Vendor Rate)
 *   - Default 10 slots = 20 columns for vendors (beyond the 5 fixed columns)
 * 
 * @throws {Error} If Items sheet not found
 * 
 * @example
 * // Call once after deploying the GAS project
 * function setup() {
 *   initItemsSheet(10);  // Creates 10 vendor pair slots
 * }
 */
function initItemsSheet(vendorSlots = DEFAULT_VENDOR_SLOTS) {
  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
    if (!sheet) {
      throw new Error(`Sheet "${APP_CONFIG.SHEETS.ITEMS}" not found.`);
    }

    _writeItemHeaders(sheet, vendorSlots);
    SpreadsheetApp.flush();

    Logger.log(
      `[initItemsSheet] Items sheet initialized with ${vendorSlots} vendor slots.`
    );

    return buildResponse(
      true,
      { vendorSlots },
      `Items sheet initialized successfully.`
    );
  } catch (error) {
    Log.error('[initItemsSheet] Error:', error.message);
    return buildResponse(
      false,
      null,
      'Failed to initialize Items sheet: ' + error.message
    );
  }
}

/**
 * migrateItemsMasterUnitColumns()
 *
 * One-time migration for spreadsheets created before Base Unit / Purchase
 * Unit / Weight per Base Unit existed. Inserts 3 columns before the current
 * Vendor Data block, writes the 3 new headers, and defaults every existing
 * row's Base Unit and Purchase Unit to 'Pcs' — matching the system's prior
 * implicit assumption, so no historical rate/stock numbers change.
 *
 * Safe to call again: detects the new headers are already present (column F
 * header is 'Base Unit') and does nothing.
 *
 * Run manually once from the Apps Script editor after deploying this change.
 *
 * @returns {Object} API response
 */
function migrateItemsMasterUnitColumns() {
  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
    if (!sheet) {
      throw new Error(`Sheet "${APP_CONFIG.SHEETS.ITEMS}" not found.`);
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastCol >= 6) {
      const existingHeader = String(sheet.getRange(1, 6).getValue() || '').trim();
      if (existingHeader === 'Base Unit') {
        return buildResponse(true, { migrated: false }, 'Items Master already has unit columns. Nothing to migrate.');
      }
    }

    // Shift everything from the old Vendor Data start (col 6) rightward by 3.
    sheet.insertColumnsBefore(6, 3);

    sheet.getRange(1, 6, 1, 3)
      .setValues([['Base Unit', 'Purchase Unit', 'Weight per Base Unit (g)']])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    if (lastRow >= 2) {
      const defaults = Array.from({ length: lastRow - 1 }, () => ['Pcs', 'Pcs', '']);
      sheet.getRange(2, 6, defaults.length, 3).setValues(defaults);
    }

    SpreadsheetApp.flush();

    const msg = `Migrated Items Master: inserted unit columns, defaulted ${Math.max(lastRow - 1, 0)} row(s) to Base/Purchase Unit 'Pcs'.`;
    logAction('UPDATE', APP_CONFIG.SHEETS.ITEMS, 'MIGRATE_UNIT_COLUMNS', msg, 'SUCCESS');
    return buildResponse(true, { migrated: true, rowsUpdated: Math.max(lastRow - 1, 0) }, msg);
  } catch (error) {
    Log.error('[migrateItemsMasterUnitColumns] Error:', error.message);
    logAction('ERROR', 'migrateItemsMasterUnitColumns', 'MIGRATE_UNIT_COLUMNS', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to migrate Items Master unit columns: ' + error.message);
  }
}

/**
 * Test function to verify Items sheet is properly initialized
 * Call this from the Apps Script editor to diagnose sheet issues
 * 
 * @returns {Object} Diagnostic information
 */
function testItemsSheet() {
  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
    if (!sheet) {
      return { success: false, message: 'Items sheet not found' };
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const headerRow = lastRow > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];

    return {
      success: true,
      message: 'Items sheet exists and is accessible',
      diagnostics: {
        lastRow,
        lastCol,
        headerCount: headerRow.length,
        headers: headerRow.slice(0, 10),  // Show first 10 headers
        isInitialized: lastRow > 0 && headerRow[0] === 'Item Name'
      }
    };
  } catch (error) {
    return {
      success: false,
      message: 'Failed to test Items sheet: ' + error.message
    };
  }
}

/**
 * Imports items from the Stock sheet into Items Master.
 * For each Stock row whose (name, size) pair does not already appear in Items Master
 * (under any narration), a new Item Master row is created with blank metadata and
 * no vendor rates. The user can then enrich those records via the Edit flow.
 *
 * Runs under a document lock to prevent concurrent mutations.
 *
 * @returns {Object} API response
 *   - data: { added, skipped }
 *   - message: Human-readable summary
 */
function importItemsFromStock() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(ITEMS_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const itemsSheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
    if (!itemsSheet) throw new Error(`Sheet "${APP_CONFIG.SHEETS.ITEMS}" not found.`);

    let stockSheet;
    try {
      stockSheet = getSheet(APP_CONFIG.SHEETS.STOCK);
    } catch (e) {
      return buildResponse(false, null, 'Stock sheet not found. Nothing to import.');
    }

    const stockLastRow = stockSheet.getLastRow();
    if (stockLastRow < 2) {
      return buildResponse(true, { added: 0, skipped: 0 }, 'Stock sheet is empty. Nothing to import.');
    }

    // Build a set of all (name|size) pairs already in Items Master (any narration)
    const existingKeys = new Set();
    const itemsLastRow = itemsSheet.getLastRow();
    if (itemsLastRow >= 2) {
      const itemsData = itemsSheet
        .getRange(2, ITEMS_COL.ITEM_NAME, itemsLastRow - 1, 2)
        .getValues();
      for (let i = 0; i < itemsData.length; i++) {
        const n = String(itemsData[i][0] || '').trim().toLowerCase();
        const s = String(itemsData[i][1] || '').trim().toLowerCase();
        if (n) existingKeys.add(n + '|' + s);
      }
    }

    // Read Stock rows: only name (col 1) and size (col 2)
    const stockData = stockSheet
      .getRange(2, STOCK_COL.ITEM_NAME, stockLastRow - 1, 2)
      .getValues();

    const newRows = [];
    let skipped = 0;

    for (let i = 0; i < stockData.length; i++) {
      const name = String(stockData[i][0] || '').trim();
      const size = String(stockData[i][1] || '').trim();
      if (!name) continue;

      const key = name.toLowerCase() + '|' + size.toLowerCase();
      if (existingKeys.has(key)) {
        skipped++;
        continue;
      }

      // [Item Name, Size, Remarks, Narration, Specification] — vendor columns left empty
      newRows.push([name, size, '', '', '']);
      existingKeys.add(key); // prevent duplicates within the same import batch
    }

    if (newRows.length > 0) {
      const appendStart = itemsSheet.getLastRow() + 1;
      itemsSheet.getRange(appendStart, 1, newRows.length, 5).setValues(newRows);
      SpreadsheetApp.flush();
      invalidateListCache(MASTER_DATA_CACHE_KEYS.ITEMS);
    }

    const added = newRows.length;
    const msg = added > 0
      ? `Imported ${added} item(s) from Stock. ${skipped} already existed and were skipped.`
      : `All ${skipped} Stock item(s) already exist in Item Master. Nothing new to import.`;

    logAction('CREATE', APP_CONFIG.SHEETS.ITEMS, 'SYNC_FROM_STOCK', msg, 'SUCCESS');
    return buildResponse(true, { added, skipped }, msg);
  } catch (error) {
    Log.error('[importItemsFromStock] Error:', error.message);
    logAction('ERROR', 'importItemsFromStock', 'SYNC_FROM_STOCK', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to import from Stock: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Called from the onEdit trigger when a single Item Name/Size cell is
 * changed directly on the Stock sheet. Repoints the matching Item Master
 * row (and its PO/Bill/BOM history) to the new identity in place, so a
 * manual rename in Stock doesn't get mistaken for a delete-and-add — which
 * would otherwise leave the old name behind as an orphaned Item Master row
 * and create a second, blank-metadata row for the new name.
 *
 * No-ops (silently) if there's nothing to rename or if the new identity
 * already exists elsewhere in Item Master — in the latter case the old row
 * is left as-is and will surface via the orphan-review flow instead.
 *
 * Best-effort: errors are logged, not thrown, since this runs from a simple
 * trigger with no UI to report failures to.
 *
 * @param {string} oldName
 * @param {string} oldSize
 * @param {string} newName
 * @param {string} newSize
 */
function renameItemMasterForStockEdit(oldName, oldSize, newName, newSize) {
  if (oldName === newName && oldSize === newSize) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(ITEMS_LOCK_TIMEOUT_MS)) return;

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
    if (!sheet) return;

    const { matchRow } = _findItemRow(sheet, oldName, oldSize);
    if (matchRow === -1) return;

    const { matchRow: collisionRow } = _findItemRow(sheet, newName, newSize);
    if (collisionRow !== -1) return;

    sheet.getRange(matchRow, ITEMS_COL.ITEM_NAME, 1, 2).setValues([[newName, newSize]]);
    SpreadsheetApp.flush();
    invalidateListCache(MASTER_DATA_CACHE_KEYS.ITEMS);

    _propagateItemIdentityChange(oldName, oldSize, newName, newSize);

    logAction(
      'UPDATE',
      APP_CONFIG.SHEETS.ITEMS,
      newName,
      `Renamed from "${oldName}|${oldSize}" via direct edit on Stock sheet`,
      'SUCCESS'
    );
  } catch (error) {
    Log.error('[renameItemMasterForStockEdit] Error:', error.message);
    logAction('ERROR', 'renameItemMasterForStockEdit', `${oldName}|${oldSize} -> ${newName}|${newSize}`, error.message, 'ERROR');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Resolves an Item Master row that has no matching Stock row (surfaced to
 * the user via the Item Master "Sync Review" panel) by creating the missing
 * Stock row with the given initial stock quantity.
 *
 * @param {string} name
 * @param {string} size
 * @param {number} initialStock
 * @returns {Object} API response
 */
function keepOrphanItem(name, size, initialStock) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(ITEMS_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const validName = _validateItemName(name);
    const validSize = sanitizeString(size || '', 'size');
    const qty = Number(initialStock) || 0;

    syncStockForItem('ensure', { name: validName, size: validSize, initialStock: qty });

    const msg = `Stock row created for "${validName}" (size: "${validSize}") with initial stock ${qty}.`;
    logAction('CREATE', APP_CONFIG.SHEETS.STOCK, validName, msg, 'SUCCESS');
    return buildResponse(true, { name: validName, size: validSize, initialStock: qty }, msg);
  } catch (error) {
    Log.error('[keepOrphanItem] Error:', error.message);
    logAction('ERROR', 'keepOrphanItem', name, error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to create Stock row: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Bulk variant of keepOrphanItem() — creates the missing Stock row for each
 * given Item Master entry under a single document lock, used by the Sync
 * Review modal's "Keep Selected" bulk action.
 *
 * @param {Array<{name: string, size: string, initialStock: number}>} items
 * @returns {Object} API response with { created } count
 */
function keepOrphanItemsBulk(items) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(ITEMS_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const toEnsure = (items || []).map(item => ({
      name: _validateItemName(item.name),
      size: sanitizeString(item.size || '', 'size'),
      initialStock: Number(item.initialStock) || 0
    }));
    // Same "processed" count as before (not "actually inserted") — this bulk
    // action is only ever invoked on items already known to be missing a
    // Stock row (the Sync Review modal's orphan list), so every item here is
    // expected to insert; see _ensureStockRowsBulk for the batched write.
    _ensureStockRowsBulk(toEnsure);
    const created = toEnsure.length;

    const msg = `Stock row created for ${created} item(s).`;
    logAction('CREATE', APP_CONFIG.SHEETS.STOCK, 'BULK', msg, 'SUCCESS');
    return buildResponse(true, { created }, msg);
  } catch (error) {
    Log.error('[keepOrphanItemsBulk] Error:', error.message);
    logAction('ERROR', 'keepOrphanItemsBulk', '', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to create Stock rows: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Ensures every Item Master row has a matching Stock row (Item Name + Size).
 * Mirrors importItemsFromStock() in the opposite direction — used to catch
 * items that were added directly on the Items Master sheet (bypassing
 * saveItem) so Stock stays automatically in sync either way.
 *
 * Idempotent: relies on syncStockForItem('ensure', ...), which only appends
 * a Stock row when one doesn't already exist for that name + size.
 *
 * @returns {Object} API response with { checked } count
 */
function syncAllItemsToStock() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(ITEMS_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const itemsSheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
    if (!itemsSheet) throw new Error(`Sheet "${APP_CONFIG.SHEETS.ITEMS}" not found.`);

    const lastRow = itemsSheet.getLastRow();
    if (lastRow < 2) {
      return buildResponse(true, { checked: 0 }, 'Item Master is empty. Nothing to sync.');
    }

    const data = itemsSheet.getRange(2, ITEMS_COL.ITEM_NAME, lastRow - 1, 2).getValues();
    const seen = new Set();
    const toEnsure = [];
    let checked = 0;

    for (let i = 0; i < data.length; i++) {
      const name = String(data[i][0] || '').trim();
      const size = String(data[i][1] || '').trim();
      if (!name) continue;

      const key = name.toLowerCase() + '|' + size.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);

      toEnsure.push({ name, size });
      checked++;
    }

    // One Stock-sheet read + one batched write for every missing row, instead
    // of syncStockForItem('ensure', ...) re-reading the whole Stock sheet
    // once per Item Master row (O(items × stockRows) — this runs over the
    // ENTIRE Item Master every time).
    _ensureStockRowsBulk(toEnsure);
    SpreadsheetApp.flush();
    return buildResponse(true, { checked }, 'Stock sheet synced with Item Master.');
  } catch (error) {
    Log.error('[syncAllItemsToStock] Error:', error.message);
    return buildResponse(false, null, 'Failed to sync stock: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes multiple items in bulk from the master sheet.
 * @param {Array<Object>} items - Array of items to delete, each having name, size
 * @returns {Object} Standardized API response
 */
function deleteItemsBulk(items) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(ITEMS_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
    if (!sheet) {
      throw new Error(`Sheet "${APP_CONFIG.SHEETS.ITEMS}" not found.`);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return buildResponse(true, null, 'No items to delete.');
    }

    // Normalize the requested items, then skip any still referenced by a
    // Product's BOM or an ITEM-sourced Process recipe — deleting them would
    // orphan that reference (see _getItemKeysInUse).
    const requestedItems = (items || [])
      .map(item => ({ name: String(item.name || '').trim(), size: String(item.size || '').trim() }))
      .filter(it => it.name);
    const inUseKeys = _getItemKeysInUse(requestedItems);
    const keyOf = it => it.name.toLowerCase() + '|' + it.size.toLowerCase();
    const skipped = requestedItems.filter(it => inUseKeys.has(keyOf(it)));
    const deletable = requestedItems.filter(it => !inUseKeys.has(keyOf(it)));

    const targetSet = new Set(deletable.map(keyOf));

    const { rowsDeleted, deletedRows } = rewriteSheetExcludingRows(sheet, 2, row => {
      const rowName = String(row[ITEMS_COL.ITEM_NAME - 1] || '').trim().toLowerCase();
      const rowSize = String(row[ITEMS_COL.SIZE - 1] || '').trim().toLowerCase();
      return targetSet.has(`${rowName}|${rowSize}`);
    });

    SpreadsheetApp.flush();
    invalidateListCache(MASTER_DATA_CACHE_KEYS.ITEMS);

    // Remove Stock entries for items no longer present in the master
    // (skipped/in-use items were never deleted from Items Master, so their
    // Stock entry must be left alone). One post-delete read of Items Master
    // builds a Set of remaining (name,size) keys, checked once per deletable
    // item in O(1) — instead of _nameSizeStillExists() re-reading and
    // linearly scanning the whole Items Master once per deletable item.
    const remainingLastRow = sheet.getLastRow();
    const remainingKeys = new Set();
    if (remainingLastRow >= 2) {
      sheet.getRange(2, ITEMS_COL.ITEM_NAME, remainingLastRow - 1, 2).getValues().forEach(row => {
        remainingKeys.add(String(row[0] || '').trim().toLowerCase() + '|' + String(row[1] || '').trim().toLowerCase());
      });
    }

    const stockKeysToCheck = new Map();
    deletable.forEach(({ name, size }) => {
      stockKeysToCheck.set(name.toLowerCase() + '|' + size.toLowerCase(), { name, size });
    });

    stockKeysToCheck.forEach(({ name, size }, key) => {
      if (!remainingKeys.has(key)) {
        syncStockForItem('remove', { name, size });
      }
    });

    const deletedItems = deletedRows.map(row => ({
      name: String(row[ITEMS_COL.ITEM_NAME - 1] || '').trim(),
      size: String(row[ITEMS_COL.SIZE - 1] || '').trim(),
    }));

    let msg = `Successfully deleted ${rowsDeleted} item(s) from master.`;
    if (skipped.length > 0) {
      const labels = skipped.map(it => it.size ? `${it.name} (${it.size})` : it.name);
      msg += ` Skipped ${skipped.length} item(s) referenced by a Product's BOM or a Process recipe: ${labels.join(', ')}.`;
    }
    logAction('BULK_DELETE', APP_CONFIG.SHEETS.ITEMS, 'multiple', msg, 'SUCCESS');
    return buildResponse(true, { deletedItems, skipped }, msg);
  } catch (error) {
    Log.error('[deleteItemsBulk] Error:', error.message);
    logAction('ERROR', 'deleteItemsBulk', 'multiple', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete items: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

