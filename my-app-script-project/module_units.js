/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_units.gs — UNIT MASTER MODULE
 *
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Small, rarely-edited master list of fixed-ratio unit conversions (Pcs,
 * Dozen, Gross, Kg, ...), grouped by FAMILY. Used to convert quantities/rates
 * entered on PO/Bill lines (in whatever unit the vendor quotes) into the
 * item's Base Unit, which is what Stock/BOM/Production actually track.
 *
 * Crossing families (e.g. an item bought by Kg but used by Pcs) is NOT a
 * fixed unit ratio — it needs the item's own Weight per Base Unit
 * (ITEMS_COL.WEIGHT_PER_BASE_UNIT, see module_items.js), so the conversion
 * helpers here take the item as a parameter.
 *
 * CRUD here mirrors module_vendors.js's saveVendor/deleteVendor pattern —
 * same shape as the Vendors master, just a different sheet.
 * ═══════════════════════════════════════════════════════════════════════════
 */

const UNITS_LOCK_TIMEOUT_MS = 15000;

/**
 * Seed rows written by initUnitsSheet() the first time the sheet is created.
 * Pcs / Gram / ML are each their family's base unit (factor 1).
 */
const DEFAULT_UNIT_SEED_ROWS = Object.freeze([
  ['Pcs',   'Count',  1,    'Base unit for count-based items'],
  ['Dozen', 'Count',  12,   ''],
  ['Gross', 'Count',  144,  '12 dozen'],
  ['Box',   'Count',  50,   'Default box size — edit per actual pack size'],
  ['Gram',  'Weight', 1,    'Base unit for weight-based items'],
  ['Kg',    'Weight', 1000, ''],
  ['ML',    'Volume', 1,    'Base unit for volume-based items'],
  ['L',     'Volume', 1000, '']
]);

/**
 * Run once manually from the Apps Script editor to set up the Unit Master
 * sheet headers and seed it with standard units. Safe to call again — only
 * seeds rows when the sheet is otherwise empty.
 */
function initUnitsSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.UNITS);
    if (!sheet) {
      sheet = ss.insertSheet(APP_CONFIG.SHEETS.UNITS);
    }

    sheet.getRange(1, 1, 1, 4)
      .setValues([['Unit Name', 'Family', 'Factor to Base', 'Remarks']])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    if (sheet.getLastRow() < 2) {
      sheet.getRange(2, 1, DEFAULT_UNIT_SEED_ROWS.length, 4)
        .setValues(DEFAULT_UNIT_SEED_ROWS.map(r => r.slice()));
    }

    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Unit Master sheet initialized successfully.');
  } catch (error) {
    console.error('[initUnitsSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize Unit Master sheet: ' + error.message);
  }
}

/**
 * Retrieves all units from the Unit Master sheet.
 * @returns {Object} API response with data: Array<{unitName, family, factorToBase, remarks}>
 */
function getUnitsData() {
  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.UNITS);
    if (!sheet) return buildResponse(false, null, 'Unit Master sheet not found. Please run initUnitsSheet().');

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, []);

    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const units = data
      .map(r => ({
        unitName: String(r[UNITS_COL.UNIT_NAME - 1] || '').trim(),
        family: String(r[UNITS_COL.FAMILY - 1] || '').trim(),
        factorToBase: Number(r[UNITS_COL.FACTOR_TO_BASE - 1]) || 0,
        remarks: String(r[UNITS_COL.REMARKS - 1] || '').trim()
      }))
      .filter(u => u.unitName !== '');

    units.sort((a, b) => a.unitName.localeCompare(b.unitName));

    return buildResponse(true, units);
  } catch (error) {
    console.error('[getUnitsData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load units: ' + error.message);
  }
}

/**
 * Creates or updates a unit conversion entry.
 * @param {Object} formData - { unitName, family, factorToBase, remarks, originalUnitName? }
 * @returns {Object} API response
 */
function saveUnit(formData) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(UNITS_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.UNITS);
    if (!sheet) throw new Error('Unit Master sheet not found.');

    const newName = sanitizeString(formData.unitName || '', 'unitName');
    if (!newName) throw new Error('Unit name must not be empty.');

    const family = sanitizeString(formData.family || '', 'family');
    if (!family) throw new Error('Family must not be empty.');

    const factor = Number(formData.factorToBase);
    if (!Number.isFinite(factor) || factor <= 0) {
      throw new Error('Factor to Base must be a positive number.');
    }

    const remarks = sanitizeString(formData.remarks || '', 'remarks');

    const isEdit = !!formData.originalUnitName;
    const originalName = isEdit ? sanitizeString(formData.originalUnitName, 'originalUnitName') : newName;

    const lastRow = sheet.getLastRow();
    const data = lastRow >= 2 ? sheet.getRange(2, UNITS_COL.UNIT_NAME, lastRow - 1, 1).getValues() : [];

    let targetRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === originalName.toLowerCase()) {
        targetRow = i + 2;
        break;
      }
    }

    if (isEdit && targetRow === -1) throw new Error('Original unit not found in sheet.');
    if (!isEdit && targetRow !== -1) throw new Error(`Unit "${newName}" already exists.`);

    if (isEdit && newName.toLowerCase() !== originalName.toLowerCase()) {
      const dupeRow = data.findIndex(r => String(r[0]).trim().toLowerCase() === newName.toLowerCase());
      if (dupeRow !== -1) throw new Error(`Another unit named "${newName}" already exists.`);
    }

    const rowData = [newName, family, factor, remarks];

    if (isEdit) {
      sheet.getRange(targetRow, 1, 1, 4).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    SpreadsheetApp.flush();

    const msg = isEdit ? 'Unit updated successfully.' : 'Unit added successfully.';
    logAction(isEdit ? 'UPDATE' : 'CREATE', APP_CONFIG.SHEETS.UNITS, newName, msg, 'SUCCESS');

    return buildResponse(true, { unitName: newName }, msg);
  } catch (error) {
    console.error('[saveUnit] Error:', error.message);
    logAction('ERROR', 'saveUnit', formData ? formData.unitName : 'unknown', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to save unit: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes a unit conversion entry.
 * @param {string} unitName
 * @returns {Object} API response
 */
function deleteUnit(unitName) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(UNITS_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.UNITS);
    if (!sheet) throw new Error('Unit Master sheet not found.');

    const targetName = String(unitName || '').trim().toLowerCase();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('Unit not found.');

    const data = sheet.getRange(2, UNITS_COL.UNIT_NAME, lastRow - 1, 1).getValues();
    let targetRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === targetName) {
        targetRow = i + 2;
        break;
      }
    }

    if (targetRow === -1) throw new Error('Unit not found.');

    sheet.deleteRow(targetRow);
    SpreadsheetApp.flush();

    logAction('DELETE', APP_CONFIG.SHEETS.UNITS, unitName, 'Unit deleted', 'SUCCESS');
    return buildResponse(true, null, `Unit "${unitName}" deleted.`);
  } catch (error) {
    console.error('[deleteUnit] Error:', error.message);
    logAction('ERROR', 'deleteUnit', unitName, error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete unit: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

// ─────────────────────────────────────────────────────────────────────────
// CONVERSION HELPERS
// Shared by module_items.js (ratePerBaseUnit), module_po.js / module_bill.js
// (baseQty/baseRate on save), and the client (via google.script.run) for
// live conversion previews on PO/Bill forms.
// ─────────────────────────────────────────────────────────────────────────

/**
 * Cached within a single GAS execution — Unit Master rarely changes mid-call
 * and every PO/Bill line lookup would otherwise re-read the sheet.
 */
let _cachedUnitsMap = null;

/**
 * Reads the Unit Master sheet once and returns a lookup map.
 * @returns {Object} { [unitNameLowercase]: { unitName, family, factorToBase } }
 */
function _getUnitsMap() {
  if (_cachedUnitsMap) return _cachedUnitsMap;

  const map = {};
  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.UNITS);
    const lastRow = sheet ? sheet.getLastRow() : 0;
    if (sheet && lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
      data.forEach(r => {
        const unitName = String(r[UNITS_COL.UNIT_NAME - 1] || '').trim();
        if (!unitName) return;
        map[unitName.toLowerCase()] = {
          unitName,
          family: String(r[UNITS_COL.FAMILY - 1] || '').trim(),
          factorToBase: Number(r[UNITS_COL.FACTOR_TO_BASE - 1]) || 1
        };
      });
    }
  } catch (e) {
    // Unit Master not found/accessible — callers fall back to treating
    // unknown units as factor-1 (today's implicit behavior).
  }

  _cachedUnitsMap = map;
  return map;
}

/**
 * Looks up a unit's {family, factorToBase}, defaulting to factor 1 / the
 * item's own base-unit family for unknown units (so a typo'd unit doesn't
 * hard-fail a save — it's just treated as already being the base unit).
 * @private
 */
function _lookupUnit(unitName, fallbackFamily, unitsMap) {
  const entry = unitsMap[String(unitName || '').trim().toLowerCase()];
  if (entry) return entry;
  return { unitName: unitName, family: fallbackFamily, factorToBase: 1 };
}

/**
 * Converts a quantity entered in `fromUnit` into the item's Base Unit.
 *
 * Same family as the item's Base Unit (e.g. Gross -> Pcs, both 'Count'):
 *   baseQty = qty * factor(fromUnit) / factor(item.baseUnit)
 *
 * Crossing families (e.g. Kg -> Pcs, 'Weight' -> 'Count'): requires
 * item.weightPerBaseUnit (grams per Base Unit) and converts via grams.
 *
 * @param {number} qty
 * @param {string} fromUnit
 * @param {{baseUnit: string, weightPerBaseUnit?: number}} item
 * @param {Object} [unitsMap] - Pass _getUnitsMap() result to avoid re-reading the sheet
 * @returns {number} Quantity in the item's Base Unit
 * @throws {Error} If crossing families without item.weightPerBaseUnit set
 */
function convertQtyToBaseUnit(qty, fromUnit, item, unitsMap) {
  const map = unitsMap || _getUnitsMap();
  const baseUnitName = item.baseUnit || 'Pcs';
  const baseEntry = _lookupUnit(baseUnitName, 'Count', map);
  const fromEntry = _lookupUnit(fromUnit, baseEntry.family, map);

  const q = Number(qty) || 0;

  if (fromEntry.family === baseEntry.family) {
    return (q * fromEntry.factorToBase) / baseEntry.factorToBase;
  }

  // Crossing families — only Weight -> Count (buy by weight, use by count)
  // is supported, via the item's grams-per-Base-Unit.
  const weightPerBaseUnit = Number(item.weightPerBaseUnit) || 0;
  if (weightPerBaseUnit <= 0) {
    throw new Error(
      `Cannot convert "${fromUnit}" (${fromEntry.family}) to "${baseUnitName}" ` +
      `(${baseEntry.family}) — set a Weight per ${baseUnitName} on this item first.`
    );
  }

  if (fromEntry.family === 'Weight' && baseEntry.family === 'Count') {
    const qtyInGrams = q * fromEntry.factorToBase;
    return qtyInGrams / weightPerBaseUnit;
  }

  throw new Error(`Unsupported unit conversion from "${fromUnit}" to "${baseUnitName}".`);
}

/**
 * Converts a rate (price per `fromUnit`) into a rate per the item's Base Unit.
 * Inverse of convertQtyToBaseUnit: converting 1 fromUnit yields N base units,
 * so the per-base-unit rate is rate / N.
 *
 * @param {number} rate
 * @param {string} fromUnit
 * @param {{baseUnit: string, weightPerBaseUnit?: number}} item
 * @param {Object} [unitsMap]
 * @returns {number} Rate per the item's Base Unit
 */
function convertRateToBaseUnit(rate, fromUnit, item, unitsMap) {
  const map = unitsMap || _getUnitsMap();
  const baseUnitsPerFromUnit = convertQtyToBaseUnit(1, fromUnit, item, map);
  if (baseUnitsPerFromUnit <= 0) return 0;
  return (Number(rate) || 0) / baseUnitsPerFromUnit;
}
