/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_tags.gs — COLOR MASTER & MODEL MASTER
 *
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Two small, rarely-edited name lists that exist purely to feed datalist
 * suggestions on the BOM (Color, Model/Product Name) and Production
 * (Color) forms — same role as module_units.js's Unit Master plays for
 * Base/Purchase Unit fields. Free text is still accepted everywhere these
 * are used; the master just gives consistent spelling to pick from instead
 * of every operator retyping "Red" slightly differently.
 *
 * Both masters share the same two-column shape (Name, Remarks), so the CRUD
 * here is implemented once against a sheetKey/label pair and exposed twice
 * under distinct function names for the client to call.
 * ═══════════════════════════════════════════════════════════════════════════
 */

const TAG_LOCK_TIMEOUT_MS = 15000;

// Maps each tag sheetKey to its own MASTER_DATA_CACHE_KEYS entry — the three
// tag types are independent lists/caches, so a write to one must never
// invalidate the other two.
const TAG_CACHE_KEY_BY_SHEET_KEY = Object.freeze({
  COLOR_MASTER: MASTER_DATA_CACHE_KEYS.COLOR_MASTER,
  MODEL_MASTER: MASTER_DATA_CACHE_KEYS.MODEL_MASTER,
  PROCESS_TYPE_MASTER: MASTER_DATA_CACHE_KEYS.PROCESS_TYPE_MASTER
});

/**
 * Run once manually from the Apps Script editor to create the sheet and
 * write its header row. Safe to call again — does nothing if already set up.
 * @private
 */
function _initTagSheet(sheetKey, label) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = APP_CONFIG.SHEETS[sheetKey];
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    sheet.getRange(1, 1, 1, 2)
      .setValues([['Name', 'Remarks']])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    SpreadsheetApp.flush();
    return buildResponse(true, null, `${label} sheet initialized successfully.`);
  } catch (error) {
    Log.error(`[_initTagSheet:${sheetKey}] Error:`, error.message);
    return buildResponse(false, null, `Failed to initialize ${label} sheet: ` + error.message);
  }
}

/** Run once manually to set up the Color Master sheet. */
function initColorMasterSheet() {
  return _initTagSheet('COLOR_MASTER', 'Color Master');
}

/** Run once manually to set up the Model Master sheet. */
function initModelMasterSheet() {
  return _initTagSheet('MODEL_MASTER', 'Model Master');
}

/** Run once manually to set up the Process Type Master sheet. */
function initProcessTypeMasterSheet() {
  return _initTagSheet('PROCESS_TYPE_MASTER', 'Process Type Master');
}

/**
 * Fetches a tag master sheet, auto-creating it (with headers) on first
 * access — same self-healing pattern as getBOMData()'s initBOMSheet() — so
 * nobody needs to manually run the one-time init function first.
 * @private
 */
function _getOrCreateTagSheet(sheetKey, label) {
  try {
    return getSheet(APP_CONFIG.SHEETS[sheetKey]);
  } catch (e) {
    _initTagSheet(sheetKey, label);
    return getSheet(APP_CONFIG.SHEETS[sheetKey]);
  }
}

/**
 * Retrieves every entry from a tag master sheet, sorted by name.
 * @returns {Object} API response with data: Array<{name, remarks}>
 * @private
 */
function _getTagData(sheetKey, label) {
  try {
    const sheet = _getOrCreateTagSheet(sheetKey, label);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, []);

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const tags = data
      .map(r => ({
        name: String(r[TAG_COL.NAME - 1] || '').trim(),
        remarks: String(r[TAG_COL.REMARKS - 1] || '').trim()
      }))
      .filter(t => t.name !== '');

    tags.sort((a, b) => a.name.localeCompare(b.name));

    return buildResponse(true, tags);
  } catch (error) {
    Log.error(`[_getTagData:${sheetKey}] Error:`, error.message);
    return buildResponse(false, null, `Failed to load ${label}: ` + error.message);
  }
}

function getColors() {
  return getCachedListResponse(MASTER_DATA_CACHE_KEYS.COLOR_MASTER, () => _getTagData('COLOR_MASTER', 'Color Master'));
}

function getModels() {
  return getCachedListResponse(MASTER_DATA_CACHE_KEYS.MODEL_MASTER, () => _getTagData('MODEL_MASTER', 'Model Master'));
}

function getProcessTypes() {
  return getCachedListResponse(MASTER_DATA_CACHE_KEYS.PROCESS_TYPE_MASTER, () => _getTagData('PROCESS_TYPE_MASTER', 'Process Type Master'));
}

/** Title-cases a single color word/phrase, e.g. "RED" -> "Red", "navy blue" -> "Navy Blue". @private */
function _titleCaseColor(s) {
  return s.toLowerCase().replace(/\b\w/g, ch => ch.toUpperCase());
}

/**
 * Scans Items Master (Item Name, Narration, Specification) for hyphen-joined
 * combinations of colors that already exist in Color Master — e.g. if "Red"
 * and "White" are both existing colors and an item name contains
 * "CRYSTA-BASKET---RED-WHITE", that combo is reported as one new entry
 * ("Red-White"). Only combinations are reported: a lone existing color
 * mentioned by itself is not re-reported, since it's already in the master
 * and a separate entry for it isn't needed. Nothing is guessed beyond the
 * colors already on file — there's no built-in color dictionary.
 * Does not write anything — the client confirms before calling saveColor
 * for each reported combo.
 * @returns {Object} API response with data: { newColors: string[], scannedCount: number }
 */
function extractColorsFromItemMaster() {
  try {
    const itemsSheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
    const lastRow = itemsSheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, { newColors: [], scannedCount: 0 });

    const rows = itemsSheet.getRange(2, 1, lastRow - 1, ITEMS_COL.SPECIFICATION).getValues();

    const existingColors = _getTagData('COLOR_MASTER', 'Color Master');
    const existingList = (existingColors.data || []).map(c => c.name);
    if (existingList.length === 0) {
      return buildResponse(true, { newColors: [], scannedCount: rows.length },
        'Color Master is empty — add the base colors first, then this can detect their combinations.');
    }
    const existingNames = new Set(existingList.map(n => n.toLowerCase()));

    // Longest name first, so a multi-word existing color (e.g. "Navy Blue")
    // is matched whole instead of matching just part of it.
    const sortedNames = [...existingList].sort((a, b) => b.length - a.length);
    const escaped = sortedNames.map(c => c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
    const colorAlt = '(?:' + escaped.join('|') + ')';
    // Two or more existing colors chained by single hyphens (e.g. "RED-WHITE")
    // — captured as one match so it's reported as a single combo entry.
    const pattern = new RegExp('\\b' + colorAlt + '(?:-' + colorAlt + ')+\\b', 'gi');

    const found = new Map(); // lowercase combo -> title-cased combo, hyphens preserved

    rows.forEach(row => {
      const text = [row[ITEMS_COL.ITEM_NAME - 1], row[ITEMS_COL.NARRATION - 1], row[ITEMS_COL.SPECIFICATION - 1]]
        .map(v => String(v || '')).join(' ');

      let match;
      pattern.lastIndex = 0;
      while ((match = pattern.exec(text)) !== null) {
        const combo = match[0].split('-').map(part => _titleCaseColor(part.trim())).join('-');
        const lower = combo.toLowerCase();
        if (!existingNames.has(lower) && !found.has(lower)) {
          found.set(lower, combo);
        }
      }
    });

    const newColors = Array.from(found.values()).sort((a, b) => a.localeCompare(b));
    return buildResponse(true, { newColors, scannedCount: rows.length });
  } catch (error) {
    Log.error('[extractColorsFromItemMaster] Error:', error.message);
    return buildResponse(false, null, 'Failed to scan Item Master for colors: ' + error.message);
  }
}

/**
 * Creates or renames a tag master entry.
 * @param {Object} formData - { name, remarks, originalName? }
 * @private
 */
function _saveTag(sheetKey, label, formData) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(TAG_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = _getOrCreateTagSheet(sheetKey, label);

    const newName = sanitizeString(formData.name || '', 'name');
    if (!newName) throw new Error('Name must not be empty.');

    const remarks = sanitizeString(formData.remarks || '', 'remarks');

    const isEdit = !!formData.originalName;
    const originalName = isEdit ? sanitizeString(formData.originalName, 'originalName') : newName;

    const lastRow = sheet.getLastRow();
    const data = lastRow >= 2 ? sheet.getRange(2, TAG_COL.NAME, lastRow - 1, 1).getValues() : [];

    let targetRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === originalName.toLowerCase()) {
        targetRow = i + 2;
        break;
      }
    }

    if (isEdit && targetRow === -1) throw new Error(`Original ${label} entry not found.`);
    if (!isEdit && targetRow !== -1) throw new Error(`"${newName}" already exists in ${label}.`);

    if (isEdit && newName.toLowerCase() !== originalName.toLowerCase()) {
      const dupeRow = data.findIndex(r => String(r[0]).trim().toLowerCase() === newName.toLowerCase());
      if (dupeRow !== -1) throw new Error(`Another entry named "${newName}" already exists in ${label}.`);
    }

    const rowData = [newName, remarks];

    if (isEdit) {
      sheet.getRange(targetRow, 1, 1, 2).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    SpreadsheetApp.flush();
    if (TAG_CACHE_KEY_BY_SHEET_KEY[sheetKey]) invalidateListCache(TAG_CACHE_KEY_BY_SHEET_KEY[sheetKey]);

    // Cascade a genuine rename to every sheet that stores this tag's name as
    // a plain string reference (mirrors _renameVendorEverywhere/_renameUnitEverywhere
    // — same class of bug, fixed for Vendor/Contractor/Unit but not Color/
    // ProcessType until now). Plain string compare, NOT case-insensitive —
    // a casing-only rename ("abc" -> "ABC") must still cascade, matching the
    // fix already applied to the other rename cascades for the same reason.
    if (isEdit && newName !== originalName) {
      if (sheetKey === 'COLOR_MASTER' && typeof _renameColorEverywhere === 'function') {
        _renameColorEverywhere(originalName, newName);
      } else if (sheetKey === 'PROCESS_TYPE_MASTER' && typeof _renameProcessTypeEverywhere === 'function') {
        _renameProcessTypeEverywhere(originalName, newName);
      }
      // MODEL_MASTER: no cascade — Model Master's name isn't stored as a
      // reference anywhere else in the schema (confirmed by a full-codebase
      // check), it's a standalone tag with no downstream rows to update.
    }

    const msg = isEdit ? `${label} entry updated successfully.` : `${label} entry added successfully.`;
    logAction(isEdit ? 'UPDATE' : 'CREATE', APP_CONFIG.SHEETS[sheetKey], newName, msg, 'SUCCESS');

    return buildResponse(true, { name: newName }, msg);
  } catch (error) {
    Log.error(`[_saveTag:${sheetKey}] Error:`, error.message);
    logAction('ERROR', `saveTag:${sheetKey}`, formData ? formData.name : 'unknown', error.message, 'ERROR');
    return buildResponse(false, null, `Failed to save ${label} entry: ` + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Cascades a Color Master rename to every sheet that stores a color name as
 * a plain string reference. Mirrors module_units.js#_renameUnitEverywhere's
 * shape/helper. Skips PROCESS_COMPONENTS_COL.COLOR_AXIS (an axis LABEL like
 * "Mudguard Color", not a Color Master name) and BOM_COL blank cells (mean
 * "common to every color", never a match).
 * @private
 */
function _renameColorEverywhere(oldName, newName) {
  const tOldRaw = String(oldName || '').trim();
  const tOld = tOldRaw.toLowerCase();
  const tNew = String(newName || '').trim();
  if (!tOld || !tNew || tOldRaw === tNew) return;

  // A color cell isn't always a single literal name — computeColorAxesForProcess/
  // _mergeLinkedAxes can produce (and an operator can hand-enter, e.g. a
  // Color Sub-Group's colorGroup or a Warehouse Pool Opening entry matching
  // an existing merged bucket) a COMPOSITE value joining 2+ independent
  // axes with COLOR_COMBO_DELIMITER (e.g. "BCP / Blue-White" — see
  // config.js). A plain whole-cell compare only ever matches a single-token
  // cell; this renames whichever ONE token matches and rejoins, so a
  // composite cell containing the renamed color as just one of its parts
  // still gets updated instead of silently going stale. A plain
  // (non-composite) value is unaffected — same exact-match behavior as
  // before.
  const renameColorToken = (value) => {
    const str = String(value == null ? '' : value);
    if (!str.includes(COLOR_COMBO_DELIMITER)) {
      return str.trim().toLowerCase() === tOld ? tNew : str;
    }
    let tokenChanged = false;
    const renamedTokens = str.split(COLOR_COMBO_DELIMITER).map(t => {
      if (t.trim().toLowerCase() === tOld) { tokenChanged = true; return tNew; }
      return t;
    });
    return tokenChanged ? renamedTokens.join(COLOR_COMBO_DELIMITER) : str;
  };

  const renameInColumn = (sheetName, col, startRow) => {
    let sheet;
    try {
      sheet = getSheet(sheetName);
    } catch (e) {
      return; // Sheet doesn't exist yet, nothing to rename
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < startRow) return;

    const range = sheet.getRange(startRow, col, lastRow - startRow + 1, 1);
    const values = range.getValues();
    let changed = false;

    for (let i = 0; i < values.length; i++) {
      if (!values[i][0]) continue; // blank cell (e.g. BOM's "common to every color") never matches
      const renamed = renameColorToken(values[i][0]);
      if (renamed !== values[i][0]) {
        values[i][0] = renamed;
        changed = true;
      }
    }

    if (changed) range.setValues(values);
  };

  // Process Components: COLOR_GROUP is 'COMMON', a single Color Master
  // name, or (a legacy cross-multiplied recipe row) a COLOR_COMBO_DELIMITER
  // -joined composite of several — renameInColumn/renameColorToken handles
  // all three; 'COMMON' only ever matches if a color is literally named that.
  renameInColumn(APP_CONFIG.SHEETS.PROCESS_COMPONENTS, PROCESS_COMPONENTS_COL.COLOR_GROUP, 2);

  // BOM: per-row color variant (blank cells mean "common to every color" and
  // never match a real color name, so they're correctly left untouched).
  renameInColumn(APP_CONFIG.SHEETS.BOM, BOM_COL.COLOR, 2);

  // Warehouse Pool Opening: manually-recorded balances, keyed partly by color.
  renameInColumn(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING, WAREHOUSE_POOL_OPENING_COL.COLOR, 2);

  // Process Color Links: a paired color lives in EITHER column on a row
  // (storage is directional A/B, but a color can appear in either slot).
  let linksSheet;
  try {
    linksSheet = getSheet(APP_CONFIG.SHEETS.PROCESS_COLOR_LINKS);
  } catch (e) {
    linksSheet = null;
  }
  if (linksSheet) {
    const lastRow = linksSheet.getLastRow();
    if (lastRow >= 2) {
      const range = linksSheet.getRange(2, PROCESS_COLOR_LINKS_COL.COLOR_A, lastRow - 1,
        PROCESS_COLOR_LINKS_COL.COLOR_B - PROCESS_COLOR_LINKS_COL.COLOR_A + 1);
      const values = range.getValues();
      let changed = false;
      for (let i = 0; i < values.length; i++) {
        for (let c = 0; c < values[i].length; c++) {
          if (!values[i][c]) continue;
          const renamed = renameColorToken(values[i][c]);
          if (renamed !== values[i][c]) {
            values[i][c] = renamed;
            changed = true;
          }
        }
      }
      if (changed) range.setValues(values);
    }
  }

  // Production: COLOR is a display string (comma-joined when a lot produced
  // multiple colors) and COLOR_BREAKDOWN is a JSON [{color, qty}, ...] array
  // — both need structured rewriting, not a whole-cell string match.
  let prodSheet;
  try {
    prodSheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
  } catch (e) {
    prodSheet = null;
  }
  if (prodSheet && prodSheet.getLastColumn() >= PRODUCTION_COL.COLOR_BREAKDOWN) {
    const lastRow = prodSheet.getLastRow();
    if (lastRow >= 2) {
      const range = prodSheet.getRange(2, PRODUCTION_COL.COLOR, lastRow - 1,
        PRODUCTION_COL.COLOR_BREAKDOWN - PRODUCTION_COL.COLOR + 1);
      const values = range.getValues();
      let changed = false;

      for (let i = 0; i < values.length; i++) {
        const colorCell = String(values[i][0] || '').trim();
        if (colorCell) {
          // Split on the comma that joins ENTRIES (one per colorBreakdown
          // row — see saveProduction's `color = colorBreakdown.map(...).join(', ')`
          // in module_production.js), THEN let renameColorToken handle each
          // entry's own possible COLOR_COMBO_DELIMITER composite — a plain
          // per-entry exact match here would miss a renamed color that's
          // only one axis of a composite entry like "BCP / Blue-White".
          const parts = colorCell.split(',').map(p => p.trim());
          const renamedParts = parts.map(p => renameColorToken(p));
          const rejoined = renamedParts.join(', ');
          if (rejoined !== colorCell) {
            values[i][0] = rejoined;
            changed = true;
          }
        }

        const breakdownRaw = String(values[i][1] || '').trim();
        if (breakdownRaw) {
          try {
            const parsed = JSON.parse(breakdownRaw);
            if (Array.isArray(parsed)) {
              let breakdownChanged = false;
              parsed.forEach(entry => {
                if (!entry || !entry.color) return;
                const renamed = renameColorToken(entry.color);
                if (renamed !== entry.color) {
                  entry.color = renamed;
                  breakdownChanged = true;
                }
              });
              if (breakdownChanged) {
                values[i][1] = JSON.stringify(parsed);
                changed = true;
              }
            }
          } catch (e) { /* malformed/legacy cell — leave as-is */ }
        }
      }

      if (changed) range.setValues(values);
    }
  }

  // Warehouse Pool bucket colors are always DERIVED (rebuilt from scratch
  // by recalculateWarehousePool) from the sheets just renamed above — recalc
  // now so the bucket keys pick up the new name immediately, instead of
  // silently staying on the old name until some unrelated write triggers it.
  if (typeof recalculateWarehousePool === 'function') {
    recalculateWarehousePool();
  }
}

/**
 * Cascades a Process Type Master rename to Process Master's own
 * PROCESS_TYPE column (the only other place a process type name is stored).
 * @private
 */
function _renameProcessTypeEverywhere(oldName, newName) {
  const tOldRaw = String(oldName || '').trim();
  const tOld = tOldRaw.toLowerCase();
  const tNew = String(newName || '').trim();
  if (!tOld || !tNew || tOldRaw === tNew) return;

  let sheet;
  try {
    sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_MASTER);
  } catch (e) {
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || sheet.getLastColumn() < PROCESS_COL.PROCESS_TYPE) return;

  const range = sheet.getRange(2, PROCESS_COL.PROCESS_TYPE, lastRow - 1, 1);
  const values = range.getValues();
  let changed = false;

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim().toLowerCase() === tOld) {
      values[i][0] = tNew;
      changed = true;
    }
  }

  if (changed) range.setValues(values);
}

function saveColor(formData) {
  return _saveTag('COLOR_MASTER', 'Color Master', formData);
}

function saveModel(formData) {
  return _saveTag('MODEL_MASTER', 'Model Master', formData);
}

function saveProcessType(formData) {
  return _saveTag('PROCESS_TYPE_MASTER', 'Process Type Master', formData);
}

/**
 * Deletes a tag master entry by name.
 * @private
 */
function _deleteTag(sheetKey, label, name) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(TAG_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS[sheetKey]);
    if (!sheet) throw new Error(`${label} sheet not found.`);

    const targetName = String(name || '').trim().toLowerCase();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error(`${label} entry not found.`);

    const data = sheet.getRange(2, TAG_COL.NAME, lastRow - 1, 1).getValues();
    let targetRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === targetName) {
        targetRow = i + 2;
        break;
      }
    }

    if (targetRow === -1) throw new Error(`${label} entry not found.`);

    sheet.deleteRow(targetRow);
    SpreadsheetApp.flush();
    if (TAG_CACHE_KEY_BY_SHEET_KEY[sheetKey]) invalidateListCache(TAG_CACHE_KEY_BY_SHEET_KEY[sheetKey]);

    logAction('DELETE', APP_CONFIG.SHEETS[sheetKey], name, `${label} entry deleted`, 'SUCCESS');
    return buildResponse(true, null, `"${name}" deleted from ${label}.`);
  } catch (error) {
    Log.error(`[_deleteTag:${sheetKey}] Error:`, error.message);
    logAction('ERROR', `deleteTag:${sheetKey}`, name, error.message, 'ERROR');
    return buildResponse(false, null, `Failed to delete ${label} entry: ` + error.message);
  } finally {
    lock.releaseLock();
  }
}

function deleteColor(name) {
  return _deleteTag('COLOR_MASTER', 'Color Master', name);
}

function deleteModel(name) {
  return _deleteTag('MODEL_MASTER', 'Model Master', name);
}

function deleteProcessType(name) {
  return _deleteTag('PROCESS_TYPE_MASTER', 'Process Type Master', name);
}

/**
 * Finds whichever Process Type Master name appears as a substring of a
 * Process Name — same matching style as App.Utils.getModelFromOutputItemName
 * on the client (models.find(... lower.includes ...)). Only ever returns a
 * name already on file in Process Type Master; nothing is invented. Longest
 * name first, so a multi-word type (e.g. "Quality Check") wins over a
 * shorter one that happens to also match part of the text.
 * @private
 */
function _matchProcessTypeInName(processName, typeNames) {
  const lower = String(processName || '').toLowerCase();
  const sorted = [...typeNames].sort((a, b) => b.length - a.length);
  return sorted.find(t => t && lower.includes(String(t).toLowerCase())) || '';
}

/**
 * Re-derives every Process row's Process Type from its Process Name against
 * the current Process Type Master list — overwrites whatever is already
 * stored (not just blanks), since a row can be left holding a stale type
 * string from before Process Type Master was edited/trimmed down. Only
 * names already on file in Process Type Master are ever used as the group
 * title; a row whose Process Name matches none of them is cleared to blank
 * so it falls under "General" instead of keeping a stale one-off value.
 * @returns {Object} API response with data: { processesUpdated: number }
 */
function importProcessTypesFromProcessNames() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(TAG_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const processSheet = getSheet(APP_CONFIG.SHEETS.PROCESS_MASTER);
    if (!processSheet) throw new Error('Process Master sheet not found.');
    ensureProcessTypeColumn(processSheet);

    const lastRow = processSheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, { processesUpdated: 0 }, 'No processes to import from.');

    const rows = processSheet.getRange(2, 1, lastRow - 1, PROCESS_COL.PROCESS_TYPE).getValues();

    const existingTypes = (_getTagData('PROCESS_TYPE_MASTER', 'Process Type Master').data || []).map(t => t.name);
    if (existingTypes.length === 0) {
      return buildResponse(true, { processesUpdated: 0 },
        'Process Type Master is empty — add types there first, then this can match them against Process Names.');
    }

    let processesUpdated = 0;
    rows.forEach((row, i) => {
      const processName = String(row[PROCESS_COL.PROCESS_NAME - 1] || '').trim();
      const currentType = String(row[PROCESS_COL.PROCESS_TYPE - 1] || '').trim();
      const match = processName ? _matchProcessTypeInName(processName, existingTypes) : '';
      if (match === currentType) return;
      processSheet.getRange(i + 2, PROCESS_COL.PROCESS_TYPE).setValue(match);
      processesUpdated++;
    });

    SpreadsheetApp.flush();
    if (processesUpdated > 0) {
      invalidateListCache(MASTER_DATA_CACHE_KEYS.PROCESS_ALL, MASTER_DATA_CACHE_KEYS.PROCESS_ACTIVE);
    }

    const msg = `Re-matched ${processesUpdated} process(es) against Process Type Master (unmatched ones were cleared to "General").`;
    logAction('IMPORT', APP_CONFIG.SHEETS.PROCESS_TYPE_MASTER, '', msg, 'SUCCESS');

    return buildResponse(true, { processesUpdated: processesUpdated }, msg);
  } catch (error) {
    Log.error('[importProcessTypesFromProcessNames] Error:', error.message);
    logAction('ERROR', 'importProcessTypesFromProcessNames', '', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to import process types: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}
