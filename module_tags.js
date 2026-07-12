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
