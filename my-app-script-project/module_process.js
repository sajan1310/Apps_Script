/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_process.gs — PROCESS MASTER MODULE
 *
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Master list of pluggable production process types (Frame Painting, Rim
 * Assembly, Frame Fitting, Bicycle Packing, ...). Sequence defines the
 * work-in-progress (WIP) chain used by module_production.js to validate
 * how much of one process's output may be consumed by the next.
 *
 * Sheet Layout (Process Master):
 * ───────────────────────────────────────────────────────────────────────────
 * Col A (1):   Process ID (e.g. PRC-1001)
 * Col B (2):   Process Name
 * Col C (3):   Sequence (integer order in the WIP chain)
 * Col D (4):   Lot Prefix (e.g. "FP")
 * Col E (5):   Is Final Stage (TRUE/FALSE)
 * Col F (6):   Active (TRUE/FALSE)
 * Col G (7):   Remarks
 * Col H (8):   Output Item Name (Warehouse Pool item this process produces per unit)
 * ═══════════════════════════════════════════════════════════════════════════
 */

const PROCESS_LOCK_TIMEOUT_MS = 15000;

/**
 * Initializes the Process Master sheet with correct headers, and seeds it
 * with the standard bicycle production stages on first creation.
 */
function initProcessMasterSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.PROCESS_MASTER);
    const isNew = !sheet;
    if (!sheet) {
      sheet = ss.insertSheet(APP_CONFIG.SHEETS.PROCESS_MASTER);
    }

    const headers = [
      'Process ID',
      'Process Name',
      'Sequence',
      'Lot Prefix',
      'Is Final Stage',
      'Active',
      'Remarks',
      'Output Item Name'
    ];

    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    if (isNew) {
      const seedRows = [
        ['PRC-1001', 'Frame Painting', 1, 'FP', false, true, 'Auto-seeded default process', 'Painted Frame'],
        ['PRC-1002', 'Rim Assembly', 2, 'RA', false, true, 'Auto-seeded default process', 'Fitted Rim'],
        ['PRC-1003', 'Frame Fitting', 3, 'FF', false, true, 'Auto-seeded default process', 'Fitted Frame'],
        ['PRC-1004', 'Bicycle Packing', 4, 'PK', true, true, 'Auto-seeded default process', 'Packed Bicycle']
      ];
      sheet.getRange(2, 1, seedRows.length, headers.length).setValues(seedRows);
    }

    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Process Master sheet initialized successfully.');
  } catch (error) {
    console.error('[initProcessMasterSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize Process Master sheet: ' + error.message);
  }
}

/**
 * Retrieves all processes, sorted by Sequence ascending.
 * @param {boolean} [activeOnly] - If true, excludes inactive processes.
 */
function getProcessData(activeOnly) {
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_MASTER);
    } catch (e) {
      initProcessMasterSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_MASTER);
    }

    ensureProcessOutputItemColumn(sheet);
    ensureProcessTypeColumn(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, []);

    const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    const records = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const processId = String(row[PROCESS_COL.PROCESS_ID - 1] || '').trim();
      const processName = String(row[PROCESS_COL.PROCESS_NAME - 1] || '').trim();
      if (!processId || !processName) continue;

      const isActive = row[PROCESS_COL.ACTIVE - 1] === true || String(row[PROCESS_COL.ACTIVE - 1]).toUpperCase() === 'TRUE';
      if (activeOnly && !isActive) continue;

      records.push({
        rowIdx: i + 2,
        processId: processId,
        processName: processName,
        sequence: Number(row[PROCESS_COL.SEQUENCE - 1]) || 0,
        lotPrefix: String(row[PROCESS_COL.LOT_PREFIX - 1] || '').trim().toUpperCase(),
        isFinalStage: row[PROCESS_COL.IS_FINAL_STAGE - 1] === true || String(row[PROCESS_COL.IS_FINAL_STAGE - 1]).toUpperCase() === 'TRUE',
        active: isActive,
        remarks: String(row[PROCESS_COL.REMARKS - 1] || '').trim(),
        outputItemName: String(row[PROCESS_COL.OUTPUT_ITEM_NAME - 1] || '').trim(),
        processType: String(row[PROCESS_COL.PROCESS_TYPE - 1] || '').trim()
      });
    }

    records.sort((a, b) => a.sequence - b.sequence);

    return buildResponse(true, records);
  } catch (error) {
    console.error('[getProcessData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load process data: ' + error.message);
  }
}

/**
 * Backfills the "Output Item Name" column on Process Master sheets created
 * before the Warehouse Pool feature existed.
 */
function ensureProcessOutputItemColumn(sheet) {
  try {
    if (sheet.getLastColumn() < PROCESS_COL.OUTPUT_ITEM_NAME) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), PROCESS_COL.OUTPUT_ITEM_NAME - sheet.getLastColumn());
      sheet.getRange(1, PROCESS_COL.OUTPUT_ITEM_NAME, 1, 1)
        .setValues([['Output Item Name']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    console.error('[ensureProcessOutputItemColumn] Error:', error.message);
  }
}

/**
 * Backfills the "Process Type" column on Process Master sheets created
 * before the Process Type Master grouping feature existed.
 */
function ensureProcessTypeColumn(sheet) {
  try {
    if (sheet.getLastColumn() < PROCESS_COL.PROCESS_TYPE) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), PROCESS_COL.PROCESS_TYPE - sheet.getLastColumn());
      sheet.getRange(1, PROCESS_COL.PROCESS_TYPE, 1, 1)
        .setValues([['Process Type']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    console.error('[ensureProcessTypeColumn] Error:', error.message);
  }
}

/**
 * Auto-generates the next sequential Process ID.
 * Format: PRC-1001, PRC-1002, ...
 */
function getNextProcessId() {
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_MASTER);
    } catch (e) {
      initProcessMasterSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_MASTER);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return 'PRC-1001';

    const ids = sheet.getRange(2, PROCESS_COL.PROCESS_ID, lastRow - 1, 1).getValues();
    let maxNum = 1000;

    ids.forEach(row => {
      const idStr = String(row[0] || '').trim();
      const match = idStr.match(/^PRC-(\d+)$/i);
      if (match) {
        const num = parseInt(match[1], 10);
        if (num > maxNum) maxNum = num;
      }
    });

    return 'PRC-' + (maxNum + 1);
  } catch (error) {
    console.error('[getNextProcessId] Error:', error.message);
    return 'PRC-1001';
  }
}

/**
 * Creates a new process or updates an existing one.
 */
function saveProcess(formData) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(PROCESS_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_MASTER);
    } catch (e) {
      initProcessMasterSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_MASTER);
    }
    ensureProcessOutputItemColumn(sheet);
    ensureProcessTypeColumn(sheet);

    const processName = sanitizeString(formData.processName, 'processName');
    if (!processName) {
      return buildResponse(false, null, 'Process Name is required.');
    }

    const lotPrefix = sanitizeString(formData.lotPrefix, 'lotPrefix').toUpperCase();
    if (!lotPrefix || !/^[A-Z0-9]{1,6}$/.test(lotPrefix)) {
      return buildResponse(false, null, 'Lot Prefix is required (1-6 letters/numbers, e.g. "FP").');
    }

    const outputItemName = sanitizeString(formData.outputItemName, 'outputItemName');
    if (!outputItemName) {
      return buildResponse(false, null, 'Output Item Name is required (the Warehouse Pool item this process produces).');
    }

    // Optional — categorizes the process for grouping (see Process Type
    // Master). Blank is allowed; the UI groups untyped processes under
    // "General", same convention as Size/Model.
    const processType = sanitizeString(formData.processType || '', 'processType');

    const sequence = validateNumber(formData.sequence, 1, 100000);
    if (sequence <= 0) {
      return buildResponse(false, null, 'Sequence must be a positive number.');
    }

    const isFinalStage = !!formData.isFinalStage;
    const active = formData.active === undefined ? true : !!formData.active;
    const remarks = sanitizeString(formData.remarks || '', 'remarks').slice(0, APP_CONFIG.VALIDATION.MAX_REMARKS_LENGTH);

    let components;
    try {
      components = typeof formData.components === 'string'
        ? JSON.parse(formData.components)
        : (formData.components || []);
    } catch (e) {
      return buildResponse(false, null, 'Invalid component data format.');
    }
    if (!Array.isArray(components)) components = [];

    const dupComponent = _findDuplicateComponent(components);
    if (dupComponent) {
      return buildResponse(false, null,
        `Duplicate component: "${dupComponent.itemName}"${dupComponent.size ? ' (' + dupComponent.size + ')' : ''} already exists in ${dupComponent.colorGroup === COMPONENT_COLOR_GROUP_COMMON ? 'Common Components' : 'the "' + dupComponent.colorGroup + '" color sub-group'}. Each item+size combination may only appear once per group — adjust its Qty / Unit instead of adding it twice.`);
    }

    const isEdit = !!formData.processId;
    const lastRow = sheet.getLastRow();

    // Duplicate Lot Prefix / Sequence checks (excluding the row being edited)
    if (lastRow >= 2) {
      const existing = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
      for (let i = 0; i < existing.length; i++) {
        const rowProcessId = String(existing[i][PROCESS_COL.PROCESS_ID - 1] || '').trim();
        if (isEdit && rowProcessId.toLowerCase() === String(formData.processId).trim().toLowerCase()) continue;

        const rowPrefix = String(existing[i][PROCESS_COL.LOT_PREFIX - 1] || '').trim().toUpperCase();
        if (rowPrefix === lotPrefix) {
          return buildResponse(false, null, `Lot Prefix "${lotPrefix}" is already used by another process.`);
        }
      }
    }

    if (isEdit) {
      const processId = sanitizeString(formData.processId, 'processId');
      const data = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];
      let targetRow = -1;
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim().toLowerCase() === processId.toLowerCase()) {
          targetRow = i + 2;
          break;
        }
      }

      if (targetRow === -1) {
        return buildResponse(false, null, `Process with ID "${processId}" not found.`);
      }

      const oldOutputItemName = String(sheet.getRange(targetRow, PROCESS_COL.OUTPUT_ITEM_NAME).getValue() || '').trim();

      sheet.getRange(targetRow, 1, 1, 9).setValues([[
        processId, processName, sequence, lotPrefix, isFinalStage, active, remarks, outputItemName, processType
      ]]);

      _saveProcessComponentsForProcess(processId, components);

      // Renaming the Output Item Name doesn't retroactively touch Production
      // lots already saved under the old name (it's de-normalized onto each
      // lot at save time — see PRODUCTION_COL.OUTPUT_ITEM_NAME), so without
      // this, old completed lots keep crediting the Warehouse Pool under the
      // stale name while new BOMs look up the new name and see zero stock.
      if (oldOutputItemName && oldOutputItemName.toLowerCase() !== outputItemName.toLowerCase()) {
        _renamePoolOutputItemNameEverywhere(oldOutputItemName, outputItemName);
        if (typeof recalculateWarehousePool === 'function') {
          recalculateWarehousePool();
        }
      }

      SpreadsheetApp.flush();
      logAction('UPDATE', APP_CONFIG.SHEETS.PROCESS_MASTER, processId, `Process updated: ${processName}`, 'SUCCESS');
      return buildResponse(true, { processId: processId }, 'Process updated successfully.');
    }

    const newProcessId = getNextProcessId();
    sheet.appendRow([newProcessId, processName, sequence, lotPrefix, isFinalStage, active, remarks, outputItemName, processType]);

    _saveProcessComponentsForProcess(newProcessId, components);

    SpreadsheetApp.flush();
    logAction('CREATE', APP_CONFIG.SHEETS.PROCESS_MASTER, newProcessId, `Process created: ${processName}`, 'SUCCESS');
    return buildResponse(true, { processId: newProcessId }, 'Process created successfully.');
  } catch (error) {
    console.error('[saveProcess] Error:', error.message);
    logAction('ERROR', 'saveProcess', formData.processId || 'NEW', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to save process: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * @private
 * Propagates a Process Master Output Item Name rename to every place that
 * holds a de-normalized snapshot of it, so the Warehouse Pool stays keyed
 * consistently under the new name:
 *  - Production sheet: each lot's own OUTPUT_ITEM_NAME (the credit side)
 *  - Production sheet: any lot's COMPONENTS_CONSUMED JSON that consumed this
 *    item as a POOL-sourced component (the debit side, on other processes' lots)
 *  - Process Components sheet: any other process's recipe that references
 *    this item as a POOL-sourced component
 * Matching is case-insensitive/trimmed; values are rewritten to the new
 * name exactly as entered.
 */
function _renamePoolOutputItemNameEverywhere(oldName, newName) {
  const oldLower = String(oldName || '').trim().toLowerCase();
  const newTrimmed = String(newName || '').trim();
  if (!oldLower || !newTrimmed || oldLower === newTrimmed.toLowerCase()) return;

  let prodSheet;
  try {
    prodSheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
  } catch (e) {
    prodSheet = null;
  }

  if (prodSheet) {
    const lastRow = prodSheet.getLastRow();
    if (lastRow >= 2) {
      const outputRange = prodSheet.getRange(2, PRODUCTION_COL.OUTPUT_ITEM_NAME, lastRow - 1, 1);
      const outputCol = outputRange.getValues();
      let outputChanged = false;
      outputCol.forEach(row => {
        if (String(row[0] || '').trim().toLowerCase() === oldLower) {
          row[0] = newTrimmed;
          outputChanged = true;
        }
      });
      if (outputChanged) outputRange.setValues(outputCol);

      const componentsRange = prodSheet.getRange(2, PRODUCTION_COL.COMPONENTS_CONSUMED, lastRow - 1, 1);
      const componentsCol = componentsRange.getValues();
      let componentsChanged = false;
      componentsCol.forEach(row => {
        const raw = String(row[0] || '').trim();
        if (!raw) return;
        let parsed;
        try {
          parsed = JSON.parse(raw);
        } catch (e) {
          return;
        }
        if (!Array.isArray(parsed)) return;

        let rowChanged = false;
        parsed.forEach(comp => {
          if (String(comp.sourceType || '').trim().toUpperCase() === COMPONENT_SOURCE_TYPES.POOL &&
              String(comp.itemName || '').trim().toLowerCase() === oldLower) {
            comp.itemName = newTrimmed;
            rowChanged = true;
          }
        });
        if (rowChanged) {
          row[0] = JSON.stringify(parsed);
          componentsChanged = true;
        }
      });
      if (componentsChanged) componentsRange.setValues(componentsCol);
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
      const range = compSheet.getRange(2, 1, lastRow - 1, PROCESS_COMPONENTS_COL.SOURCE_TYPE);
      const rows = range.getValues();
      let changed = false;
      rows.forEach(row => {
        const sourceType = String(row[PROCESS_COMPONENTS_COL.SOURCE_TYPE - 1] || '').trim().toUpperCase();
        const itemName = String(row[PROCESS_COMPONENTS_COL.ITEM_NAME - 1] || '').trim();
        if (sourceType === COMPONENT_SOURCE_TYPES.POOL && itemName.toLowerCase() === oldLower) {
          row[PROCESS_COMPONENTS_COL.ITEM_NAME - 1] = newTrimmed;
          changed = true;
        }
      });
      if (changed) range.setValues(rows);
    }
  }
}

/**
 * ONE-TIME REPAIR — run manually from the Apps Script editor to fix data
 * left stale by a Process Master Output Item Name rename that happened
 * before _renamePoolOutputItemNameEverywhere() was wired into saveProcess().
 * Example: repairPoolOutputItemNameRename('Fitted Rim 14 inch', 'Fitted Rim 14 inch ED');
 */
function repairPoolOutputItemNameRename(oldName, newName) {
  _renamePoolOutputItemNameEverywhere(oldName, newName);
  const result = recalculateWarehousePool();
  console.log('[repairPoolOutputItemNameRename] ' + (result && result.message));
  return result;
}

/**
 * ONE-CLICK REPAIR for the "Fitted Rim 14 inch" -> "Fitted Rim 14 inch ED"
 * rename. Run this once from the Apps Script editor (select it in the
 * function dropdown and click Run), then delete it — it's a single-use fix.
 */
function repairFittedRim14InchEDRename() {
  return repairPoolOutputItemNameRename('Fitted Rim 14 inch', 'Fitted Rim 14 inch ED');
}

/**
 * Deletes a process. Blocked if any Production lot already references it.
 */
function deleteProcess(processId) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(PROCESS_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_MASTER);
    if (!sheet) throw new Error('Process Master sheet not found.');

    const idClean = String(processId || '').trim();
    if (!idClean) return buildResponse(false, null, 'Process ID is required.');

    let prodSheet;
    try {
      prodSheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
    } catch (e) {
      prodSheet = null;
    }

    if (prodSheet) {
      const pLastRow = prodSheet.getLastRow();
      if (pLastRow >= 2) {
        const refs = prodSheet.getRange(2, PRODUCTION_COL.PROCESS_ID, pLastRow - 1, 1).getValues();
        const inUse = refs.some(row => String(row[0]).trim().toLowerCase() === idClean.toLowerCase());
        if (inUse) {
          return buildResponse(false, null, `Cannot delete process: "${idClean}" is already referenced by Production lots.`);
        }
      }
    }

    const rowsDeleted = deleteRowsById(idClean, sheet, 2, PROCESS_COL.PROCESS_ID);

    // Also remove this process's component checklist rows
    try {
      const compSheet = getSheet(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);
      deleteRowsById(idClean, compSheet, 2, PROCESS_COMPONENTS_COL.PROCESS_ID);
    } catch (e) {
      // Process Components sheet doesn't exist yet, nothing to clean up
    }

    SpreadsheetApp.flush();

    const msg = `Process "${idClean}" deleted (${rowsDeleted} row(s) removed).`;
    logAction('DELETE', APP_CONFIG.SHEETS.PROCESS_MASTER, idClean, msg, 'SUCCESS');

    return buildResponse(true, null, msg);
  } catch (error) {
    console.error('[deleteProcess] Error:', error.message);
    logAction('ERROR', 'deleteProcess', processId, error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete process: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes multiple processes at once. Processes already referenced by
 * Production lots are skipped (mirrors deleteProcess's single-item check).
 * @param {string[]} processIds
 */
function deleteProcessesBulk(processIds) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(PROCESS_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_MASTER);
    if (!sheet) throw new Error('Process Master sheet not found.');

    const requested = (processIds || []).map(id => String(id || '').trim()).filter(Boolean);
    if (requested.length === 0) {
      return buildResponse(true, null, 'No processes selected.');
    }

    let inUseSet = new Set();
    let prodSheet;
    try {
      prodSheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
    } catch (e) {
      // Production sheet does not exist yet, so nothing is in use
    }

    if (prodSheet) {
      const pLastRow = prodSheet.getLastRow();
      if (pLastRow >= 2) {
        const refs = prodSheet.getRange(2, PRODUCTION_COL.PROCESS_ID, pLastRow - 1, 1).getValues();
        const usedIds = new Set(refs.map(row => String(row[0]).trim().toLowerCase()));
        requested.forEach(id => {
          if (usedIds.has(id.toLowerCase())) inUseSet.add(id);
        });
      }
    }

    const toDelete = requested.filter(id => !inUseSet.has(id));
    const targetSet = new Set(toDelete);

    let rowsDeleted = 0;
    if (targetSet.size > 0) {
      rowsDeleted = _rewriteWithoutMatchingRowsBulk(sheet, 2, PROCESS_COL.PROCESS_ID, targetSet).rowsDeleted;

      try {
        const compSheet = getSheet(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);
        _rewriteWithoutMatchingRowsBulk(compSheet, 2, PROCESS_COMPONENTS_COL.PROCESS_ID, targetSet);
      } catch (e) {
        // Process Components sheet doesn't exist yet, nothing to clean up
      }

      SpreadsheetApp.flush();
    }

    let msg = `Deleted ${toDelete.length} process(es) (${rowsDeleted} rows removed).`;
    if (inUseSet.size > 0) {
      msg += ` Skipped ${inUseSet.size} process(es) still in use by Production: ${Array.from(inUseSet).join(', ')}.`;
    }
    logAction('BULK_DELETE', APP_CONFIG.SHEETS.PROCESS_MASTER, 'multiple', msg, 'SUCCESS');

    return buildResponse(true, { skipped: Array.from(inUseSet) }, msg);
  } catch (error) {
    console.error('[deleteProcessesBulk] Error:', error.message);
    logAction('ERROR', 'deleteProcessesBulk', 'multiple', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete processes: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Persists a manual drag-and-drop reorder from the Processes tab. Only
 * meaningful for the flat (ungrouped) view — rewrites Sequence 1..N for
 * every row matching a processId in orderedProcessIds, in that order.
 * @param {string[]} orderedProcessIds - Process IDs in their new desired order
 */
function reorderProcesses(orderedProcessIds) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(PROCESS_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_MASTER);
    if (!sheet) throw new Error('Process Master sheet not found.');

    const order = (orderedProcessIds || []).map(id => String(id || '').trim()).filter(Boolean);
    if (order.length === 0) {
      return buildResponse(false, null, 'No process order provided.');
    }

    const newSeqById = {};
    order.forEach((id, i) => { newSeqById[id.toLowerCase()] = i + 1; });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return buildResponse(true, null, 'No processes to reorder.');
    }

    const ids = sheet.getRange(2, PROCESS_COL.PROCESS_ID, lastRow - 1, 1).getValues();
    const sequences = sheet.getRange(2, PROCESS_COL.SEQUENCE, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      const id = String(ids[i][0] || '').trim().toLowerCase();
      if (newSeqById.hasOwnProperty(id)) {
        sequences[i][0] = newSeqById[id];
      }
    }
    sheet.getRange(2, PROCESS_COL.SEQUENCE, sequences.length, 1).setValues(sequences);

    SpreadsheetApp.flush();

    const msg = `Reordered ${order.length} process(es).`;
    logAction('UPDATE', APP_CONFIG.SHEETS.PROCESS_MASTER, 'multiple', msg, 'SUCCESS');

    return buildResponse(true, null, msg);
  } catch (error) {
    console.error('[reorderProcesses] Error:', error.message);
    logAction('ERROR', 'reorderProcesses', 'multiple', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to reorder processes: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

// ── Process Components (per-process item checklist) ────────────────────

/**
 * Initializes the Process Components sheet with correct headers.
 */
function initProcessComponentsSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);
    if (!sheet) {
      sheet = ss.insertSheet(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);
    }
    sheet.getRange(1, 1, 1, 8)
      .setValues([['Process ID', 'Item Name', 'Size', 'Narration', 'Qty Per Unit', 'Remarks', 'Source Type', 'Color Group']])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Process Components sheet initialized successfully.');
  } catch (error) {
    console.error('[initProcessComponentsSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize Process Components sheet: ' + error.message);
  }
}

/**
 * Backfills the "Qty Per Unit" column on Process Components sheets created
 * before the process-recipe feature existed, so legacy rows don't throw
 * when read/written.
 */
function ensureProcessComponentsQtyColumn(sheet) {
  try {
    if (sheet.getLastColumn() < PROCESS_COMPONENTS_COL.REMARKS) {
      const startCol = sheet.getLastColumn() + 1;
      sheet.insertColumnsAfter(sheet.getLastColumn(), PROCESS_COMPONENTS_COL.REMARKS - sheet.getLastColumn());
      const headers = startCol === PROCESS_COMPONENTS_COL.QTY_PER_UNIT
        ? [['Qty Per Unit', 'Remarks']]
        : [['Remarks']];
      sheet.getRange(1, startCol, 1, PROCESS_COMPONENTS_COL.REMARKS - startCol + 1)
        .setValues(headers)
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    console.error('[ensureProcessComponentsQtyColumn] Error:', error.message);
  }
}

/**
 * Backfills the "Source Type" column on Process Components sheets created
 * before the Warehouse Pool feature existed. Existing rows default to ITEM.
 */
function ensureProcessComponentsSourceTypeColumn(sheet) {
  try {
    if (sheet.getLastColumn() < PROCESS_COMPONENTS_COL.SOURCE_TYPE) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), PROCESS_COMPONENTS_COL.SOURCE_TYPE - sheet.getLastColumn());
      sheet.getRange(1, PROCESS_COMPONENTS_COL.SOURCE_TYPE, 1, 1)
        .setValues([['Source Type']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');

      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const col = sheet.getRange(2, PROCESS_COMPONENTS_COL.SOURCE_TYPE, lastRow - 1, 1);
        const blanks = col.getValues().map(() => [COMPONENT_SOURCE_TYPES.ITEM]);
        col.setValues(blanks);
      }
    }
  } catch (error) {
    console.error('[ensureProcessComponentsSourceTypeColumn] Error:', error.message);
  }
}

/**
 * Backfills the "Color Group" column on Process Components sheets created
 * before the color sub-group feature existed. Existing rows default to
 * COMMON (shared across every color variant), matching prior behavior where
 * every recipe row applied to every lot.
 */
function ensureProcessComponentsColorGroupColumn(sheet) {
  try {
    if (sheet.getLastColumn() < PROCESS_COMPONENTS_COL.COLOR_GROUP) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), PROCESS_COMPONENTS_COL.COLOR_GROUP - sheet.getLastColumn());
      sheet.getRange(1, PROCESS_COMPONENTS_COL.COLOR_GROUP, 1, 1)
        .setValues([['Color Group']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');

      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const col = sheet.getRange(2, PROCESS_COMPONENTS_COL.COLOR_GROUP, lastRow - 1, 1);
        const blanks = col.getValues().map(() => [COMPONENT_COLOR_GROUP_COMMON]);
        col.setValues(blanks);
      }
    }
  } catch (error) {
    console.error('[ensureProcessComponentsColorGroupColumn] Error:', error.message);
  }
}

/**
 * Retrieves the component checklist for a single process.
 * @param {string} processId
 */
function getProcessComponentsData(processId) {
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);
    } catch (e) {
      initProcessComponentsSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);
    }

    ensureProcessComponentsQtyColumn(sheet);
    ensureProcessComponentsSourceTypeColumn(sheet);
    ensureProcessComponentsColorGroupColumn(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, []);

    const targetId = String(processId || '').trim().toLowerCase();
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

    const components = data
      .map(row => ({
        processId: String(row[PROCESS_COMPONENTS_COL.PROCESS_ID - 1] || '').trim(),
        itemName: String(row[PROCESS_COMPONENTS_COL.ITEM_NAME - 1] || '').trim(),
        size: String(row[PROCESS_COMPONENTS_COL.SIZE - 1] || '').trim(),
        narration: String(row[PROCESS_COMPONENTS_COL.NARRATION - 1] || '').trim(),
        qtyPerUnit: Number(row[PROCESS_COMPONENTS_COL.QTY_PER_UNIT - 1]) || 0,
        remarks: String(row[PROCESS_COMPONENTS_COL.REMARKS - 1] || '').trim(),
        sourceType: String(row[PROCESS_COMPONENTS_COL.SOURCE_TYPE - 1] || '').trim().toUpperCase() === COMPONENT_SOURCE_TYPES.POOL
          ? COMPONENT_SOURCE_TYPES.POOL
          : COMPONENT_SOURCE_TYPES.ITEM,
        colorGroup: String(row[PROCESS_COMPONENTS_COL.COLOR_GROUP - 1] || '').trim() || COMPONENT_COLOR_GROUP_COMMON
      }))
      .filter(c => c.itemName && (!targetId || c.processId.toLowerCase() === targetId));

    return buildResponse(true, components);
  } catch (error) {
    console.error('[getProcessComponentsData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load process components: ' + error.message);
  }
}

/**
 * Returns the distinct color sub-group names a process should offer on the
 * Production Lot form, sorted alphabetically. Two independent sources feed
 * this, unioned together:
 *   1. Colors explicitly configured on the process's own recipe (a
 *      component row scoped to a Color Master name instead of COMMON) —
 *      the original mechanism, still useful for color-specific raw
 *      materials (e.g. a specific paint).
 *   2. Colors that currently exist in the Warehouse Pool for any
 *      POOL-sourced component this recipe consumes — so a downstream
 *      process (e.g. Frame Fitting consuming Painted Frame) automatically
 *      becomes color-selectable the moment its upstream process has
 *      actually produced more than one color, with no manual recipe setup
 *      required. This reflects live pool state, not configuration, so it
 *      can change as upstream production happens.
 * An empty result means neither source found a color variant, and
 * Production logging proceeds with just the COMMON components as before.
 * @param {string} processId
 */
function getProcessColorGroups(processId) {
  try {
    const componentsResp = getProcessComponentsData(processId);
    const components = (componentsResp && componentsResp.data) || [];
    const poolRows = typeof getWarehousePoolData === 'function'
      ? ((getWarehousePoolData() || {}).data || [])
      : [];
    return buildResponse(true, computeColorGroupsForProcess(components, poolRows));
  } catch (error) {
    console.error('[getProcessColorGroups] Error:', error.message);
    return buildResponse(false, null, 'Failed to load process color groups: ' + error.message);
  }
}

/**
 * Shared core of getProcessColorGroups — pulled out so the bulk variant
 * (getAllProcessColorGroups) can read the Process Components and Warehouse
 * Pool sheets ONCE for every process instead of once per process.
 * @param {Array} components Process Components rows already filtered to one process.
 * @param {Array} poolRows Full Warehouse Pool rows (shared across all processes).
 */
function computeColorGroupsForProcess(components, poolRows) {
  const colors = new Set();
  components.forEach(c => {
    if (c.colorGroup && c.colorGroup !== COMPONENT_COLOR_GROUP_COMMON) colors.add(c.colorGroup);
  });

  const poolItemNames = new Set(
    components.filter(c => c.sourceType === COMPONENT_SOURCE_TYPES.POOL).map(c => c.itemName.toLowerCase())
  );
  if (poolItemNames.size > 0) {
    // Only count colors per pool item — an item that exists in just ONE
    // pool color (e.g. a Fitted Rim that's always Black) isn't a real
    // output-color choice for this lot, it's a fixed input, so it must
    // not leak into the output checklist. Only items with 2+ pool colors
    // (e.g. a Painted Frame painted to match each output color) are a
    // genuine per-output-color variant and contribute to this union.
    const colorsByItem = new Map();
    poolRows.forEach(r => {
      const key = r.outputItemName.toLowerCase();
      if (!r.color || !poolItemNames.has(key)) return;
      if (!colorsByItem.has(key)) colorsByItem.set(key, new Set());
      colorsByItem.get(key).add(r.color);
    });

    // Items sharing the EXACT same color set (e.g. two recipe rows both
    // sourced from the same upstream "Painted Frame") are one color axis.
    // Items with genuinely different color sets (e.g. a Fitted Rim colored
    // BCP/Black assembled alongside a Frame colored Blue-White/Orange-White)
    // are independent axes — a real output unit needs one color from EACH
    // axis at once, so the producible variants are the cross product of all
    // axes, not a flat union (which would falsely offer "BCP" alone as a
    // complete output color with no frame color attached).
    const axesBySignature = new Map();
    colorsByItem.forEach(itemColors => {
      if (itemColors.size <= 1) return;
      const sorted = Array.from(itemColors).sort((a, b) => a.localeCompare(b));
      const signature = sorted.map(c => c.toLowerCase()).join('|');
      if (!axesBySignature.has(signature)) axesBySignature.set(signature, sorted);
    });

    const axes = Array.from(axesBySignature.values());
    if (axes.length === 1) {
      axes[0].forEach(color => colors.add(color));
    } else if (axes.length > 1) {
      let combos = [''];
      axes.forEach(axisColors => {
        const next = [];
        combos.forEach(prefix => {
          axisColors.forEach(color => next.push(prefix ? `${prefix}${COLOR_COMBO_DELIMITER}${color}` : color));
        });
        combos = next;
      });
      combos.forEach(combo => colors.add(combo));
    }
  }

  return Array.from(colors).sort((a, b) => a.localeCompare(b));
}

/**
 * Bulk variant of getProcessColorGroups — returns every process's color
 * groups in one call, keyed by Process ID. Reads Process Components and
 * Warehouse Pool ONCE (not once per process — see computeColorGroupsForProcess)
 * so this stays fast regardless of process count. Used by the Warehouse Pool
 * table to list every known color variant of each process (not just the ones
 * that already have a stock bucket), so the user can add initial/opening
 * stock for a variant that hasn't produced anything yet.
 */
function getAllProcessColorGroups() {
  try {
    const processResp = getProcessData(false);
    const processes = (processResp && processResp.data) || [];

    const allComponentsResp = getProcessComponentsData('');
    const allComponents = (allComponentsResp && allComponentsResp.data) || [];
    const componentsByProcess = new Map();
    allComponents.forEach(c => {
      if (!componentsByProcess.has(c.processId)) componentsByProcess.set(c.processId, []);
      componentsByProcess.get(c.processId).push(c);
    });

    const poolResp = typeof getWarehousePoolData === 'function' ? getWarehousePoolData() : null;
    const poolRows = (poolResp && poolResp.data) || [];

    const result = {};
    processes.forEach(p => {
      const components = componentsByProcess.get(p.processId) || [];
      result[p.processId] = computeColorGroupsForProcess(components, poolRows);
    });

    return buildResponse(true, result);
  } catch (error) {
    console.error('[getAllProcessColorGroups] Error:', error.message);
    return buildResponse(false, null, 'Failed to load process color groups: ' + error.message);
  }
}

/**
 * Scans a process's submitted component list for a duplicate item, where
 * uniqueness = Item Name + Size + Color Group (Common counts as one group,
 * same as every named color sub-group counts as its own group — see
 * View_Products.html's Common Components / Color Sub-Groups sections).
 * Matching is case-insensitive/trimmed. Returns the first duplicate's own
 * (already-trimmed) field values for the error message, or null if none.
 * @param {Array<Object>} components
 * @returns {{itemName: string, size: string, colorGroup: string}|null}
 * @private
 */
function _findDuplicateComponent(components) {
  const seen = new Set();
  for (let i = 0; i < (components || []).length; i++) {
    const comp = components[i] || {};
    const itemName = String(comp.itemName || '').trim();
    if (!itemName) continue;
    const size = String(comp.size || '').trim();
    const colorGroup = String(comp.colorGroup || '').trim() || COMPONENT_COLOR_GROUP_COMMON;
    const key = itemName.toLowerCase() + '|' + size.toLowerCase() + '|' + colorGroup.toLowerCase();
    if (seen.has(key)) {
      return { itemName, size, colorGroup };
    }
    seen.add(key);
  }
  return null;
}

/**
 * Replaces a process's entire component checklist with the given list.
 * Runs under the caller's existing document lock (called from saveProcess)
 * — does not acquire its own.
 * @private
 */
function _saveProcessComponentsForProcess(processId, components) {
  let sheet;
  try {
    sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);
  } catch (e) {
    initProcessComponentsSheet();
    sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);
  }

  ensureProcessComponentsQtyColumn(sheet);
  ensureProcessComponentsSourceTypeColumn(sheet);
  ensureProcessComponentsColorGroupColumn(sheet);
  deleteRowsById(processId, sheet, 2, PROCESS_COMPONENTS_COL.PROCESS_ID);

  const rowsToWrite = (components || [])
    .map(comp => ({
      itemName: sanitizeString(comp.itemName || '', 'itemName'),
      size: sanitizeString(comp.size || '', 'size'),
      narration: sanitizeString(comp.narration || '', 'narration'),
      qtyPerUnit: validateNumber(comp.qtyPerUnit, 0.0001, 1000000) || 1,
      remarks: sanitizeString(comp.remarks || '', 'remarks'),
      sourceType: String(comp.sourceType || '').trim().toUpperCase() === COMPONENT_SOURCE_TYPES.POOL
        ? COMPONENT_SOURCE_TYPES.POOL
        : COMPONENT_SOURCE_TYPES.ITEM,
      colorGroup: sanitizeString(comp.colorGroup || '', 'colorGroup') || COMPONENT_COLOR_GROUP_COMMON
    }))
    .filter(c => c.itemName)
    .map(c => [processId, c.itemName, c.size, c.narration, c.qtyPerUnit, c.remarks, c.sourceType, c.colorGroup]);

  if (rowsToWrite.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToWrite.length, 8).setValues(rowsToWrite);
  }
}
