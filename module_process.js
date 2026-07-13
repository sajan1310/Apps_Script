/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_process.gs — PROCESS MASTER MODULE
 *
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Master list of pluggable production process types (Frame Painting, Rim
 * Assembly, Frame Fitting, Bicycle Packing, ...). WIP chaining between
 * processes is validated by module_production.js against live Warehouse
 * Pool stock (a process's Output Item Name consumed by a downstream
 * process's POOL-sourced component) — Sequence itself is display/sort
 * order only (the Processes table's default order and manual drag-reorder;
 * see reorderProcesses), not a consumption gate.
 *
 * Sheet Layout (Process Master):
 * ───────────────────────────────────────────────────────────────────────────
 * Col A (1):   Process ID (e.g. PRC-1001)
 * Col B (2):   Process Name
 * Col C (3):   Sequence (sort order in the Processes table)
 * Col D (4):   Lot Prefix (e.g. "FP")
 * Col E (5):   Is Final Stage (TRUE/FALSE)
 * Col F (6):   Active (TRUE/FALSE)
 * Col G (7):   Remarks
 * Col H (8):   Output Item Name (Warehouse Pool item this process produces per unit)
 * Col I (9):   Process Type (Reference to Process Type Master.Name; optional)
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
    Log.error('[initProcessMasterSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize Process Master sheet: ' + error.message);
  }
}

/**
 * Retrieves all processes, sorted by Sequence ascending.
 * @param {boolean} [activeOnly] - If true, excludes inactive processes.
 */
function getProcessData(activeOnly) {
  const cacheKey = activeOnly ? MASTER_DATA_CACHE_KEYS.PROCESS_ACTIVE : MASTER_DATA_CACHE_KEYS.PROCESS_ALL;
  return getCachedListResponse(cacheKey, () => {
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
      ensureProcessPrimaryColorAxisColumn(sheet);

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return buildResponse(true, []);

      const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
      const records = [];

      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const processId = String(row[PROCESS_COL.PROCESS_ID - 1] || '').trim();
        const processName = String(row[PROCESS_COL.PROCESS_NAME - 1] || '').trim();
        if (!processId || !processName) continue;

        const isActive = row[PROCESS_COL.ACTIVE - 1] === true || String(row[PROCESS_COL.ACTIVE - 1]).toUpperCase() === 'TRUE';
        if (activeOnly && !isActive) continue;

        records.push({
          processId: processId,
          processName: processName,
          sequence: Number(row[PROCESS_COL.SEQUENCE - 1]) || 0,
          lotPrefix: String(row[PROCESS_COL.LOT_PREFIX - 1] || '').trim().toUpperCase(),
          isFinalStage: row[PROCESS_COL.IS_FINAL_STAGE - 1] === true || String(row[PROCESS_COL.IS_FINAL_STAGE - 1]).toUpperCase() === 'TRUE',
          active: isActive,
          remarks: String(row[PROCESS_COL.REMARKS - 1] || '').trim(),
          outputItemName: String(row[PROCESS_COL.OUTPUT_ITEM_NAME - 1] || '').trim(),
          processType: String(row[PROCESS_COL.PROCESS_TYPE - 1] || '').trim(),
          primaryColorAxis: String(row[PROCESS_COL.PRIMARY_COLOR_AXIS - 1] || '').trim()
        });
      }

      records.sort((a, b) => a.sequence - b.sequence);

      return buildResponse(true, records);
    } catch (error) {
      Log.error('[getProcessData] Error:', error.message);
      return buildResponse(false, null, 'Failed to load process data: ' + error.message);
    }
  });
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
    Log.error('[ensureProcessOutputItemColumn] Error:', error.message);
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
    Log.error('[ensureProcessTypeColumn] Error:', error.message);
  }
}

/**
 * Backfills the "Primary Color Axis" column on Process Master sheets created
 * before the Color Axes feature existed (see computeColorGroupsForProcess).
 */
function ensureProcessPrimaryColorAxisColumn(sheet) {
  try {
    if (sheet.getLastColumn() < PROCESS_COL.PRIMARY_COLOR_AXIS) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), PROCESS_COL.PRIMARY_COLOR_AXIS - sheet.getLastColumn());
      sheet.getRange(1, PROCESS_COL.PRIMARY_COLOR_AXIS, 1, 1)
        .setValues([['Primary Color Axis']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    Log.error('[ensureProcessPrimaryColorAxisColumn] Error:', error.message);
  }
}

/**
 * @private Updates just one process's Primary Color Axis cell, without
 * touching its recipe/components/color links — used when the operator picks
 * or changes the Primary Axis directly on the Production Lot form (see
 * saveProduction in module_production.js) instead of going through the full
 * Process editor / saveProcess. Silently no-ops if the process row can't be
 * found, since this is a best-effort "remember this choice for next time"
 * side effect, not something the Production save itself should fail over.
 * @param {string} processId
 * @param {string} primaryColorAxis
 */
function _setProcessPrimaryColorAxis(processId, primaryColorAxis) {
  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_MASTER);
    ensureProcessPrimaryColorAxisColumn(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const targetId = String(processId || '').trim().toLowerCase();
    const ids = sheet.getRange(2, PROCESS_COL.PROCESS_ID, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '').trim().toLowerCase() === targetId) {
        sheet.getRange(i + 2, PROCESS_COL.PRIMARY_COLOR_AXIS).setValue(primaryColorAxis);
        invalidateListCache(MASTER_DATA_CACHE_KEYS.PROCESS_ALL, MASTER_DATA_CACHE_KEYS.PROCESS_ACTIVE);
        return;
      }
    }
  } catch (error) {
    Log.error('[_setProcessPrimaryColorAxis] Error:', error.message);
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
    Log.error('[getNextProcessId] Error:', error.message);
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
    ensureProcessPrimaryColorAxisColumn(sheet);

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

    // Linked Processes: pairs this process's colors with another process's
    // colors (see _saveProcessColorLinksForProcess / computeColorGroupsForProcess)
    // so a downstream process consuming pool items from both treats them as
    // one paired choice instead of cross-multiplying.
    let colorLinks;
    try {
      colorLinks = typeof formData.colorLinks === 'string'
        ? JSON.parse(formData.colorLinks)
        : (formData.colorLinks || []);
    } catch (e) {
      return buildResponse(false, null, 'Invalid color link data format.');
    }
    if (!Array.isArray(colorLinks)) colorLinks = [];

    // Primary Color Axis: when this process has 2+ independent color axes
    // (see computeColorGroupsForProcess), the operator picks which one's
    // checked "Colors to Produce" rows determine a lot's total quantity —
    // blank keeps the process on the legacy single-list/cross-product path.
    // Only a light shape check here (mirrors the low-friction validation
    // already used for colorLinks below) — a stale/mismatched label just
    // falls back to legacy behavior client-side rather than blocking save.
    const primaryColorAxis = sanitizeString(formData.primaryColorAxis || '', 'primaryColorAxis');

    const dupComponent = _findDuplicateComponent(components);
    if (dupComponent) {
      return buildResponse(false, null,
        `Duplicate component: "${dupComponent.itemName}"${dupComponent.size ? ' (' + dupComponent.size + ')' : ''} already exists in ${dupComponent.colorGroup === COMPONENT_COLOR_GROUP_COMMON ? 'Common Components' : 'the "' + dupComponent.colorGroup + '" color sub-group'}. Each item+size combination may only appear once per group — adjust its Qty / Unit instead of adding it twice.`);
    }

    const isEdit = !!formData.processId;
    const lastRow = sheet.getLastRow();

    // Duplicate Lot Prefix / Output Item Name checks (excluding the row being edited).
    // Output Item Name identity matters beyond display: Warehouse Pool buckets
    // (module_warehouse.js#_poolKey) are keyed by outputItemName+productTag+color
    // with NO Process ID in the key, so two ACTIVE processes sharing one Output
    // Item Name would silently merge/misattribute each other's pool stock. An
    // inactive process is exempt (it can no longer produce new lots, so it can't
    // grow that ambiguity) — deactivate the old one before reusing its name.
    if (lastRow >= 2) {
      const existing = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
      for (let i = 0; i < existing.length; i++) {
        const rowProcessId = String(existing[i][PROCESS_COL.PROCESS_ID - 1] || '').trim();
        if (isEdit && rowProcessId.toLowerCase() === String(formData.processId).trim().toLowerCase()) continue;

        const rowPrefix = String(existing[i][PROCESS_COL.LOT_PREFIX - 1] || '').trim().toUpperCase();
        if (rowPrefix === lotPrefix) {
          return buildResponse(false, null, `Lot Prefix "${lotPrefix}" is already used by another process.`);
        }

        if (active) {
          const rowActiveCell = existing[i][PROCESS_COL.ACTIVE - 1];
          const rowActive = rowActiveCell === true || String(rowActiveCell).toUpperCase() === 'TRUE';
          const rowOutputItemName = String(existing[i][PROCESS_COL.OUTPUT_ITEM_NAME - 1] || '').trim();
          if (rowActive && rowOutputItemName.toLowerCase() === outputItemName.toLowerCase()) {
            const rowProcessName = String(existing[i][PROCESS_COL.PROCESS_NAME - 1] || '').trim();
            return buildResponse(false, null,
              `Output Item Name "${outputItemName}" is already used by another active process ("${rowProcessName}"). Two active processes can't share one Warehouse Pool item — pick a different name, or deactivate "${rowProcessName}" first.`);
          }
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

      const oldRow = sheet.getRange(targetRow, 1, 1, 10).getValues()[0];
      const oldOutputItemName = String(oldRow[PROCESS_COL.OUTPUT_ITEM_NAME - 1] || '').trim();
      const oldProcessName = String(oldRow[PROCESS_COL.PROCESS_NAME - 1] || '').trim();
      const oldIsFinalStageCell = oldRow[PROCESS_COL.IS_FINAL_STAGE - 1];
      const oldIsFinalStage = oldIsFinalStageCell === true || String(oldIsFinalStageCell).toUpperCase() === 'TRUE';

      sheet.getRange(targetRow, 1, 1, 10).setValues([[
        processId, processName, sequence, lotPrefix, isFinalStage, active, remarks, outputItemName, processType, primaryColorAxis
      ]]);

      _saveProcessComponentsForProcess(processId, components);
      _saveProcessColorLinksForProcess(processId, colorLinks);

      // Renaming the Output Item Name doesn't retroactively touch Production
      // lots already saved under the old name (it's de-normalized onto each
      // lot at save time — see PRODUCTION_COL.OUTPUT_ITEM_NAME), so without
      // this, old completed lots keep crediting the Warehouse Pool under the
      // stale name while new BOMs look up the new name and see zero stock.
      const outputItemNameChanged = oldOutputItemName && oldOutputItemName.toLowerCase() !== outputItemName.toLowerCase();
      if (outputItemNameChanged) {
        _renamePoolOutputItemNameEverywhere(oldOutputItemName, outputItemName);
      }

      // recalculateWarehousePool()'s Pass 3 (Dispatch debits) and
      // _computeReadyToDispatchMap both key off each process's CURRENT
      // isFinalStage flag, not whatever it was when a lot was produced —
      // so flipping this flag changes which pool buckets are dispatchable
      // and must rebuild the pool immediately, or the Warehouse Pool/
      // Dispatch views stay stale until an unrelated Production/Dispatch
      // save happens to trigger the next rebuild.
      if (outputItemNameChanged || oldIsFinalStage !== isFinalStage) {
        if (typeof recalculateWarehousePool === 'function') {
          recalculateWarehousePool();
        }
      }

      // Contractor Rates keys its rate card on Process Name as a free string
      // (see CONTRACTOR_RATES_COL), not Process ID — without this, a renamed
      // process's rate card entry stays under the old name while
      // saveProduction resolves a lot's contractor rate by the process's
      // CURRENT name, silently zeroing every future lot's payable for it.
      if (oldProcessName && oldProcessName.toLowerCase() !== processName.toLowerCase()) {
        _renameProcessNameInContractorRates(oldProcessName, processName);
      }

      SpreadsheetApp.flush();
      invalidateListCache(MASTER_DATA_CACHE_KEYS.PROCESS_ALL, MASTER_DATA_CACHE_KEYS.PROCESS_ACTIVE);
      logAction('UPDATE', APP_CONFIG.SHEETS.PROCESS_MASTER, processId, `Process updated: ${processName}`, 'SUCCESS');
      return buildResponse(true, { processId: processId }, 'Process updated successfully.');
    }

    const newProcessId = getNextProcessId();
    sheet.appendRow([newProcessId, processName, sequence, lotPrefix, isFinalStage, active, remarks, outputItemName, processType, primaryColorAxis]);

    _saveProcessComponentsForProcess(newProcessId, components);
    _saveProcessColorLinksForProcess(newProcessId, colorLinks);

    SpreadsheetApp.flush();
    invalidateListCache(MASTER_DATA_CACHE_KEYS.PROCESS_ALL, MASTER_DATA_CACHE_KEYS.PROCESS_ACTIVE);
    logAction('CREATE', APP_CONFIG.SHEETS.PROCESS_MASTER, newProcessId, `Process created: ${processName}`, 'SUCCESS');
    return buildResponse(true, { processId: newProcessId }, 'Process created successfully.');
  } catch (error) {
    Log.error('[saveProcess] Error:', error.message);
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
 *  - Warehouse Pool Opening sheet: manually-seeded opening balances, which
 *    (per its own header comment) are never wiped/rekeyed by
 *    recalculateWarehousePool() — without this they'd stay credited to the
 *    old name forever, stranding that stock from every bucket the new name
 *    actually resolves to
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

  let openingSheet;
  try {
    openingSheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
  } catch (e) {
    openingSheet = null;
  }

  if (openingSheet) {
    const lastRow = openingSheet.getLastRow();
    if (lastRow >= 2) {
      const range = openingSheet.getRange(2, WAREHOUSE_POOL_OPENING_COL.OUTPUT_ITEM_NAME, lastRow - 1, 1);
      const values = range.getValues();
      let changed = false;
      values.forEach(row => {
        if (String(row[0] || '').trim().toLowerCase() === oldLower) {
          row[0] = newTrimmed;
          changed = true;
        }
      });
      if (changed) range.setValues(values);
    }
  }
}

/**
 * @private
 * Renames a Process across Contractor Rates' rate card, which keys rows by
 * Process Name as a free string (see CONTRACTOR_RATES_COL's header comment
 * in config.js), not by Process ID. Without this, an existing rate card
 * entry stays keyed under the old name while module_production.js's
 * saveProduction resolves a lot's contractor rate by the process's CURRENT
 * name at save time (_getContractorRate(assignedTo, process.processName)),
 * so a rename silently zeroes out every future lot's contractor payable for
 * that process until someone notices and re-enters the rate by hand.
 */
function _renameProcessNameInContractorRates(oldName, newName) {
  const tOld = String(oldName || '').trim().toLowerCase();
  const tNew = String(newName || '').trim();
  if (!tOld || !tNew || tOld === tNew.toLowerCase()) return;

  let sheet;
  try {
    sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTOR_RATES);
  } catch (e) {
    return; // Rate card sheet doesn't exist yet, nothing to rename
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, CONTRACTOR_RATES_COL.PROCESS_NAME, lastRow - 1, 1);
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

/**
 * ONE-TIME REPAIR — run manually from the Apps Script editor to fix data
 * left stale by a Process Master Output Item Name rename that happened
 * before _renamePoolOutputItemNameEverywhere() was wired into saveProcess().
 * Example: repairPoolOutputItemNameRename('Fitted Rim 14 inch', 'Fitted Rim 14 inch ED');
 */
function repairPoolOutputItemNameRename(oldName, newName) {
  _renamePoolOutputItemNameEverywhere(oldName, newName);
  const result = recalculateWarehousePool();
  Logger.log('[repairPoolOutputItemNameRename] ' + (result && result.message));
  return result;
}

/**
 * @private
 * Returns the subset of processIds (matched case-insensitively) that are
 * currently referenced elsewhere and therefore unsafe to delete:
 *  - a Production lot logged against the process (PRODUCTION_COL.PROCESS_ID)
 *  - a Product's BOM component group scoped to it (BOM_COL.PROCESS_GROUP) —
 *    deleting a Process still in a Product's recipe would silently orphan
 *    that reference (getBOMFinalStageProcessMap / getBOMProductionData both
 *    key off it) instead of surfacing the conflict to the operator.
 *  - a manually-seeded Warehouse Pool Opening balance
 *    (WAREHOUSE_POOL_OPENING_COL.PROCESS_ID) — deleting a Process that has
 *    only opening stock (no Production lot yet) would otherwise sail past
 *    the check above and leave that row's Process ID dangling.
 * Shared by deleteProcess and deleteProcessesBulk so both guard against the
 * same set of references with one read of each sheet.
 * @param {string[]} processIds
 * @returns {Set<string>} lower-cased process IDs found in use
 */
function _getProcessIdsInUse(processIds) {
  const requested = new Set((processIds || []).map(id => String(id || '').trim().toLowerCase()).filter(Boolean));
  const inUse = new Set();
  if (requested.size === 0) return inUse;

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
      refs.forEach(row => {
        const id = String(row[0] || '').trim().toLowerCase();
        if (requested.has(id)) inUse.add(id);
      });
    }
  }

  let bomSheet;
  try {
    bomSheet = getSheet(APP_CONFIG.SHEETS.BOM);
  } catch (e) {
    bomSheet = null;
  }
  if (bomSheet) {
    const bLastRow = bomSheet.getLastRow();
    if (bLastRow >= 2) {
      const refs = bomSheet.getRange(2, BOM_COL.PROCESS_GROUP, bLastRow - 1, 1).getValues();
      refs.forEach(row => {
        const id = String(row[0] || '').trim().toLowerCase();
        if (requested.has(id)) inUse.add(id);
      });
    }
  }

  let openingSheet;
  try {
    openingSheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
  } catch (e) {
    openingSheet = null;
  }
  if (openingSheet) {
    const oLastRow = openingSheet.getLastRow();
    if (oLastRow >= 2) {
      const refs = openingSheet.getRange(2, WAREHOUSE_POOL_OPENING_COL.PROCESS_ID, oLastRow - 1, 1).getValues();
      refs.forEach(row => {
        const id = String(row[0] || '').trim().toLowerCase();
        if (requested.has(id)) inUse.add(id);
      });
    }
  }

  return inUse;
}

/**
 * Deletes a process. Blocked if any Production lot or Product BOM already
 * references it (see _getProcessIdsInUse).
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

    if (_getProcessIdsInUse([idClean]).has(idClean.toLowerCase())) {
      return buildResponse(false, null, `Cannot delete process: "${idClean}" is already referenced by Production lots, a Product's recipe (BOM), or a Warehouse Pool opening balance.`);
    }

    const rowsDeleted = deleteRowsById(idClean, sheet, 2, PROCESS_COL.PROCESS_ID);

    // Also remove this process's component checklist rows
    try {
      const compSheet = getSheet(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);
      deleteRowsById(idClean, compSheet, 2, PROCESS_COMPONENTS_COL.PROCESS_ID);
    } catch (e) {
      // Process Components sheet doesn't exist yet, nothing to clean up
    }

    // A color link is just a preference, not a real-quantity reference like
    // Production/BOM/Opening-balance rows — cascade-delete instead of
    // blocking, so a deleted process never leaves a dangling link row.
    try {
      const linksSheet = getSheet(APP_CONFIG.SHEETS.PROCESS_COLOR_LINKS);
      deleteRowsById(idClean, linksSheet, 2, PROCESS_COLOR_LINKS_COL.PROCESS_A_ID);
      deleteRowsById(idClean, linksSheet, 2, PROCESS_COLOR_LINKS_COL.PROCESS_B_ID);
    } catch (e) {
      // Process Color Links sheet doesn't exist yet, nothing to clean up
    }

    SpreadsheetApp.flush();
    invalidateListCache(MASTER_DATA_CACHE_KEYS.PROCESS_ALL, MASTER_DATA_CACHE_KEYS.PROCESS_ACTIVE);

    const msg = `Process "${idClean}" deleted (${rowsDeleted} row(s) removed).`;
    logAction('DELETE', APP_CONFIG.SHEETS.PROCESS_MASTER, idClean, msg, 'SUCCESS');

    return buildResponse(true, null, msg);
  } catch (error) {
    Log.error('[deleteProcess] Error:', error.message);
    logAction('ERROR', 'deleteProcess', processId, error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete process: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes multiple processes at once. Processes already referenced by
 * Production lots or a Product's BOM are skipped (mirrors deleteProcess's
 * single-item check — see _getProcessIdsInUse).
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

    const inUseLower = _getProcessIdsInUse(requested);
    const inUseSet = new Set(requested.filter(id => inUseLower.has(id.toLowerCase())));

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

      try {
        const linksSheet = getSheet(APP_CONFIG.SHEETS.PROCESS_COLOR_LINKS);
        _rewriteWithoutMatchingRowsBulk(linksSheet, 2, PROCESS_COLOR_LINKS_COL.PROCESS_A_ID, targetSet);
        _rewriteWithoutMatchingRowsBulk(linksSheet, 2, PROCESS_COLOR_LINKS_COL.PROCESS_B_ID, targetSet);
      } catch (e) {
        // Process Color Links sheet doesn't exist yet, nothing to clean up
      }

      SpreadsheetApp.flush();
      invalidateListCache(MASTER_DATA_CACHE_KEYS.PROCESS_ALL, MASTER_DATA_CACHE_KEYS.PROCESS_ACTIVE);
    }

    let msg = `Deleted ${toDelete.length} process(es) (${rowsDeleted} rows removed).`;
    if (inUseSet.size > 0) {
      msg += ` Skipped ${inUseSet.size} process(es) still in use by Production, a Product's BOM, or a Warehouse Pool opening balance: ${Array.from(inUseSet).join(', ')}.`;
    }
    logAction('BULK_DELETE', APP_CONFIG.SHEETS.PROCESS_MASTER, 'multiple', msg, 'SUCCESS');

    return buildResponse(true, { skipped: Array.from(inUseSet) }, msg);
  } catch (error) {
    Log.error('[deleteProcessesBulk] Error:', error.message);
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
    invalidateListCache(MASTER_DATA_CACHE_KEYS.PROCESS_ALL, MASTER_DATA_CACHE_KEYS.PROCESS_ACTIVE);

    const msg = `Reordered ${order.length} process(es).`;
    logAction('UPDATE', APP_CONFIG.SHEETS.PROCESS_MASTER, 'multiple', msg, 'SUCCESS');

    return buildResponse(true, null, msg);
  } catch (error) {
    Log.error('[reorderProcesses] Error:', error.message);
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
    Log.error('[initProcessComponentsSheet] Error:', error.message);
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
    Log.error('[ensureProcessComponentsQtyColumn] Error:', error.message);
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
    Log.error('[ensureProcessComponentsSourceTypeColumn] Error:', error.message);
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
    Log.error('[ensureProcessComponentsColorGroupColumn] Error:', error.message);
  }
}

/**
 * Backfills the "Color Axis" column on Process Components sheets created
 * before the Color Axes feature existed. Existing rows default to blank
 * (not manually assigned to an axis) — see PROCESS_COMPONENTS_COL.COLOR_AXIS
 * and computeColorGroupsForProcess.
 */
function ensureProcessComponentsColorAxisColumn(sheet) {
  try {
    if (sheet.getLastColumn() < PROCESS_COMPONENTS_COL.COLOR_AXIS) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), PROCESS_COMPONENTS_COL.COLOR_AXIS - sheet.getLastColumn());
      sheet.getRange(1, PROCESS_COMPONENTS_COL.COLOR_AXIS, 1, 1)
        .setValues([['Color Axis']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    Log.error('[ensureProcessComponentsColorAxisColumn] Error:', error.message);
  }
}

/**
 * Optional Unit column (see PROCESS_COMPONENTS_COL.UNIT) — blank on every
 * pre-existing row, preserving today's "qty is already in the item's Base
 * Unit" behavior exactly. Only a row where the user explicitly picks a Unit
 * going forward triggers real conversion in module_stock.js.
 */
function ensureProcessComponentsUnitColumn(sheet) {
  try {
    if (sheet.getLastColumn() < PROCESS_COMPONENTS_COL.UNIT) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), PROCESS_COMPONENTS_COL.UNIT - sheet.getLastColumn());
      sheet.getRange(1, PROCESS_COMPONENTS_COL.UNIT, 1, 1)
        .setValues([['Unit']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    Log.error('[ensureProcessComponentsUnitColumn] Error:', error.message);
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
    ensureProcessComponentsColorAxisColumn(sheet);
    ensureProcessComponentsUnitColumn(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, []);

    const targetId = String(processId || '').trim().toLowerCase();
    const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

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
        colorGroup: String(row[PROCESS_COMPONENTS_COL.COLOR_GROUP - 1] || '').trim() || COMPONENT_COLOR_GROUP_COMMON,
        colorAxis: String(row[PROCESS_COMPONENTS_COL.COLOR_AXIS - 1] || '').trim(),
        // Blank = "already in the item's Base Unit" (preserves pre-existing
        // recipes' behavior exactly) — see PROCESS_COMPONENTS_COL.UNIT.
        unit: String(row[PROCESS_COMPONENTS_COL.UNIT - 1] || '').trim()
      }))
      .filter(c => c.itemName && (!targetId || c.processId.toLowerCase() === targetId));

    return buildResponse(true, components);
  } catch (error) {
    Log.error('[getProcessComponentsData] Error:', error.message);
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
    const colorLinks = _getAllProcessColorLinks();
    return buildResponse(true, computeColorGroupsForProcess(components, poolRows, colorLinks));
  } catch (error) {
    Log.error('[getProcessColorGroups] Error:', error.message);
    return buildResponse(false, null, 'Failed to load process color groups: ' + error.message);
  }
}

/**
 * @private Looks up one process's own record (for its primaryColorAxis
 * field) without the caller having to fetch/filter the full process list.
 */
function _getProcessRecordById(processId) {
  const targetId = String(processId || '').trim().toLowerCase();
  if (!targetId) return null;
  const resp = getProcessData(false);
  const records = (resp && resp.data) || [];
  return records.find(p => p.processId.toLowerCase() === targetId) || null;
}

/**
 * Shared core of getProcessColorGroups — pulled out so the bulk variant
 * (getAllProcessColorGroups) can read the Process Components, Warehouse
 * Pool, and Process Color Links sheets ONCE for every process instead of
 * once per process.
 *
 * Branches on whether 2+ independent color axes actually exist (via
 * computeColorAxesForProcess) — NOT on whether primaryColorAxis has been
 * saved onto the process yet. This must match Script.html's
 * renderGroupedColorChecklist, which renders one independent checkbox
 * group per axis (individual colors, e.g. "Blue", "Blue-White") the moment
 * 2+ axes are detected, regardless of whether a Primary Axis has been
 * configured — the picker for choosing Primary only appears once that split
 * view is already showing. Branching on primaryColorAxis instead (the
 * original design) meant any process with 2+ auto-detected pool axes but no
 * saved Primary Axis yet would render individual-color checkboxes on the
 * form but validate against the OLD cross-multiplied composite list (e.g.
 * "Blue / BCP / Blue-White") here — rejecting every real color on every
 * submission until someone visited the Process editor to save a Primary
 * Axis (symptom: "Color X is not a configured color sub-group" firing on
 * every save for a freshly multi-axis process, unfixable by refreshing).
 *
 * A process with fewer than 2 axes (0 or 1) still gets the exact original
 * _legacyColorGroupList behavior — colors explicitly tagged on recipe rows
 * (including tag-only colors with no colorAxis label, which
 * computeColorAxesForProcess intentionally excludes) plus any single pool
 * axis's own colors, unchanged.
 * @param {Array} components Process Components rows already filtered to one process.
 * @param {Array} poolRows Full Warehouse Pool rows (shared across all processes).
 * @param {Array} colorLinks Full Process Color Links rows (shared across all processes) — see _getAllProcessColorLinks.
 */
function computeColorGroupsForProcess(components, poolRows, colorLinks) {
  const axes = computeColorAxesForProcess(components, poolRows, colorLinks);
  if (axes.length < 2) {
    return _legacyColorGroupList(components, poolRows, colorLinks);
  }

  const colors = new Set();
  axes.forEach(axis => {
    axis.colors.forEach(c => colors.add(c));
  });
  return Array.from(colors).sort((a, b) => a.localeCompare(b));
}

/**
 * @private Adds `value` into `map` (used as a case-insensitive Set) keyed by
 * its trimmed lowercase form, keeping the FIRST-seen casing as the value.
 * Two user-typed color/axis names that are the same real thing except for
 * casing (e.g. "Red" then "red", typed on different recipe rows or in
 * different production lots) must collapse to ONE entry — a plain Set built
 * from raw strings treats them as two, producing duplicate/phantom color
 * checkboxes and spuriously-multi-color axes. Read back with
 * `Array.from(map.values())`.
 */
function _addUniqueCaseInsensitive(map, value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return;
  const key = raw.toLowerCase();
  if (!map.has(key)) map.set(key, raw);
}

/**
 * @private The original, unmodified computeColorGroupsForProcess algorithm
 * — see computeColorGroupsForProcess for why this stays byte-for-byte the
 * pre-Color-Axes behavior.
 */
function _legacyColorGroupList(components, poolRows, colorLinks) {
  const colors = new Map();
  components.forEach(c => {
    if (c.colorGroup && c.colorGroup !== COMPONENT_COLOR_GROUP_COMMON) _addUniqueCaseInsensitive(colors, c.colorGroup);
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
    const processIdByItem = new Map();
    poolRows.forEach(r => {
      const key = r.outputItemName.toLowerCase();
      if (!r.color || !poolItemNames.has(key)) return;
      if (!colorsByItem.has(key)) colorsByItem.set(key, new Map());
      _addUniqueCaseInsensitive(colorsByItem.get(key), r.color);
      if (!processIdByItem.has(key) && r.processId) processIdByItem.set(key, r.processId);
    });

    // Items sharing the EXACT same color set (e.g. two recipe rows both
    // sourced from the same upstream "Painted Frame") are one color axis.
    // Items with genuinely different color sets (e.g. a Fitted Rim colored
    // BCP/Black assembled alongside a Frame colored Blue-White/Orange-White)
    // are independent axes — a real output unit needs one color from EACH
    // axis at once, so the producible variants are the cross product of all
    // axes, not a flat union (which would falsely offer "BCP" alone as a
    // complete output color with no frame color attached) — UNLESS an
    // explicit Process Color Link (see _mergeLinkedAxes) declares two axes
    // correlated, in which case they're paired instead of cross-multiplied.
    const axesBySignature = new Map();
    colorsByItem.forEach((itemColors, itemKey) => {
      if (itemColors.size <= 1) return;
      const sorted = Array.from(itemColors.values()).sort((a, b) => a.localeCompare(b));
      const signature = sorted.map(c => c.toLowerCase()).join('|');
      if (!axesBySignature.has(signature)) {
        axesBySignature.set(signature, { colors: sorted, processIds: new Set() });
      }
      const pid = processIdByItem.get(itemKey);
      if (pid) axesBySignature.get(signature).processIds.add(pid);
    });

    let axes = Array.from(axesBySignature.values());
    if (axes.length > 1 && colorLinks && colorLinks.length > 0) {
      axes = _mergeLinkedAxes(axes, colorLinks);
    }

    if (axes.length === 1) {
      axes[0].colors.forEach(color => _addUniqueCaseInsensitive(colors, color));
    } else if (axes.length > 1) {
      let combos = [''];
      axes.forEach(axis => {
        const next = [];
        combos.forEach(prefix => {
          axis.colors.forEach(color => next.push(prefix ? `${prefix}${COLOR_COMBO_DELIMITER}${color}` : color));
        });
        combos = next;
      });
      combos.forEach(combo => _addUniqueCaseInsensitive(colors, combo));
    }
  }

  return Array.from(colors.values()).sort((a, b) => a.localeCompare(b));
}

/**
 * Builds the independent "Color Axes" breakdown for a process — one entry
 * per contributing axis, NEVER cross-multiplied into composite strings.
 * Two independent sources feed this, exactly like _legacyColorGroupList's
 * two sources, just kept as separate groups instead of being unioned/
 * cross-multiplied into one flat list:
 *   1. Warehouse Pool axes — auto-detected from live pool color history of
 *      this recipe's POOL-sourced components, same detection as
 *      _legacyColorGroupList, just NOT cross-multiplied. Labeled by the
 *      pool item name(s) that drive them (e.g. "Painted Frame").
 *   2. Explicitly-tagged axes — recipe rows carrying a literal Color
 *      Master name in COLOR_GROUP (not the pool-axis mechanism above) AND
 *      an explicit COLOR_AXIS label (e.g. a Mudguard component tagged
 *      colorGroup "Red" with colorAxis "Mudguard Color"). Rows with a
 *      colorGroup but no colorAxis label are left out of this breakdown
 *      entirely (only the legacy flat list sees them) — see
 *      PROCESS_COMPONENTS_COL.COLOR_AXIS.
 * @param {Array} components
 * @param {Array} poolRows
 * @param {Array} colorLinks
 * @returns {Array<{key: string, label: string, colors: string[], source: 'pool'|'tag'}>}
 */
function computeColorAxesForProcess(components, poolRows, colorLinks) {
  const axes = [];

  const poolItemNames = new Set(
    components.filter(c => c.sourceType === COMPONENT_SOURCE_TYPES.POOL).map(c => c.itemName.toLowerCase())
  );
  if (poolItemNames.size > 0) {
    const colorsByItem = new Map();
    const processIdByItem = new Map();
    const itemNameByKey = new Map();
    components.forEach(c => {
      if (c.sourceType !== COMPONENT_SOURCE_TYPES.POOL) return;
      const key = c.itemName.toLowerCase();
      if (!itemNameByKey.has(key)) itemNameByKey.set(key, c.itemName);
    });
    poolRows.forEach(r => {
      const key = r.outputItemName.toLowerCase();
      if (!r.color || !poolItemNames.has(key)) return;
      if (!colorsByItem.has(key)) colorsByItem.set(key, new Map());
      _addUniqueCaseInsensitive(colorsByItem.get(key), r.color);
      if (!processIdByItem.has(key) && r.processId) processIdByItem.set(key, r.processId);
    });

    const axesBySignature = new Map();
    colorsByItem.forEach((itemColors, itemKey) => {
      if (itemColors.size <= 1) return;
      const sorted = Array.from(itemColors.values()).sort((a, b) => a.localeCompare(b));
      const signature = sorted.map(c => c.toLowerCase()).join('|');
      if (!axesBySignature.has(signature)) {
        axesBySignature.set(signature, { colors: sorted, processIds: new Set(), itemNames: new Set() });
      }
      const entry = axesBySignature.get(signature);
      const pid = processIdByItem.get(itemKey);
      if (pid) entry.processIds.add(pid);
      const itemName = itemNameByKey.get(itemKey);
      if (itemName) entry.itemNames.add(itemName);
    });

    let poolAxes = Array.from(axesBySignature.values());
    if (poolAxes.length > 1 && colorLinks && colorLinks.length > 0) {
      poolAxes = _mergeLinkedAxes(poolAxes, colorLinks);
    }

    poolAxes.forEach((axis, idx) => {
      const itemNames = axis.itemNames ? Array.from(axis.itemNames).sort((a, b) => a.localeCompare(b)) : [];
      const label = itemNames.length > 0 ? itemNames.join(', ') : `Color Group ${idx + 1}`;
      axes.push({ key: 'pool:' + label.toLowerCase(), label, colors: axis.colors, source: 'pool' });
    });
  }

  // Keyed by the axis label's lowercase form so "Mudguard Color" and
  // "mudguard color" (typed on different recipe rows) collapse into one
  // real axis instead of two — same casing-drift hazard as the colors
  // themselves, just one level up (see _addUniqueCaseInsensitive).
  const rawTagGroups = new Map();
  components.forEach(c => {
    if (!c.colorGroup || c.colorGroup === COMPONENT_COLOR_GROUP_COMMON) return;
    const axisLabel = String(c.colorAxis || '').trim();
    if (!axisLabel) return;
    const axisKey = axisLabel.toLowerCase();
    if (!rawTagGroups.has(axisKey)) rawTagGroups.set(axisKey, { label: axisLabel, colors: new Map() });
    _addUniqueCaseInsensitive(rawTagGroups.get(axisKey).colors, c.colorGroup);
  });
  rawTagGroups.forEach(({ label, colors: colorMap }) => {
    axes.push({
      key: 'tag:' + label.toLowerCase(),
      label,
      colors: Array.from(colorMap.values()).sort((a, b) => a.localeCompare(b)),
      source: 'tag'
    });
  });

  return axes;
}

/**
 * Returns the full Color Axes breakdown for one process — used by the
 * Production Lot form's split checklist UI (see computeColorAxesForProcess)
 * and by the Process editor's Primary Axis picker. Unlike getProcessColorGroups,
 * this always returns the axis breakdown regardless of whether the process
 * has opted into a primaryColorAxis yet, so the Process editor can populate
 * the picker's options before the operator has chosen one.
 * @param {string} processId
 */
function getProcessColorAxes(processId) {
  try {
    const componentsResp = getProcessComponentsData(processId);
    const components = (componentsResp && componentsResp.data) || [];
    const poolRows = typeof getWarehousePoolData === 'function'
      ? ((getWarehousePoolData() || {}).data || [])
      : [];
    const colorLinks = _getAllProcessColorLinks();
    const process = _getProcessRecordById(processId);
    const primaryColorAxis = String((process && process.primaryColorAxis) || '').trim();

    const axes = computeColorAxesForProcess(components, poolRows, colorLinks);
    const primaryAxisKey = primaryColorAxis
      ? (axes.find(a => a.label.toLowerCase() === primaryColorAxis.toLowerCase()) || {}).key || ''
      : '';

    return buildResponse(true, { axes, primaryColorAxis, primaryAxisKey });
  } catch (error) {
    Log.error('[getProcessColorAxes] Error:', error.message);
    return buildResponse(false, null, 'Failed to load process color axes: ' + error.message);
  }
}

/**
 * @private
 * Merges axes belonging to explicitly-linked processes (see Process Color
 * Links, config.js PROCESS_COLOR_LINKS_COL) into single paired axes instead
 * of leaving them to be cross-multiplied. Only axes contributed by exactly
 * one process are link-eligible (an axis already collapsed from 2+
 * processes by identical-signature matching has no single process to key
 * the graph on, and is left untouched). Chained links (B-D, D-A) group 3+
 * processes transitively via BFS with no dedicated N-way data structure.
 * @param {Array<{colors: string[], processIds: Set<string>}>} axes
 * @param {Array<{processAId: string, colorA: string, processBId: string, colorB: string}>} colorLinks
 * @returns {Array<{colors: string[]}>}
 */
function _mergeLinkedAxes(axes, colorLinks) {
  const adjacency = new Map(); // processId -> [{ otherProcessId, map: Map(myColorLower -> theirColor) }]
  function addEdge(pFrom, colorFrom, pTo, colorTo) {
    if (!pFrom || !pTo) return;
    if (!adjacency.has(pFrom)) adjacency.set(pFrom, []);
    let entry = adjacency.get(pFrom).find(e => e.otherProcessId === pTo);
    if (!entry) {
      entry = { otherProcessId: pTo, map: new Map() };
      adjacency.get(pFrom).push(entry);
    }
    entry.map.set(String(colorFrom || '').trim().toLowerCase(), colorTo);
  }
  colorLinks.forEach(link => {
    addEdge(link.processAId, link.colorA, link.processBId, link.colorB);
    addEdge(link.processBId, link.colorB, link.processAId, link.colorA);
  });

  // Only axes contributed by exactly one process can be placed on the graph.
  const axisIndexByProcessId = new Map();
  axes.forEach((axis, idx) => {
    if (axis.processIds.size === 1) {
      axisIndexByProcessId.set(Array.from(axis.processIds)[0], idx);
    }
  });

  const visited = new Set();
  const mergedAxes = [];
  const usedAxisIdx = new Set();

  axisIndexByProcessId.forEach((idx, pid) => {
    if (visited.has(pid)) return;

    const queue = [pid];
    visited.add(pid);
    const componentProcessIds = [pid];
    while (queue.length > 0) {
      const cur = queue.shift();
      (adjacency.get(cur) || []).forEach(edge => {
        if (axisIndexByProcessId.has(edge.otherProcessId) && !visited.has(edge.otherProcessId)) {
          visited.add(edge.otherProcessId);
          queue.push(edge.otherProcessId);
          componentProcessIds.push(edge.otherProcessId);
        }
      });
    }

    if (componentProcessIds.length <= 1) return; // no link partner present in this recipe — leave axis as-is

    let anchorPid = componentProcessIds[0];
    componentProcessIds.forEach(p => {
      if (axes[axisIndexByProcessId.get(p)].colors.length > axes[axisIndexByProcessId.get(anchorPid)].colors.length) {
        anchorPid = p;
      }
    });
    const otherPids = componentProcessIds.filter(p => p !== anchorPid);

    const mergedColors = [];
    axes[axisIndexByProcessId.get(anchorPid)].colors.forEach(anchorColor => {
      const parts = [anchorColor];
      const unresolved = otherPids.some(otherPid => {
        const resolved = _resolveLinkedColor(anchorPid, anchorColor, otherPid, adjacency);
        if (resolved == null) return true;
        parts.push(resolved);
        return false;
      });
      if (unresolved) {
        Log.warn('[computeColorGroupsForProcess] Linked processes ' + componentProcessIds.join(', ') +
          ' have no full color mapping for "' + anchorColor + '" — skipping that combination.');
        return;
      }
      mergedColors.push(parts.join(COLOR_COMBO_DELIMITER));
    });

    if (mergedColors.length > 0) {
      mergedAxes.push({ colors: mergedColors });
    }
    componentProcessIds.forEach(p => usedAxisIdx.add(axisIndexByProcessId.get(p)));
  });

  const result = axes.filter((axis, idx) => !usedAxisIdx.has(idx));
  return result.concat(mergedAxes);
}

/**
 * @private
 * Resolves the color on `toPid` that corresponds to `fromColor` on `fromPid`,
 * composing multiple Process Color Link edges when the two processes aren't
 * directly linked but are connected through an intermediate process (e.g.
 * B-D-A). Returns null if no chain of explicit mappings connects them for
 * this specific color value.
 */
function _resolveLinkedColor(fromPid, fromColor, toPid, adjacency) {
  const visited = new Set([fromPid]);
  const queue = [{ pid: fromPid, color: fromColor }];
  while (queue.length > 0) {
    const { pid, color } = queue.shift();
    if (pid === toPid) return color;
    (adjacency.get(pid) || []).forEach(edge => {
      if (visited.has(edge.otherProcessId)) return;
      const nextColor = edge.map.get(String(color || '').trim().toLowerCase());
      if (nextColor === undefined) return;
      visited.add(edge.otherProcessId);
      queue.push({ pid: edge.otherProcessId, color: nextColor });
    });
  }
  return null;
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
    const colorLinks = _getAllProcessColorLinks();

    const result = {};
    processes.forEach(p => {
      const components = componentsByProcess.get(p.processId) || [];
      result[p.processId] = computeColorGroupsForProcess(components, poolRows, colorLinks);
    });

    return buildResponse(true, result);
  } catch (error) {
    Log.error('[getAllProcessColorGroups] Error:', error.message);
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
  ensureProcessComponentsColorAxisColumn(sheet);
  ensureProcessComponentsUnitColumn(sheet);
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
      colorGroup: sanitizeString(comp.colorGroup || '', 'colorGroup') || COMPONENT_COLOR_GROUP_COMMON,
      colorAxis: sanitizeString(comp.colorAxis || '', 'colorAxis'),
      // Blank = "already in the item's Base Unit" — see PROCESS_COMPONENTS_COL.UNIT.
      unit: sanitizeString(comp.unit || '', 'unit')
    }))
    .filter(c => c.itemName)
    .map(c => [processId, c.itemName, c.size, c.narration, c.qtyPerUnit, c.remarks, c.sourceType, c.colorGroup, c.colorAxis, c.unit]);

  if (rowsToWrite.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToWrite.length, 10).setValues(rowsToWrite);
  }
}

/**
 * Rewrites Process Components rows referencing a renamed/merged Items
 * Master item, so process recipes stay attributed to the item's current
 * name + size. Only touches SOURCE_TYPE 'ITEM' rows — a 'POOL' row's
 * itemName is a Process's Output Item Name (a different identity space,
 * handled by _renamePoolOutputItemNameEverywhere), not an Items Master
 * item, so it must never be rewritten here even on a name+size coincidence.
 *
 * Called from module_items.js's _propagateItemIdentityChange, same as
 * backfillBillItemRefs/backfillPOItemRefs/backfillBOMItemRefs — without
 * this, a renamed/merged item leaves any Process's ITEM-sourced component
 * pointing at a name+size Items Master no longer has any record of, even
 * though _getItemKeysInUse blocks deleting that same item for exactly this
 * kind of orphaning.
 *
 * Runs under the caller's existing document lock — does not acquire its own.
 *
 * @param {string} oldName
 * @param {string} oldSize
 * @param {string} newName
 * @param {string} newSize
 */
function backfillProcessComponentItemRefs(oldName, oldSize, newName, newSize) {
  let sheet;
  try {
    sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_COMPONENTS);
  } catch (e) {
    return;
  }

  const startRow = 2;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  const numRows = lastRow - startRow + 1;
  const range = sheet.getRange(startRow, PROCESS_COMPONENTS_COL.ITEM_NAME, numRows, PROCESS_COMPONENTS_COL.SOURCE_TYPE - PROCESS_COMPONENTS_COL.ITEM_NAME + 1);
  const values = range.getValues();
  const sourceTypeIdx = PROCESS_COMPONENTS_COL.SOURCE_TYPE - PROCESS_COMPONENTS_COL.ITEM_NAME;

  const tOldName = String(oldName || '').trim().toLowerCase();
  const tOldSize = String(oldSize || '').trim().toLowerCase();

  let changed = false;
  for (let i = 0; i < values.length; i++) {
    const sourceType = String(values[i][sourceTypeIdx] || '').trim().toUpperCase();
    if (sourceType === COMPONENT_SOURCE_TYPES.POOL) continue;

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

/**
 * Initializes the Process Color Links sheet with correct headers.
 */
function initProcessColorLinksSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.PROCESS_COLOR_LINKS);
  if (!sheet) {
    sheet = ss.insertSheet(APP_CONFIG.SHEETS.PROCESS_COLOR_LINKS);
  }

  const headers = ['Process A ID', 'Color A', 'Process B ID', 'Color B'];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#f3f3f3');

  SpreadsheetApp.flush();
}

/**
 * @private
 * Full read of every Process Color Links row. Read once and shared across a
 * single computation (see computeColorGroupsForProcess) rather than
 * re-reading the sheet per process — same discipline as poolRows in
 * getAllProcessColorGroups.
 * @returns {Array<{processAId: string, colorA: string, processBId: string, colorB: string}>}
 */
function _getAllProcessColorLinks() {
  let sheet;
  try {
    sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_COLOR_LINKS);
  } catch (e) {
    return [];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  return data
    .map(row => ({
      processAId: String(row[PROCESS_COLOR_LINKS_COL.PROCESS_A_ID - 1] || '').trim(),
      colorA: String(row[PROCESS_COLOR_LINKS_COL.COLOR_A - 1] || '').trim(),
      processBId: String(row[PROCESS_COLOR_LINKS_COL.PROCESS_B_ID - 1] || '').trim(),
      colorB: String(row[PROCESS_COLOR_LINKS_COL.COLOR_B - 1] || '').trim()
    }))
    .filter(r => r.processAId && r.colorA && r.processBId && r.colorB);
}

/**
 * Public read endpoint: every color-link pair touching one process,
 * normalized so the caller always sees {otherProcessId, otherProcessName,
 * myColor, theirColor} regardless of which column (A or B) this process was
 * originally stored under. Powers the Process edit modal's "Linked
 * Processes" section.
 * @param {string} processId
 */
function getProcessColorLinksData(processId) {
  try {
    const targetId = String(processId || '').trim().toLowerCase();
    if (!targetId) return buildResponse(true, []);

    const allProcesses = typeof _getAllProcessRecords === 'function' ? _getAllProcessRecords() : [];
    const processNameById = {};
    allProcesses.forEach(p => { processNameById[p.processId.toLowerCase()] = p.processName; });

    const links = _getAllProcessColorLinks();
    const records = [];
    links.forEach(link => {
      if (link.processAId.toLowerCase() === targetId) {
        records.push({
          otherProcessId: link.processBId,
          otherProcessName: processNameById[link.processBId.toLowerCase()] || link.processBId,
          myColor: link.colorA,
          theirColor: link.colorB
        });
      } else if (link.processBId.toLowerCase() === targetId) {
        records.push({
          otherProcessId: link.processAId,
          otherProcessName: processNameById[link.processAId.toLowerCase()] || link.processAId,
          myColor: link.colorB,
          theirColor: link.colorA
        });
      }
    });

    return buildResponse(true, records);
  } catch (error) {
    Log.error('[getProcessColorLinksData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load process color links: ' + error.message);
  }
}

/**
 * @private
 * Replaces every color-link row touching one process with the given list.
 * Runs under the caller's existing document lock (called from saveProcess)
 * — does not acquire its own. Mirrors _saveProcessComponentsForProcess's
 * delete-then-append replace pattern. `links` is the flat
 * {otherProcessId, myColor, theirColor} shape getProcessColorLinksData
 * returns; this process is always re-written as Process A on save regardless
 * of which side originally created the row.
 * @param {string} processId
 * @param {Array<{otherProcessId: string, myColor: string, theirColor: string}>} links
 */
function _saveProcessColorLinksForProcess(processId, links) {
  let sheet;
  try {
    sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_COLOR_LINKS);
  } catch (e) {
    initProcessColorLinksSheet();
    sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_COLOR_LINKS);
  }

  deleteRowsById(processId, sheet, 2, PROCESS_COLOR_LINKS_COL.PROCESS_A_ID);
  deleteRowsById(processId, sheet, 2, PROCESS_COLOR_LINKS_COL.PROCESS_B_ID);

  const allProcesses = typeof _getAllProcessRecords === 'function' ? _getAllProcessRecords() : [];
  const validProcessIds = new Set(allProcesses.map(p => p.processId.toLowerCase()));
  const selfLower = String(processId || '').trim().toLowerCase();

  const rowsToWrite = (links || [])
    .map(link => ({
      otherProcessId: sanitizeString(link.otherProcessId || '', 'otherProcessId'),
      myColor: sanitizeString(link.myColor || '', 'myColor'),
      theirColor: sanitizeString(link.theirColor || '', 'theirColor')
    }))
    .filter(link =>
      link.otherProcessId && link.myColor && link.theirColor &&
      link.otherProcessId.toLowerCase() !== selfLower &&
      validProcessIds.has(link.otherProcessId.toLowerCase())
    )
    .map(link => [processId, link.myColor, link.otherProcessId, link.theirColor]);

  if (rowsToWrite.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToWrite.length, 4).setValues(rowsToWrite);
  }
}
