/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_production.gs — PRODUCTION LOT MODULE
 * 
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Complete CRUD operations for Production Lots including:
 * - Retrieval of production logs (getProductionData)
 * - Recording new production runs and updates (saveProduction)
 * - Deleting production logs (deleteProduction)
 * 
 * Sheet Layout (Production):
 * ───────────────────────────────────────────────────────────────────────────
 * Col A (1):   Date (DD/MM/YYYY)
 * Col B (2):   Product_ID (OPTIONAL — only set as a tag on final-stage lots)
 * Col C (3):   Product Name (OPTIONAL — de-normalized copy of the tag)
 * Col D (4):   Quantity
 * Col E (5):   Assigned by
 * Col F (6):   Assigned to (the contractor — required)
 * Col G (7):   Status
 * Col H (8):   Remarks
 *
 * Logging new production no longer requires a Product ID — only a Process
 * ID, the components consumed (drawn from that process's recipe, qty
 * editable), production quantity, and contractor assignment. Output is
 * credited to the Warehouse Pool under the process's own Output Item Name
 * (see module_warehouse.js). Only for final-stage (packing) processes is an
 * optional Product tag available, used purely so Dispatch can find the
 * finished, fully-packed output.
 * ═══════════════════════════════════════════════════════════════════════════
 */

const PRODUCTION_LOCK_TIMEOUT_MS = 15000;

/**
 * Initializes the Production sheet with correct headers.
 * Run once manually or triggered automatically when sheet is missing.
 */
function initProductionSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.PRODUCTION);
    if (!sheet) {
      sheet = ss.insertSheet(APP_CONFIG.SHEETS.PRODUCTION);
    }
    
    const headers = [
      'Date',
      'Product_ID (Tag)',
      'Product Name (Tag)',
      'Quantity',
      'Assigned by',
      'Assigned to',
      'Status',
      'Remarks',
      'Custom Components (JSON)',
      'Sheet Remarks',
      'Process ID',
      'Lot Number',
      'Consumed From Process Qty',
      'Contractor Rate',
      'Contractor Payable',
      'Output Item Name',
      'Components Consumed (JSON)',
      'Color',
      'Color Breakdown (JSON)'
    ];

    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Production sheet initialized successfully.');
  } catch (error) {
    Log.error('[initProductionSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize Production sheet: ' + error.message);
  }
}

/**
 * Backfills the "Custom Components (JSON)" and "Sheet Remarks" columns on
 * Production sheets created before this feature existed, so legacy rows
 * don't throw when read/written.
 */
function ensureProductionExtraColumns(sheet) {
  try {
    if (sheet.getLastColumn() < PRODUCTION_COL.SHEET_REMARKS) {
      const startCol = sheet.getLastColumn() + 1;
      sheet.insertColumnsAfter(sheet.getLastColumn(), PRODUCTION_COL.SHEET_REMARKS - sheet.getLastColumn());
      sheet.getRange(1, startCol, 1, PRODUCTION_COL.SHEET_REMARKS - startCol + 1)
        .setValues([['Custom Components (JSON)', 'Sheet Remarks']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    Log.error('[ensureProductionExtraColumns] Error:', error.message);
  }
}

/**
 * Backfills the "Process ID", "Lot Number" and "Consumed From Process Qty"
 * columns on Production sheets created before the multi-process feature
 * existed, so legacy rows don't throw when read/written.
 */
function ensureProductionProcessColumns(sheet) {
  try {
    if (sheet.getLastColumn() < PRODUCTION_COL.CONSUMED_FROM_PROCESS_QTY) {
      const startCol = sheet.getLastColumn() + 1;
      sheet.insertColumnsAfter(sheet.getLastColumn(), PRODUCTION_COL.CONSUMED_FROM_PROCESS_QTY - sheet.getLastColumn());
      sheet.getRange(1, startCol, 1, PRODUCTION_COL.CONSUMED_FROM_PROCESS_QTY - startCol + 1)
        .setValues([['Process ID', 'Lot Number', 'Consumed From Process Qty']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    Log.error('[ensureProductionProcessColumns] Error:', error.message);
  }
}

/**
 * Backfills the "Contractor Rate" and "Contractor Payable" columns on
 * Production sheets created before the contractor rate card feature
 * existed, so legacy rows don't throw when read/written.
 */
function ensureProductionContractorColumns(sheet) {
  try {
    if (sheet.getLastColumn() < PRODUCTION_COL.CONTRACTOR_PAYABLE) {
      const startCol = sheet.getLastColumn() + 1;
      sheet.insertColumnsAfter(sheet.getLastColumn(), PRODUCTION_COL.CONTRACTOR_PAYABLE - sheet.getLastColumn());
      sheet.getRange(1, startCol, 1, PRODUCTION_COL.CONTRACTOR_PAYABLE - startCol + 1)
        .setValues([['Contractor Rate', 'Contractor Payable']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    Log.error('[ensureProductionContractorColumns] Error:', error.message);
  }
}

/**
 * Backfills the "Output Item Name" and "Components Consumed (JSON)" columns
 * on Production sheets created before the Warehouse Pool feature existed.
 */
function ensureProductionWarehouseColumns(sheet) {
  try {
    if (sheet.getLastColumn() < PRODUCTION_COL.COMPONENTS_CONSUMED) {
      const startCol = sheet.getLastColumn() + 1;
      sheet.insertColumnsAfter(sheet.getLastColumn(), PRODUCTION_COL.COMPONENTS_CONSUMED - sheet.getLastColumn());
      sheet.getRange(1, startCol, 1, PRODUCTION_COL.COMPONENTS_CONSUMED - startCol + 1)
        .setValues([['Output Item Name', 'Components Consumed (JSON)']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    Log.error('[ensureProductionWarehouseColumns] Error:', error.message);
  }
}

/**
 * Backfills the "Color" column on Production sheets created before the
 * color sub-group feature existed, so legacy rows don't throw when read/written.
 */
function ensureProductionColorColumn(sheet) {
  try {
    if (sheet.getLastColumn() < PRODUCTION_COL.COLOR) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), PRODUCTION_COL.COLOR - sheet.getLastColumn());
      sheet.getRange(1, PRODUCTION_COL.COLOR, 1, 1)
        .setValues([['Color']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    Log.error('[ensureProductionColorColumn] Error:', error.message);
  }
}

/**
 * Backfills the "Color Breakdown" column on Production sheets created before
 * a single lot could cover a multi-color batch, so legacy rows don't throw
 * when read/written.
 */
function ensureProductionColorBreakdownColumn(sheet) {
  try {
    if (sheet.getLastColumn() < PRODUCTION_COL.COLOR_BREAKDOWN) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), PRODUCTION_COL.COLOR_BREAKDOWN - sheet.getLastColumn());
      sheet.getRange(1, PRODUCTION_COL.COLOR_BREAKDOWN, 1, 1)
        .setValues([['Color Breakdown (JSON)']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    Log.error('[ensureProductionColorBreakdownColumn] Error:', error.message);
  }
}

/**
 * Structured Order Number reference (see PRODUCTION_COL.ORDER_NUMBER) —
 * closes the gap where deleteClientOrder/deleteClientOrdersBulk could only
 * check Dispatch, never Production, since the PI->Production link used to
 * live only in a free-text Remarks note.
 */
function ensureProductionOrderNumberColumn(sheet) {
  try {
    if (sheet.getLastColumn() < PRODUCTION_COL.ORDER_NUMBER) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), PRODUCTION_COL.ORDER_NUMBER - sheet.getLastColumn());
      sheet.getRange(1, PRODUCTION_COL.ORDER_NUMBER, 1, 1)
        .setValues([['Order Number']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    Log.error('[ensureProductionOrderNumberColumn] Error:', error.message);
  }
}

/**
 * Reads all Process Master rows (active and inactive) directly from the
 * sheet, sorted by Sequence ascending. Internal helper shared by
 * saveProduction/_computeProcessWipMap so they don't pay the
 * buildResponse()-wrapping overhead of getProcessData() on every call.
 * @private
 */
function _getAllProcessRecords() {
  let sheet;
  try {
    sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_MASTER);
  } catch (e) {
    initProcessMasterSheet();
    sheet = getSheet(APP_CONFIG.SHEETS.PROCESS_MASTER);
  }

  ensureProcessOutputItemColumn(sheet);
  ensureProcessPrimaryColorAxisColumn(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  const records = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const processId = String(row[PROCESS_COL.PROCESS_ID - 1] || '').trim();
    if (!processId) continue;

    records.push({
      processId: processId,
      processName: String(row[PROCESS_COL.PROCESS_NAME - 1] || '').trim(),
      sequence: Number(row[PROCESS_COL.SEQUENCE - 1]) || 0,
      lotPrefix: String(row[PROCESS_COL.LOT_PREFIX - 1] || '').trim().toUpperCase(),
      isFinalStage: row[PROCESS_COL.IS_FINAL_STAGE - 1] === true || String(row[PROCESS_COL.IS_FINAL_STAGE - 1]).toUpperCase() === 'TRUE',
      outputItemName: String(row[PROCESS_COL.OUTPUT_ITEM_NAME - 1] || '').trim(),
      primaryColorAxis: String(row[PROCESS_COL.PRIMARY_COLOR_AXIS - 1] || '').trim()
    });
  }

  records.sort((a, b) => a.sequence - b.sequence);
  return records;
}

/**
 * Public read endpoint: for a given process, returns the current Warehouse
 * Pool availability of every POOL-sourced component in its recipe, so the
 * UI can show an operator how much of an upstream stage's output is
 * available before they commit a lot that consumes it.
 * @param {string} processId
 */
function getProcessWipData(processId) {
  try {
    const componentsResp = getProcessComponentsData(processId);
    const components = (componentsResp && componentsResp.data) || [];

    const poolQtyMap = getPoolAvailableQtyMap();
    const records = components
      .filter(c => c.sourceType === COMPONENT_SOURCE_TYPES.POOL)
      .map(c => {
        const entry = poolQtyMap[String(c.itemName || '').trim().toLowerCase()];
        const colorGroup = String(c.colorGroup || '').trim();
        const isColorScoped = colorGroup && colorGroup.toUpperCase() !== COMPONENT_COLOR_GROUP_COMMON;
        const availableQty = !entry ? 0 : (isColorScoped ? (entry.byColor[colorGroup.toLowerCase()] || 0) : entry.total);
        return {
          outputItemName: c.itemName,
          availableQty: availableQty
        };
      });

    return buildResponse(true, records);
  } catch (error) {
    Log.error('[getProcessWipData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load process WIP data: ' + error.message);
  }
}

/**
 * Generates the next sequential Lot Number for a given process, scoped to
 * that process's Lot Prefix (e.g. "LOT-FP-0012"). Scans the existing Lot
 * Number column for the current max, mirroring getNextDispatchNumber's
 * regex-scan approach.
 * @private
 */
function _generateLotNumber(sheet, lotPrefix) {
  const prefix = String(lotPrefix || '').trim().toUpperCase();
  const lastRow = sheet.getLastRow();
  let maxNum = 0;

  if (lastRow >= 2) {
    const lotNumbers = sheet.getRange(2, PRODUCTION_COL.LOT_NUMBER, lastRow - 1, 1).getValues();
    const pattern = new RegExp('^LOT-' + prefix + '-(\\d+)$', 'i');
    lotNumbers.forEach(row => {
      const match = String(row[0] || '').trim().match(pattern);
      if (match) {
        const num = parseInt(match[1], 10);
        if (num > maxNum) maxNum = num;
      }
    });
  }

  return 'LOT-' + prefix + '-' + String(maxNum + 1).padStart(4, '0');
}

/**
 * Retrieves all production lots from the sheet.
 * Includes the physical sheet row index (rowIdx) for edit/delete targeting.
 */
function getProductionData() {
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
    } catch (e) {
      initProductionSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
    }

    ensureProductionExtraColumns(sheet);
    ensureProductionProcessColumns(sheet);
    ensureProductionContractorColumns(sheet);
    ensureProductionWarehouseColumns(sheet);
    ensureProductionColorColumn(sheet);
    ensureProductionColorBreakdownColumn(sheet);
    ensureProductionOrderNumberColumn(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return buildResponse(true, []);
    }

    const numCols = Math.max(sheet.getLastColumn(), PRODUCTION_COL.COLOR_BREAKDOWN);
    const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    const records = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rawDate = row[PRODUCTION_COL.DATE - 1];
      const dateStr = rawDate instanceof Date ? toSafeDateString(rawDate) : String(rawDate || '');

      const processId = String(row[PRODUCTION_COL.PROCESS_ID - 1] || '').trim();
      if (!processId) continue;

      let componentsConsumed = [];
      const componentsRaw = String(row[PRODUCTION_COL.COMPONENTS_CONSUMED - 1] || '').trim();
      if (componentsRaw) {
        try {
          const parsed = JSON.parse(componentsRaw);
          if (Array.isArray(parsed)) componentsConsumed = parsed;
        } catch (e) {
          Log.error('[getProductionData] Invalid componentsConsumed JSON on row', i + 2, ':', e.message);
        }
      }

      let customComponents = [];
      const customComponentsRaw = String(row[PRODUCTION_COL.CUSTOM_COMPONENTS - 1] || '').trim();
      if (customComponentsRaw) {
        try {
          const parsed = JSON.parse(customComponentsRaw);
          if (Array.isArray(parsed)) customComponents = parsed;
        } catch (e) {
          Log.error('[getProductionData] Invalid customComponents JSON on row', i + 2, ':', e.message);
        }
      }

      let colorBreakdown = [];
      const colorBreakdownRaw = String(row[PRODUCTION_COL.COLOR_BREAKDOWN - 1] || '').trim();
      if (colorBreakdownRaw) {
        try {
          const parsed = JSON.parse(colorBreakdownRaw);
          if (Array.isArray(parsed)) colorBreakdown = parsed;
        } catch (e) {
          Log.error('[getProductionData] Invalid colorBreakdown JSON on row', i + 2, ':', e.message);
        }
      }

      records.push({
        rowIdx: i + 2, // 2-indexed row position for direct sheet operations
        date: dateStr,
        dateRaw: rawDate instanceof Date ? rawDate.toISOString() : null,
        productId: String(row[PRODUCTION_COL.PRODUCT_ID - 1] || '').trim(),
        productName: String(row[PRODUCTION_COL.PRODUCT_NAME - 1] || '').trim(),
        qty: Number(row[PRODUCTION_COL.QTY - 1]) || 0,
        assignedBy: String(row[PRODUCTION_COL.ASSIGNED_BY - 1] || '').trim(),
        assignedTo: String(row[PRODUCTION_COL.ASSIGNED_TO - 1] || '').trim(),
        status: String(row[PRODUCTION_COL.STATUS - 1] || '').trim(),
        remarks: String(row[PRODUCTION_COL.REMARKS - 1] || '').trim(),
        customComponents: customComponents,
        sheetRemarks: String(row[PRODUCTION_COL.SHEET_REMARKS - 1] || '').trim(),
        processId: processId,
        lotNumber: String(row[PRODUCTION_COL.LOT_NUMBER - 1] || '').trim(),
        contractorPayable: Number(row[PRODUCTION_COL.CONTRACTOR_PAYABLE - 1]) || 0,
        outputItemName: String(row[PRODUCTION_COL.OUTPUT_ITEM_NAME - 1] || '').trim(),
        componentsConsumed: componentsConsumed,
        color: String(row[PRODUCTION_COL.COLOR - 1] || '').trim(),
        colorBreakdown: colorBreakdown
      });
    }
    
    // Sort production lots by date descending, then rowIdx descending (newest
    // first). Timestamp is computed once per record (not per comparison) —
    // records.length comparisons in a sort would otherwise re-parse the same
    // Date string O(n log n) times instead of O(n).
    records.forEach(r => { r._sortTs = r.dateRaw ? new Date(r.dateRaw).getTime() : 0; });
    records.sort((a, b) => {
      if (b._sortTs !== a._sortTs) return b._sortTs - a._sortTs;
      return b.rowIdx - a.rowIdx;
    });
    records.forEach(r => { delete r._sortTs; });
    
    return buildResponse(true, records);
  } catch (error) {
    Log.error('[getProductionData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load production data: ' + error.message);
  }
}

/**
 * @private
 * Key for a POOL-sourced component need — color-scoped needs (colorGroup
 * other than COMMON) are tracked per-color since the pool tracks colors
 * separately; a COMMON need is tracked against the item's total across
 * every color bucket. Shared by saveProduction and updateProductionStatus
 * so both validate pool availability the same way.
 */
function _poolNeedKey(itemNameLower, colorGroup) {
  const isColorScoped = colorGroup && colorGroup.toUpperCase() !== COMPONENT_COLOR_GROUP_COMMON;
  return itemNameLower + '||' + (isColorScoped ? colorGroup.toLowerCase() : '');
}

/**
 * @private
 * Builds a _poolNeedKey -> {itemName, colorGroup, isColorScoped, qty} map of
 * a lot's POOL-sourced component needs from a components array (either the
 * client's freshly-submitted list, or a parsed COMPONENTS_CONSUMED snapshot).
 */
function _buildPoolNeededMap(components) {
  const poolNeeded = {};
  // Lazily loaded only if some component actually carries a non-blank Unit
  // (see PROCESS_COMPONENTS_COL.UNIT) — mirrors module_warehouse.js#
  // recalculateWarehousePool's Pass 2, which is what actually debits the
  // pool once this lot completes. This map must convert the same way that
  // debit does, or a lot with a unit-converted POOL component could pass
  // this availability check yet still overdraw the pool once completed
  // (or vice versa, be rejected here despite fitting the true debit).
  let itemUnitMap = null;
  let unitsMap = null;
  (components || []).forEach(c => {
    if (String(c.sourceType || '').toUpperCase() !== COMPONENT_SOURCE_TYPES.POOL) return;
    const itemName = String(c.itemName || '').trim();
    if (!itemName) return;
    const colorGroup = String(c.colorGroup || '').trim();
    const itemNameLower = itemName.toLowerCase();
    const isColorScoped = colorGroup && colorGroup.toUpperCase() !== COMPONENT_COLOR_GROUP_COMMON;
    const key = _poolNeedKey(itemNameLower, colorGroup);
    if (!poolNeeded[key]) poolNeeded[key] = { itemName, colorGroup: isColorScoped ? colorGroup : '', isColorScoped, qty: 0 };

    let qty = Number(c.qty) || 0;
    const unit = String(c.unit || '').trim();
    if (unit && typeof convertQtyToBaseUnit === 'function') {
      if (!itemUnitMap) itemUnitMap = typeof _getItemUnitInfoMap === 'function' ? _getItemUnitInfoMap() : {};
      if (!unitsMap) unitsMap = typeof _getUnitsMap === 'function' ? _getUnitsMap() : {};
      const unitInfo = typeof _lookupItemUnitInfo === 'function'
        ? _lookupItemUnitInfo(itemUnitMap, itemName, '')
        : { baseUnit: 'Pcs', purchaseUnit: 'Pcs', weightPerBaseUnit: 0 };
      try {
        qty = convertQtyToBaseUnit(qty, unit, unitInfo, unitsMap);
      } catch (e) {
        // Unconvertible — fall back to the as-entered qty, same as the
        // Warehouse Pool debit this is validating against.
      }
    }

    poolNeeded[key].qty += qty;
  });
  return poolNeeded;
}

/**
 * @private
 * Validates a poolNeeded map (see _buildPoolNeededMap) against current
 * Warehouse Pool availability. `alreadyConsumed` is an optional map, same
 * keys as poolNeeded, of qty this same lot already debited under its
 * PREVIOUS saved state — pass it ONLY when that previous state was itself
 * Completed (i.e. actually already subtracted from the pool by
 * recalculateWarehousePool, which debits Completed rows only — see
 * module_warehouse.js). Passing it unconditionally would let a lot's own
 * stale/never-debited claim get counted back into "available", permitting
 * an overdraft once the lot is later completed.
 * @returns {string|null} an error message, or null if everything fits.
 */
function _validatePoolAvailability(poolNeeded, alreadyConsumed) {
  const poolAvailableMap = getPoolAvailableQtyMap();
  for (const key in poolNeeded) {
    const need = poolNeeded[key];
    const itemNameLower = need.itemName.toLowerCase();
    const entry = poolAvailableMap[itemNameLower];
    const currentAvailableQty = !entry ? 0 : (need.isColorScoped ? (entry.byColor[need.colorGroup.toLowerCase()] || 0) : entry.total);
    const availableForThisLot = currentAvailableQty + ((alreadyConsumed && alreadyConsumed[key]) || 0);
    if (need.qty > availableForThisLot + 0.0001) {
      const label = need.isColorScoped ? `${need.itemName}" in color "${need.colorGroup}` : need.itemName;
      return `Only ${availableForThisLot} unit(s) of "${label}" are available in the Warehouse Pool.`;
    }
  }
  return null;
}

/**
 * Saves a production run (creates a new entry or updates an existing row).
 */
function saveProduction(formData) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(PRODUCTION_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }
  
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
    } catch (e) {
      initProductionSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
    }

    ensureProductionExtraColumns(sheet);
    ensureProductionProcessColumns(sheet);
    ensureProductionContractorColumns(sheet);
    ensureProductionWarehouseColumns(sheet);
    ensureProductionColorColumn(sheet);
    ensureProductionColorBreakdownColumn(sheet);
    ensureProductionOrderNumberColumn(sheet);

    const processId = sanitizeString(formData.processId, 'processId');
    if (!processId) {
      return buildResponse(false, null, 'A Process must be selected for this lot.');
    }

    const assignedTo = sanitizeString(formData.assignedTo || '', 'assignedTo');
    if (!assignedTo) {
      return buildResponse(false, null, 'A Contractor (Assigned To) is required for this lot.');
    }

    const allProcesses = _getAllProcessRecords();
    const process = allProcesses.find(p => p.processId.toLowerCase() === processId.toLowerCase());
    if (!process) {
      return buildResponse(false, null, `Process "${processId}" was not found. It may have been deleted.`);
    }

    // Color: whenever this process has at least one color sub-group defined
    // on its recipe (see getProcessColorGroups), the lot is logged as a
    // single-row batch covering every color produced in this run — qty is
    // derived as the sum of the breakdown instead of being sent separately,
    // so a multi-color run never has to be split into multiple lots/rows.
    // Process Components/Warehouse Pool/Process Color Links are read once
    // here and reused below for the Primary Axis resolution (~line 640) —
    // getProcessColorGroups() and computeColorAxesForProcess() both derive
    // from the exact same three inputs, so fetching them twice per lot save
    // (this is the busiest write path in the app) doubled the sheet reads
    // for no benefit.
    const _colorComponents = (getProcessComponentsData(processId).data || []);
    const _colorPoolRows = typeof getWarehousePoolData === 'function' ? (getWarehousePoolData().data || []) : [];
    const _colorLinks = _getAllProcessColorLinks();
    // Process Color Overrides (see excludeWarehousePoolColors/
    // includeWarehousePoolColor) must apply here too, not just to the
    // client's checklist — otherwise a color the operator just excluded via
    // the Warehouse Pool dialog would still validate as accepted on a stale
    // form that had it checked before the exclusion, and a color they just
    // force-INCLUDEd wouldn't validate as accepted at all.
    const _colorOverrides = (typeof _getAllProcessColorOverrides === 'function' ? _getAllProcessColorOverrides() : {})[String(processId || '').trim().toLowerCase()];
    const availableColorGroups = computeColorGroupsForProcess(processId, _colorComponents, _colorPoolRows, _colorLinks, undefined, _colorOverrides);

    let qty;
    let color = '';
    let colorBreakdown = [];

    // Parsed eagerly (not just when availableColorGroups > 0) because a
    // client-added custom sub-group (see Script.html's addCustomColorRow)
    // may be the ONLY color on this lot, for a process whose recipe/pool
    // otherwise has zero auto-detected color groups.
    let rawBreakdown = [];
    if (formData.colorBreakdown) {
      try {
        rawBreakdown = typeof formData.colorBreakdown === 'string'
          ? JSON.parse(formData.colorBreakdown)
          : formData.colorBreakdown;
      } catch (e) {
        return buildResponse(false, null, 'Invalid color breakdown data format.');
      }
      if (!Array.isArray(rawBreakdown)) rawBreakdown = [];
    }
    const hasCustomBreakdown = rawBreakdown.some(c => c && c.isCustom);

    if (availableColorGroups.length > 0 || hasCustomBreakdown) {
      colorBreakdown = rawBreakdown
        .map(c => ({
          color: sanitizeString(c.color || '', 'color'),
          size: sanitizeString(c.size || '', 'size'),
          // Zero and negative quantities are both allowed — a lot can be
          // logged as a correction/reversal (e.g. a prior over-count)
          // without editing history, mirroring
          // adjustStockManually/adjustWarehousePoolManually.
          qty: validateNumber(c.qty, -10000000, 10000000),
          isCustom: !!c.isCustom,
          // See Script.html's getCheckedColorQtys — false for a true
          // Color-Axis non-primary row, or a legacy-grouping "Other" row
          // that's segment-matching a checked row elsewhere in the
          // checklist (both describe the same physical batch as their
          // counterpart, not an additional one). Defaults true when absent
          // so an older client payload / a row with no such distinction
          // still counts, same as before this field existed.
          countsTowardTotal: c.countsTowardTotal !== false,
          // This row's checklist group/axis key at check time (see
          // Script.html's getCheckedColorQtys — row.dataset.group), used
          // below to identify a genuine primary-axis row by axis IDENTITY
          // rather than by color name, which can collide across axes.
          axisKey: sanitizeString(c.axisKey || '', 'colorAxisKey')
        }))
        .filter(c => c.color);

      if (colorBreakdown.length === 0) {
        return buildResponse(false, null, 'At least one Color is required for this lot (this process has color-specific components).');
      }

      // A custom sub-group (operator-typed at production time, not part of
      // this process's recipe/pool colors) is deliberately exempt — that's
      // the whole point of the "+ Add Custom Sub-Group" control. Every other
      // entry still must match a real configured color sub-group, guarding
      // against a stale dropdown (e.g. a color removed from the recipe after
      // this form was opened).
      const invalidColor = colorBreakdown.find(c => !c.isCustom && !availableColorGroups.some(ag => ag.toLowerCase() === c.color.toLowerCase()));
      if (invalidColor) {
        return buildResponse(false, null, `Color "${invalidColor.color}" is not a configured color sub-group for this process. It may have been removed — refresh and re-select.`);
      }

      // Size is descriptive only — recorded per color for this lot's own
      // record-keeping (lot list, print sheet) but not a Warehouse Pool
      // inventory dimension; Pool buckets stay keyed by Output Item Name +
      // Product Tag + Color exactly as before, regardless of Size.
      color = colorBreakdown.map(c => c.size ? `${c.color} (${c.size})` : c.color).join(', ');

      // Quantity: when this process has a Primary Color Axis configured
      // (see PROCESS_COL.PRIMARY_COLOR_AXIS / computeColorAxesForProcess),
      // its checklist is split into independent groups — e.g. Base, Rim
      // Color, Mudguard Color — and only the PRIMARY group's rows describe
      // real, additional output batches. A non-primary group's checked
      // rows (e.g. "Mudguard Color: Red") describe the SAME batch from a
      // different angle, not extra units, so summing every row (legacy
      // behavior below) would double-count them. Sum only the rows whose
      // color belongs to the primary axis instead. Falls back to legacy
      // (sum everything) if the configured primary axis can't be resolved
      // — e.g. it was renamed/removed since — so a stale config never
      // silently zeroes out a lot's quantity.
      //
      // The operator can also pick (or change) the Primary Axis directly on
      // the Production Lot form (see Script.html's
      // _buildColorAxisGroupHeader/setPrimaryColorAxisChoice) — that choice
      // arrives as formData.primaryColorAxis and takes precedence over the
      // process's stored default for THIS lot, so it's correct even if the
      // process has never been configured (or was just changed) in the
      // Process editor. It's then persisted back onto the process below so
      // future lots default to it too.
      const submittedPrimaryColorAxis = sanitizeString(formData.primaryColorAxis || '', 'primaryColorAxis');
      const primaryColorAxis = submittedPrimaryColorAxis || String(process.primaryColorAxis || '').trim();
      let primaryAxisColorsLower = null;
      let primaryAxisKeyLower = null;
      if (primaryColorAxis) {
        const axes = computeColorAxesForProcess(processId, _colorComponents, _colorPoolRows, _colorLinks);
        const primaryAxis = axes.find(a => a.label.toLowerCase() === primaryColorAxis.toLowerCase());
        if (primaryAxis) {
          primaryAxisColorsLower = new Set(primaryAxis.colors.map(c => c.toLowerCase()));
          primaryAxisKeyLower = primaryAxis.key.toLowerCase();
        }
      }

      // No formal Primary Color Axis resolved (either never configured, or
      // this is the legacy pool-signature-grouped checklist — see
      // Script.html's renderGroupedColorChecklist): fall back to each
      // entry's own countsTowardTotal instead of blindly summing everything,
      // so a legacy "Other" row the client auto-matched to a pool-backed
      // group's row (e.g. "Pink" <- "Pink / BCP / Pink-White") doesn't
      // double-count that same physical batch into qty.
      //
      // A custom sub-group (see Script.html's addCustomColorRow) is by
      // definition NOT part of any configured recipe axis, so it can never
      // appear in primaryAxisColorsLower on its own — counted here via its
      // own countsTowardTotal instead, exactly like the legacy branch below
      // already does for every entry. The operator picks which real group a
      // custom color belongs to when 2+ exist (see addCustomColorRow): one
      // placed in the Primary group, or left independent of any group,
      // arrives with countsTowardTotal true (a genuine extra output batch);
      // one placed in a NON-primary group arrives false (it describes the
      // SAME batch a different way, same as a real non-primary row) and
      // must not also add to the total, or it would double-count exactly
      // like the pre-Color-Axes bug this whole feature fixed. Without the
      // isCustom branch here at all, EVERY custom color's quantity was
      // silently dropped from the saved lot's total, since a custom name
      // can never match a real recipe-known axis color.
      // Matching purely by color NAME against primaryAxisColorsLower breaks
      // once two independent axes happen to share a color name (e.g. a Rim
      // axis and a separate Frame axis both having their own "Purple" —
      // each axis's colors come from its own pool item's independent color
      // history, with no cross-axis uniqueness guarantee — see
      // computeColorAxesForProcess). A non-primary axis row sharing the
      // primary axis's color name then got wrongly counted as an extra
      // primary batch, inflating qty. Prefer the row's own axisKey (its
      // dataset-group at check time — see Script.html's getCheckedColorQtys)
      // when present: exact axis identity, immune to name collisions.
      // Falls back to the name-only match for older submitted payloads with
      // no axisKey, or a row from the legacy (non-axis) checklist.
      const isPrimaryAxisRow = c => {
        const isKnownPrimaryColor = c.axisKey
          ? c.axisKey.toLowerCase() === primaryAxisKeyLower
          : primaryAxisColorsLower.has(c.color.toLowerCase());
        const countsAsCustom = c.isCustom && c.countsTowardTotal !== false;
        return isKnownPrimaryColor || countsAsCustom;
      };

      qty = primaryAxisColorsLower
        ? colorBreakdown.reduce((sum, c) => sum + (isPrimaryAxisRow(c) ? c.qty : 0), 0)
        : colorBreakdown.reduce((sum, c) => sum + (c.countsTowardTotal ? c.qty : 0), 0);

      // The total itself may legitimately be zero or negative (a
      // correction/reversal lot, or an explicit zero-output record) — see
      // the colorBreakdown mapping above. What's still rejected is a
      // Primary-Axis-configured process where NO checked row actually
      // belongs to the primary group at all (e.g. only a non-primary
      // "Mudguard Color" row was checked) — that's an operator forgetting
      // to address the primary axis, not a deliberate zero, and would
      // otherwise silently save attribute-only consumption against zero
      // real output.
      if (primaryAxisColorsLower && !colorBreakdown.some(isPrimaryAxisRow)) {
        return buildResponse(false, null, `At least one "${primaryColorAxis}" color is required for this lot.`);
      }

      // The operator explicitly picked/changed the Primary Axis on this
      // lot's form and it resolved to a real axis (primaryAxisColorsLower
      // is only set once a match was found) — adopt it as the process's own
      // default going forward, same as configuring it in the Process editor.
      if (submittedPrimaryColorAxis && primaryAxisColorsLower &&
          submittedPrimaryColorAxis.toLowerCase() !== String(process.primaryColorAxis || '').trim().toLowerCase() &&
          typeof _setProcessPrimaryColorAxis === 'function') {
        _setProcessPrimaryColorAxis(processId, submittedPrimaryColorAxis);
      }
    } else {
      // Zero and negative quantities are both allowed — a lot can be logged
      // as a correction/reversal (e.g. a prior over-count) without editing
      // history, mirroring adjustStockManually/adjustWarehousePoolManually.
      qty = validateNumber(formData.qty, -10000000, 10000000);
    }

    // Components consumed: required list of {itemName, size, color, sourceType, qty, colorGroup}.
    // colorGroup is 'COMMON' (applies to every color in this batch, including
    // manually-added rows with no recipe-derived colorGroup) or a specific
    // Color Master name matching one of this lot's colorBreakdown entries.
    let componentsConsumed;
    try {
      componentsConsumed = typeof formData.componentsConsumed === 'string'
        ? JSON.parse(formData.componentsConsumed)
        : (formData.componentsConsumed || []);
    } catch (e) {
      return buildResponse(false, null, 'Invalid components consumed data format.');
    }
    if (!Array.isArray(componentsConsumed)) componentsConsumed = [];

    let cleanComponents = componentsConsumed
      .map(c => ({
        itemName: sanitizeString(c.itemName || '', 'itemName'),
        size: sanitizeString(c.size || '', 'size'),
        narration: sanitizeString(c.narration || '', 'narration'),
        color: sanitizeString(c.color || '', 'color'),
        sourceType: String(c.sourceType || '').trim().toUpperCase() === COMPONENT_SOURCE_TYPES.POOL
          ? COMPONENT_SOURCE_TYPES.POOL
          : COMPONENT_SOURCE_TYPES.ITEM,
        // Zero and negative consumption are both allowed — e.g. a correction/
        // reversal lot that credits stock/pool back, mirroring the output
        // qty's own zero/negative handling above.
        qty: validateNumber(c.qty, -10000000, 10000000),
        colorGroup: sanitizeString(c.colorGroup || '', 'colorGroup') || COMPONENT_COLOR_GROUP_COMMON,
        // Blank = "already in the item's Base Unit" — see PROCESS_COMPONENTS_COL.UNIT
        // and module_stock.js#_getBilledAndConsumedQtyMaps.
        unit: sanitizeString(c.unit || '', 'unit')
      }))
      .filter(c => c.itemName);

    if (colorBreakdown.length > 0) {
      // Drop any leftover row scoped to a color that isn't part of this
      // batch (e.g. left behind after a color was unchecked in the UI). A
      // breakdown color may itself be a composite of 2+ independent pool
      // axes (e.g. "BCP / Blue-White" — see COLOR_COMBO_DELIMITER); a
      // component row scoped to just one axis's literal color ("BCP") is
      // still valid for that batch even though it doesn't match the full
      // composite string, so every individual axis token is accepted too.
      const breakdownColorsLower = new Set();
      colorBreakdown.forEach(c => {
        breakdownColorsLower.add(c.color.toLowerCase());
        c.color.split(COLOR_COMBO_DELIMITER).forEach(token => breakdownColorsLower.add(token.trim().toLowerCase()));
      });
      cleanComponents = cleanComponents.filter(c =>
        c.colorGroup === COMPONENT_COLOR_GROUP_COMMON || breakdownColorsLower.has(c.colorGroup.toLowerCase())
      );
    }

    if (cleanComponents.length === 0) {
      return buildResponse(false, null, 'At least one component consumed is required for this lot.');
    }

    // Only final-stage processes may carry an optional Product tag, used
    // purely so Dispatch can find this lot's pool credit.
    const productId = process.isFinalStage ? sanitizeString(formData.productId || '', 'productId') : '';
    const productName = process.isFinalStage ? sanitizeString(formData.productName || '', 'productName') : '';

    // Process Date
    const dateObj = toSafeDateObject(formData.date) || new Date();

    const assignedBy = sanitizeString(formData.assignedBy || '', 'assignedBy');
    const status = sanitizeString(formData.status || 'Pending', 'status');
    const remarks = sanitizeString(formData.remarks || '', 'remarks');

    const rowData = [
      dateObj,
      productId,
      productName,
      qty,
      assignedBy,
      assignedTo,
      status,
      remarks
    ];

    const isEdit = !!formData.rowIdx;
    let targetRow = isEdit ? parseInt(formData.rowIdx, 10) : -1;
    let lotNumber;

    let targetRowValues = null;
    if (isEdit) {
      if (isNaN(targetRow) || targetRow < 2 || targetRow > sheet.getLastRow()) {
        return buildResponse(false, null, 'Invalid production record selected for edit.');
      }

      // Single batched read of the row — replaces 3 separate per-column
      // getRange().getValue() round-trips (process id, lot number, components
      // consumed) used below and further down in this function.
      targetRowValues = sheet.getRange(targetRow, 1, 1, PRODUCTION_COL.COLOR_BREAKDOWN).getValues()[0];

      // The Process determines lot numbering and the Output Item Name —
      // changing it on an existing lot would orphan its lot number and
      // corrupt pool totals, so it must be deleted and recreated instead.
      const currentProcessId = String(targetRowValues[PRODUCTION_COL.PROCESS_ID - 1]).trim();
      if (currentProcessId && currentProcessId.toLowerCase() !== processId.toLowerCase()) {
        return buildResponse(false, null, 'Process cannot be changed on an existing lot. Delete and recreate it under the new process instead.');
      }

      lotNumber = String(targetRowValues[PRODUCTION_COL.LOT_NUMBER - 1]).trim()
        || _generateLotNumber(sheet, process.lotPrefix);
    } else {
      lotNumber = _generateLotNumber(sheet, process.lotPrefix);
    }

    // Validate POOL-sourced consumption against current Warehouse Pool
    // availability. A component scoped to a specific color (colorGroup other
    // than COMMON) must be checked against that color's own bucket, since
    // the pool now tracks colors separately; a COMMON component is checked
    // against the item's total across every color bucket.
    const poolNeeded = _buildPoolNeededMap(cleanComponents);

    // originalPoolConsumed represents pool stock this SAME lot already
    // debited under its previously-saved state — only true when that
    // previous state's Status was itself 'Completed' (recalculateWarehousePool
    // debits Completed rows only). Gating on previousStatus here matters:
    // trusting a Pending lot's never-debited claim let an edit inflate its
    // own apparent availability, permitting an overdraft once later completed.
    let originalPoolConsumed = {};
    let originalHasItemSourced = false;
    if (isEdit) {
      const previousStatus = String(targetRowValues[PRODUCTION_COL.STATUS - 1] || '').trim().toLowerCase();
      const originalRaw = String(targetRowValues[PRODUCTION_COL.COMPONENTS_CONSUMED - 1] || '').trim();
      if (originalRaw) {
        try {
          const parsed = JSON.parse(originalRaw);
          if (Array.isArray(parsed)) {
            parsed.forEach(c => {
              if (String(c.sourceType || '').toUpperCase() !== COMPONENT_SOURCE_TYPES.POOL) {
                originalHasItemSourced = true;
                return;
              }
              if (previousStatus !== 'completed') return;
              const itemNameLower = String(c.itemName || '').trim().toLowerCase();
              const colorGroup = String(c.colorGroup || '').trim();
              const key = _poolNeedKey(itemNameLower, colorGroup);
              originalPoolConsumed[key] = (originalPoolConsumed[key] || 0) + (Number(c.qty) || 0);
            });
          }
        } catch (e) { /* ignore malformed legacy data */ }
      }
    }

    // Insufficient Warehouse Pool stock no longer blocks the save — the
    // operator may legitimately need to log a lot before its upstream stock
    // catches up on paper. _validatePoolAvailability's message is instead
    // carried through as a warning appended to the success response below,
    // and the shortfall is left to show up as a negative Available Qty in
    // the Warehouse Pool (recalculateWarehousePool computes producedQty -
    // consumedQty directly with no floor at 0 — see module_warehouse.js).
    //
    // Only worth checking when this save will actually result in a
    // 'Completed' lot — recalculateWarehousePool only ever debits Completed
    // rows, so a Pending (or Pending-staying-Pending) save can't yet affect
    // the pool, and running the check anyway produced a misleading "pool
    // will go negative" warning on lots that hadn't actually consumed
    // anything.
    const poolWarning = status.toLowerCase() === 'completed'
      ? _validatePoolAvailability(poolNeeded, originalPoolConsumed)
      : null;

    // Snapshot the contractor rate at save time (Assigned To IS the
    // contractor reference). If Assigned To has no matching rate card entry
    // for this process — e.g. it's an in-house supervisor's name — both
    // values stay 0, since the contractor link is optional.
    const contractorRate = typeof _getContractorRate === 'function'
      ? _getContractorRate(assignedTo, process.processName)
      : 0;
    const contractorPayable = contractorRate * qty;
    const componentsJson = JSON.stringify(cleanComponents);
    const colorBreakdownJson = colorBreakdown.length > 0 ? JSON.stringify(colorBreakdown) : '';

    if (isEdit) {
      sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
      sheet.getRange(targetRow, PRODUCTION_COL.PROCESS_ID, 1, 5).setValues([[processId, lotNumber, 0, contractorRate, contractorPayable]]);
      sheet.getRange(targetRow, PRODUCTION_COL.OUTPUT_ITEM_NAME, 1, 4).setValues([[process.outputItemName, componentsJson, color, colorBreakdownJson]]);
    } else {
      // Append row, then fill in the process columns on the newly created row
      sheet.appendRow(rowData);
      const newRow = sheet.getLastRow();
      sheet.getRange(newRow, PRODUCTION_COL.PROCESS_ID, 1, 5).setValues([[processId, lotNumber, 0, contractorRate, contractorPayable]]);
      sheet.getRange(newRow, PRODUCTION_COL.OUTPUT_ITEM_NAME, 1, 4).setValues([[process.outputItemName, componentsJson, color, colorBreakdownJson]]);
    }

    // recalculateStock() rebuilds the whole Stock sheet from every ITEM-sourced
    // Completed-lot consumption — skip it when this lot has no ITEM-sourced
    // component, before or after the edit, since Stock can't have changed.
    const newHasItemSourced = cleanComponents.some(c => c.sourceType !== COMPONENT_SOURCE_TYPES.POOL);
    if ((newHasItemSourced || originalHasItemSourced) && typeof recalculateStock === 'function') {
      recalculateStock();
    }
    recalculateWarehousePool();

    SpreadsheetApp.flush();

    const tagSuffix = productId ? ` (tagged: ${productName || productId})` : '';
    const colorSuffix = color ? `, Color: ${color}` : '';
    const logMsg = isEdit
      ? `Production lot updated, Process: ${process.processName}, Lot: ${lotNumber}, Qty: ${qty}${colorSuffix}${tagSuffix}`
      : `Production lot recorded, Process: ${process.processName}, Lot: ${lotNumber}, Qty: ${qty}${colorSuffix}${tagSuffix}`;

    logAction(isEdit ? 'UPDATE' : 'CREATE', APP_CONFIG.SHEETS.PRODUCTION, processId, logMsg, 'SUCCESS');

    const baseMsg = isEdit ? 'Production log updated successfully.' : 'Production log saved successfully.';
    const successMsg = poolWarning ? `${baseMsg} Warning: ${poolWarning} Warehouse Pool stock will now show negative for this item.` : baseMsg;
    return buildResponse(true, { lotNumber: lotNumber }, successMsg);
  } catch (error) {
    Log.error('[saveProduction] Error:', error.message);
    logAction('ERROR', 'saveProduction', formData.processId || 'NEW', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to save production: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Reads a Production row's Components Consumed JSON and reports whether it
 * contains at least one ITEM-sourced (non-POOL) component — i.e. whether
 * deleting/changing this row could affect the Stock sheet, so callers can
 * skip the expensive recalculateStock() full rebuild when it's a pure
 * Warehouse-Pool-only lot.
 * @private
 */
function _rowHasItemSourcedComponentFromRaw(raw) {
  try {
    if (!raw) return true;
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return true;
    return parsed.some(c => String(c.sourceType || '').toUpperCase() !== COMPONENT_SOURCE_TYPES.POOL);
  } catch (e) {
    return true;
  }
}

/**
 * Deletes a production log entry.
 */
function deleteProduction(rowIdx, expectedProductId, expectedQty) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(PRODUCTION_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
    if (!sheet) throw new Error('Production sheet not found.');

    const targetRow = parseInt(rowIdx, 10);
    if (isNaN(targetRow) || targetRow < 2 || targetRow > sheet.getLastRow()) {
      return buildResponse(false, null, 'Invalid production record selected for deletion.');
    }

    // Single batched read (Product ID through Components Consumed) instead
    // of a separate getRange().getValue() round-trip for each.
    const hasComponentsCol = sheet.getLastColumn() >= PRODUCTION_COL.COMPONENTS_CONSUMED;
    const rowWidth = (hasComponentsCol ? PRODUCTION_COL.COMPONENTS_CONSUMED : 4) - PRODUCTION_COL.PRODUCT_ID + 1;
    const idQtyRow = sheet.getRange(targetRow, PRODUCTION_COL.PRODUCT_ID, 1, rowWidth).getValues()[0];
    const productId = String(idQtyRow[PRODUCTION_COL.PRODUCT_ID - PRODUCTION_COL.PRODUCT_ID]).trim();
    const qty = Number(idQtyRow[PRODUCTION_COL.QTY - PRODUCTION_COL.PRODUCT_ID]) || 0;

    // Safety check to ensure we do not delete shifted/incorrect rows
    if (expectedProductId !== undefined && expectedQty !== undefined) {
      if (productId.toLowerCase() !== String(expectedProductId || '').trim().toLowerCase() ||
          Math.abs(qty - Number(expectedQty)) > 0.0001) {
        return buildResponse(false, null, 'Data mismatch: The record has been modified or shifted. Please refresh.');
      }
    }

    // Legacy sheets predating the Warehouse Pool model don't have this
    // column at all — default to "yes, recalc" so we don't silently skip
    // a Stock rebuild we can't actually rule out on old data.
    const hadItemSourced = !hasComponentsCol || _rowHasItemSourcedComponentFromRaw(
      String(idQtyRow[PRODUCTION_COL.COMPONENTS_CONSUMED - PRODUCTION_COL.PRODUCT_ID] || '').trim()
    );

    // Un-tagged (intermediate-stage) lots credit the Warehouse Pool; if a
    // downstream lot already drew on that credit, removing it here can leave
    // the pool negative. Non-blocking (see _checkPoolCreditRemovalWarning) —
    // only surfaced when this row actually IS untagged, since a final-stage
    // (productId-tagged) lot's credit is drawn down by Dispatch instead, a
    // different mechanism this check doesn't cover.
    let poolCreditWarning = null;
    if (!productId && sheet.getLastColumn() >= PRODUCTION_COL.COLOR_BREAKDOWN &&
        typeof _checkPoolCreditRemovalWarning === 'function') {
      const extraVals = sheet.getRange(targetRow, PRODUCTION_COL.OUTPUT_ITEM_NAME, 1,
        PRODUCTION_COL.COLOR_BREAKDOWN - PRODUCTION_COL.OUTPUT_ITEM_NAME + 1).getValues()[0];
      const outputItemName = String(extraVals[0]).trim();
      let colorBreakdown = null;
      try {
        const parsed = JSON.parse(String(extraVals[PRODUCTION_COL.COLOR_BREAKDOWN - PRODUCTION_COL.OUTPUT_ITEM_NAME] || '') || '[]');
        if (Array.isArray(parsed) && parsed.length > 0) colorBreakdown = parsed;
      } catch (e) { /* legacy/flat lot — treated as flat below */ }
      poolCreditWarning = _checkPoolCreditRemovalWarning(outputItemName, colorBreakdown, qty);
    }

    sheet.deleteRow(targetRow);

    if (hadItemSourced && typeof recalculateStock === 'function') {
      recalculateStock();
    }
    recalculateWarehousePool();

    SpreadsheetApp.flush();

    const msg = `Production record deleted for Product "${productId}" (Qty was ${qty}).`;
    logAction('DELETE', APP_CONFIG.SHEETS.PRODUCTION, productId, msg, 'SUCCESS');

    return buildResponse(true, null, poolCreditWarning
      ? `Production log deleted successfully. ${poolCreditWarning}`
      : 'Production log deleted successfully.');
  } catch (error) {
    Log.error('[deleteProduction] Error:', error.message);
    logAction('ERROR', 'deleteProduction', String(rowIdx), error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete production: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes multiple production log entries in a single batch.
 * @param {Array<number|string>} rowIdxs - 1-based sheet row indexes to delete
 * @param {Array<{rowIdx:number|string, expectedProductId:string, expectedQty:number}>} [expectedRows] -
 *   Optional per-row guard (same mismatch check as the single-row deleteProduction)
 *   so a row that shifted/changed since the client loaded its list is skipped
 *   instead of silently deleting whatever now sits at that row number.
 */
function deleteProductionBulk(rowIdxs, expectedRows) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(PRODUCTION_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
    if (!sheet) throw new Error('Production sheet not found.');

    const lastRow = sheet.getLastRow();
    let targetRows = (rowIdxs || [])
      .map(r => parseInt(r, 10))
      .filter(r => !isNaN(r) && r >= 2 && r <= lastRow);

    if (targetRows.length === 0) {
      return buildResponse(true, null, 'No production records selected.');
    }

    const expectedByRow = {};
    (expectedRows || []).forEach(e => {
      if (e && e.rowIdx !== undefined) expectedByRow[String(e.rowIdx)] = e;
    });

    // Single batched read (Product ID through Color Breakdown, when present)
    // covering every row in range, instead of one getRange().getValue()
    // round-trip per row — extended past Components Consumed (same span
    // deleteProduction() reads for one row) to also cover Output Item Name
    // and Color Breakdown, needed for the pool-credit-removal warning below.
    const hasComponentsCol = sheet.getLastColumn() >= PRODUCTION_COL.COMPONENTS_CONSUMED;
    const hasColorBreakdownCol = sheet.getLastColumn() >= PRODUCTION_COL.COLOR_BREAKDOWN;
    const rowWidth = (hasColorBreakdownCol ? PRODUCTION_COL.COLOR_BREAKDOWN
      : hasComponentsCol ? PRODUCTION_COL.COMPONENTS_CONSUMED : 4) - PRODUCTION_COL.PRODUCT_ID + 1;
    const idQtyComponentsRows = sheet.getRange(2, PRODUCTION_COL.PRODUCT_ID, lastRow - 1, rowWidth).getValues();

    let skippedMismatch = 0;
    let hadItemSourced = false;
    const poolCreditWarnings = [];
    targetRows = targetRows.filter(rowNum => {
      const rowVals = idQtyComponentsRows[rowNum - 2];
      const productId = String(rowVals[PRODUCTION_COL.PRODUCT_ID - PRODUCTION_COL.PRODUCT_ID]).trim();
      const qty = Number(rowVals[PRODUCTION_COL.QTY - PRODUCTION_COL.PRODUCT_ID]) || 0;

      const expected = expectedByRow[String(rowNum)];
      if (expected && expected.expectedProductId !== undefined && expected.expectedQty !== undefined) {
        if (productId.toLowerCase() !== String(expected.expectedProductId || '').trim().toLowerCase() ||
            Math.abs(qty - Number(expected.expectedQty)) > 0.0001) {
          skippedMismatch++;
          return false;
        }
      }

      // Legacy sheets predating the Warehouse Pool model don't have this
      // column at all — default to "yes, recalc" so we don't silently skip
      // a Stock rebuild we can't actually rule out on old data.
      if (!hasComponentsCol || _rowHasItemSourcedComponentFromRaw(
        String(rowVals[PRODUCTION_COL.COMPONENTS_CONSUMED - PRODUCTION_COL.PRODUCT_ID] || '').trim())) {
        hadItemSourced = true;
      }

      // Same non-blocking pool-credit-removal check as deleteProduction()
      // (see _checkPoolCreditRemovalWarning) — only applies to untagged
      // (intermediate-stage) rows.
      if (!productId && hasColorBreakdownCol && typeof _checkPoolCreditRemovalWarning === 'function') {
        const outputItemName = String(rowVals[PRODUCTION_COL.OUTPUT_ITEM_NAME - PRODUCTION_COL.PRODUCT_ID] || '').trim();
        let colorBreakdown = null;
        try {
          const parsed = JSON.parse(String(rowVals[PRODUCTION_COL.COLOR_BREAKDOWN - PRODUCTION_COL.PRODUCT_ID] || '') || '[]');
          if (Array.isArray(parsed) && parsed.length > 0) colorBreakdown = parsed;
        } catch (e) { /* legacy/flat lot — treated as flat below */ }
        const warning = _checkPoolCreditRemovalWarning(outputItemName, colorBreakdown, qty);
        if (warning) poolCreditWarnings.push(warning);
      }
      return true;
    });

    if (targetRows.length === 0) {
      return buildResponse(false, null, 'Data mismatch: the selected record(s) have been modified or shifted. Please refresh.');
    }

    const targetRowSet = new Set(targetRows);
    const { rowsDeleted } = rewriteSheetExcludingRows(sheet, 2, (_row, rowNum) => targetRowSet.has(rowNum));

    if (hadItemSourced && typeof recalculateStock === 'function') {
      recalculateStock();
    }
    recalculateWarehousePool();

    SpreadsheetApp.flush();

    let msg = skippedMismatch > 0
      ? `Deleted ${rowsDeleted} production record(s). Skipped ${skippedMismatch} that were modified or shifted — please refresh and retry those.`
      : `Deleted ${rowsDeleted} production record(s).`;
    if (poolCreditWarnings.length > 0) msg += ' ' + poolCreditWarnings.join(' ');
    logAction('BULK_DELETE', APP_CONFIG.SHEETS.PRODUCTION, 'multiple', msg, 'SUCCESS');

    return buildResponse(true, null, msg);
  } catch (error) {
    Log.error('[deleteProductionBulk] Error:', error.message);
    logAction('ERROR', 'deleteProductionBulk', 'multiple', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete production records: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

const PRODUCTION_STATUS_OPTIONS = Object.freeze(['Pending', 'In Progress', 'Completed', 'Cancelled']);

/**
 * Updates only the Status of a production lot — lets an operator change a
 * lot's status directly from the table row without opening the full Edit
 * Lot form.
 * @param {number|string} rowIdx - 1-based sheet row to update
 * @param {number|string} expectedQty - Qty the client last saw, to guard against editing a shifted row
 * @param {string} newStatus
 */
function updateProductionStatus(rowIdx, expectedQty, newStatus) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(PRODUCTION_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
    if (!sheet) throw new Error('Production sheet not found.');

    const targetRow = parseInt(rowIdx, 10);
    if (isNaN(targetRow) || targetRow < 2 || targetRow > sheet.getLastRow()) {
      return buildResponse(false, null, 'Invalid production record selected.');
    }

    const status = sanitizeString(newStatus || '', 'status');
    if (!PRODUCTION_STATUS_OPTIONS.includes(status)) {
      return buildResponse(false, null, `Invalid status "${status}".`);
    }

    // Single batched read (Qty + Status + Components Consumed) — Status and
    // Components Consumed are needed below to gate the pool-availability
    // check on this lot actually transitioning INTO Completed.
    const rowVals = sheet.getRange(targetRow, 1, 1, PRODUCTION_COL.COMPONENTS_CONSUMED).getValues()[0];
    const qty = Number(rowVals[PRODUCTION_COL.QTY - 1]) || 0;
    if (expectedQty !== undefined && Math.abs(qty - Number(expectedQty)) > 0.0001) {
      return buildResponse(false, null, 'Data mismatch: The record has been modified or shifted. Please refresh.');
    }

    // recalculateWarehousePool only debits a lot's POOL-sourced components
    // once its Status is Completed (see module_warehouse.js) — so completing
    // a lot from this quick row-select is exactly the transition that can
    // overdraw the pool, same as saveProduction validates when logging/
    // editing a lot straight to Completed. Any other transition (staying
    // Completed, or leaving it) doesn't change what this lot draws from the
    // pool, so no re-check is needed there.
    // Insufficient Warehouse Pool stock no longer blocks completing a lot —
    // same relaxation as saveProduction (see its comment) — the shortfall's
    // message is instead carried through as a warning on the success
    // response, and the pool is left to go negative.
    const previousStatus = String(rowVals[PRODUCTION_COL.STATUS - 1] || '').trim().toLowerCase();
    let poolWarning = null;
    if (status.toLowerCase() === 'completed' && previousStatus !== 'completed') {
      let components = [];
      try {
        const parsed = JSON.parse(String(rowVals[PRODUCTION_COL.COMPONENTS_CONSUMED - 1] || '') || '[]');
        if (Array.isArray(parsed)) components = parsed;
      } catch (e) { /* malformed/legacy — nothing to validate against */ }

      poolWarning = _validatePoolAvailability(_buildPoolNeededMap(components), null);
    }

    sheet.getRange(targetRow, PRODUCTION_COL.STATUS).setValue(status);

    // recalculateStock()'s consumedQtyMap only nets an ITEM-sourced
    // component's qty out of Stock while its Production row's Status is
    // 'Completed' (module_stock.js) — so ANY status change on a lot with at
    // least one ITEM-sourced component (into Completed, out of Completed, or
    // between two non-Completed statuses that still net differently, e.g.
    // Pending -> Cancelled has no effect but is cheap to just recompute)
    // can change what Stock should reflect. Every OTHER status-mutating path
    // (saveProduction, deleteProduction, deleteProductionBulk) already
    // triggers this same conditional recalc — this quick row-select path was
    // the one gap, leaving Current Stock stale after a status change here
    // until some unrelated action happened to trigger a full rebuild.
    if (previousStatus !== status.toLowerCase() &&
        _rowHasItemSourcedComponentFromRaw(String(rowVals[PRODUCTION_COL.COMPONENTS_CONSUMED - 1] || '').trim()) &&
        typeof recalculateStock === 'function') {
      recalculateStock();
    }

    if (typeof recalculateWarehousePool === 'function') {
      recalculateWarehousePool();
    }

    SpreadsheetApp.flush();

    const lotNumber = String(rowVals[PRODUCTION_COL.LOT_NUMBER - 1] || '').trim();
    logAction('UPDATE', APP_CONFIG.SHEETS.PRODUCTION, lotNumber || String(targetRow), `Production lot status changed to "${status}" (Lot: ${lotNumber || targetRow}).`, 'SUCCESS');

    const statusMsg = poolWarning ? `Status updated successfully. Warning: ${poolWarning} Warehouse Pool stock will now show negative for this item.` : 'Status updated successfully.';
    return buildResponse(true, { status: status }, statusMsg);
  } catch (error) {
    Log.error('[updateProductionStatus] Error:', error.message);
    logAction('ERROR', 'updateProductionStatus', String(rowIdx), error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to update status: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Saves a per-lot customization of the "Required Component Items" table
 * shown on the Production Sheet (e.g. customer-requested substitutions),
 * along with remarks describing what was changed and why/for whom.
 */
function saveProductionSheet(rowIdx, expectedProductId, expectedQty, customComponents, sheetRemarks) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(PRODUCTION_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
    if (!sheet) throw new Error('Production sheet not found.');

    ensureProductionExtraColumns(sheet);

    const targetRow = parseInt(rowIdx, 10);
    if (isNaN(targetRow) || targetRow < 2 || targetRow > sheet.getLastRow()) {
      return buildResponse(false, null, 'Invalid production record selected.');
    }

    // Safety check to ensure we do not overwrite a shifted/incorrect row.
    // Product Id and Qty are adjacent columns (2-4), so read them together.
    const idQtyRow = sheet.getRange(targetRow, PRODUCTION_COL.PRODUCT_ID, 1, 3).getValues()[0];
    const productId = String(idQtyRow[0]).trim();
    const qty = Number(idQtyRow[2]) || 0;
    if (expectedProductId !== undefined && expectedQty !== undefined) {
      if (productId.toLowerCase() !== String(expectedProductId || '').trim().toLowerCase() ||
          Math.abs(qty - Number(expectedQty)) > 0.0001) {
        return buildResponse(false, null, 'Data mismatch: The record has been modified or shifted. Please refresh.');
      }
    }

    let components;
    try {
      components = typeof customComponents === 'string'
        ? JSON.parse(customComponents)
        : (customComponents || []);
    } catch (e) {
      return buildResponse(false, null, 'Invalid component data format.');
    }

    if (!Array.isArray(components)) {
      return buildResponse(false, null, 'Invalid component data format.');
    }

    const cleanComponents = components
      .map(comp => ({
        itemName: sanitizeString(comp.itemName || '', 'itemName'),
        size: sanitizeString(comp.size || '', 'size'),
        narration: sanitizeString(comp.narration || '', 'narration'),
        color: sanitizeString(comp.color || '', 'color'),
        requiredQty: validateNumber(comp.requiredQty, 0, 10000000)
      }))
      .filter(comp => comp.itemName);

    const remarks = sanitizeString(sheetRemarks || '', 'sheetRemarks').slice(0, APP_CONFIG.VALIDATION.MAX_REMARKS_LENGTH);

    sheet.getRange(targetRow, PRODUCTION_COL.CUSTOM_COMPONENTS, 1, 2)
      .setValues([[JSON.stringify(cleanComponents), remarks]]);

    SpreadsheetApp.flush();

    logAction('UPDATE', APP_CONFIG.SHEETS.PRODUCTION, productId, `Production sheet customization saved for ${productId} (row ${targetRow}).`, 'SUCCESS');

    return buildResponse(true, { customComponents: cleanComponents, sheetRemarks: remarks }, 'Production sheet saved successfully.');
  } catch (error) {
    Log.error('[saveProductionSheet] Error:', error.message);
    logAction('ERROR', 'saveProductionSheet', String(rowIdx), error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to save production sheet: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Repoints an item's old name+size to its new name+size inside every
 * Production lot's already-saved COMPONENTS_CONSUMED JSON (and the legacy
 * CUSTOM_COMPONENTS JSON) snapshot. Called from module_items.js's
 * _propagateItemIdentityChange after a rename (saveItem) or a merge
 * (mergeItemEdit).
 *
 * Without this, a completed lot's Components Consumed snapshot keeps
 * pointing at the item's old name — module_stock.js's
 * _getBilledAndConsumedQtyMaps() keys its consumedQtyMap off that snapshot's
 * itemName, so once Stock's own row is renamed (syncStockForItem('rename'))
 * the stale snapshot silently stops matching on the next recalculateStock()
 * and that lot's consumption drops out of Current Stock math entirely.
 *
 * Only ITEM-sourced COMPONENTS_CONSUMED entries are touched — a POOL-sourced
 * entry's itemName is an upstream process's Output Item Name (Warehouse
 * Pool), not an Items Master reference, so it must never be repointed by an
 * Items Master rename. CUSTOM_COMPONENTS (module_production.js's
 * "sheet customization" display snapshot) has no sourceType field to check —
 * it's a display-only list of Items Master components, so every matching
 * entry is repointed.
 *
 * Runs under the caller's existing document lock — does not acquire its own.
 *
 * @param {string} oldName
 * @param {string} oldSize
 * @param {string} newName
 * @param {string} newSize
 */
function backfillProductionConsumedItemRefs(oldName, oldSize, newName, newSize) {
  let sheet;
  try {
    sheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
  } catch (e) {
    return;
  }

  const startRow = APP_CONFIG.PRODUCTION_SETTINGS.DATA_START_ROW;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;
  if (sheet.getLastColumn() < PRODUCTION_COL.COMPONENTS_CONSUMED) return;

  const numRows = lastRow - startRow + 1;
  const tOldName = String(oldName || '').trim().toLowerCase();
  const tOldSize = String(oldSize || '').trim().toLowerCase();

  const _repoint = function(raw, matchSourceType) {
    const trimmed = String(raw || '').trim();
    if (!trimmed) return null;

    let list;
    try {
      list = JSON.parse(trimmed);
    } catch (e) {
      return null;
    }
    if (!Array.isArray(list) || list.length === 0) return null;

    let changed = false;
    list.forEach(function(comp) {
      if (matchSourceType) {
        const sourceType = String(comp.sourceType || '').trim().toUpperCase();
        if (sourceType === COMPONENT_SOURCE_TYPES.POOL) return;
      }
      const rowName = String(comp.itemName || '').trim().toLowerCase();
      const rowSize = String(comp.size || '').trim().toLowerCase();
      if (rowName === tOldName && rowSize === tOldSize) {
        comp.itemName = newName;
        comp.size = newSize;
        changed = true;
      }
    });

    return changed ? JSON.stringify(list) : null;
  };

  const consumedRange = sheet.getRange(startRow, PRODUCTION_COL.COMPONENTS_CONSUMED, numRows, 1);
  const consumedValues = consumedRange.getValues();
  let consumedChanged = false;
  for (let i = 0; i < consumedValues.length; i++) {
    const rewritten = _repoint(consumedValues[i][0], true);
    if (rewritten !== null) {
      consumedValues[i][0] = rewritten;
      consumedChanged = true;
    }
  }
  if (consumedChanged) {
    consumedRange.setValues(consumedValues);
  }

  if (sheet.getLastColumn() >= PRODUCTION_COL.CUSTOM_COMPONENTS) {
    const customRange = sheet.getRange(startRow, PRODUCTION_COL.CUSTOM_COMPONENTS, numRows, 1);
    const customValues = customRange.getValues();
    let customChanged = false;
    for (let i = 0; i < customValues.length; i++) {
      const rewritten = _repoint(customValues[i][0], false);
      if (rewritten !== null) {
        customValues[i][0] = rewritten;
        customChanged = true;
      }
    }
    if (customChanged) {
      customRange.setValues(customValues);
    }
  }
}
