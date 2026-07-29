/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_warehouse.gs — WAREHOUSE POOL MODULE
 *
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Tracks ready stock of intermediate and finished process outputs (e.g.
 * "Painted Frame", "Fitted Rim", "Fitted Frame", "Packed Bicycle") that sit
 * between production stages. Credited when a Production lot for the
 * originating process is marked Completed; debited when a downstream
 * Production lot consumes it as a POOL-sourced component, or when a
 * Product-tagged final-stage credit is Dispatched.
 *
 * Sheet Layout (Warehouse Pool):
 * ───────────────────────────────────────────────────────────────────────────
 * Col A (1):   Output Item Name
 * Col B (2):   Process ID (originating process)
 * Col C (3):   Product Tag (blank for intermediate WIP; Product ID for
 *              tagged final-stage finished goods)
 * Col D (4):   Produced Qty
 * Col E (5):   Consumed Qty
 * Col F (6):   Available Qty (= Produced Qty - Consumed Qty)
 * Col G (7):   Color (Color Master name this bucket was credited under;
 *              blank for color-agnostic/uncolored output — a multi-color
 *              lot's output is split into one bucket per color so a
 *              downstream process can draw a specific color out of the pool)
 * ═══════════════════════════════════════════════════════════════════════════
 */

const WAREHOUSE_LOCK_TIMEOUT_MS = 15000;

/**
 * Initializes the Warehouse Pool sheet with correct headers.
 */
function initWarehousePoolSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
    if (!sheet) {
      sheet = ss.insertSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
    }

    const headers = ['Output Item Name', 'Process ID', 'Product Tag', 'Produced Qty', 'Consumed Qty', 'Available Qty', 'Color'];

    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Warehouse Pool sheet initialized successfully.');
  } catch (error) {
    Log.error('[initWarehousePoolSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize Warehouse Pool sheet: ' + error.message);
  }
}

/**
 * Backfills the "Color" column on Warehouse Pool sheets created before
 * per-color pool tracking existed, so legacy sheets don't throw when
 * read/written.
 */
function ensureWarehousePoolColorColumn(sheet) {
  try {
    if (sheet.getLastColumn() < WAREHOUSE_POOL_COL.COLOR) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), WAREHOUSE_POOL_COL.COLOR - sheet.getLastColumn());
      sheet.getRange(1, WAREHOUSE_POOL_COL.COLOR, 1, 1)
        .setValues([['Color']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    Log.error('[ensureWarehousePoolColorColumn] Error:', error.message);
  }
}

/**
 * @private
 * Builds the key used to bucket a pool row: Output Item Name + Product Tag
 * (blank tag for intermediate WIP) + Color (blank for color-agnostic output).
 */
function _poolKey(outputItemName, productTag, color) {
  return String(outputItemName || '').trim().toLowerCase() + '||' + String(productTag || '').trim().toLowerCase() + '||' + String(color || '').trim().toLowerCase();
}

/**
 * Resolves a single-axis-token colorGroup (e.g. "BCP") against a set of
 * live bucket color strings for the same item, one of which may be a
 * composite of 2+ independent pool axes (e.g. "BCP / Blue-White" — see
 * COLOR_COMBO_DELIMITER). A Process Component recipe row's Color Sub-Group
 * is configured manually from Color Master (see Script_Process.html's
 * addColorGroup) — independently of whatever string the upstream item's
 * own credits actually landed under — so an exact-string bucket match
 * alone can miss a composite bucket that legitimately contains this
 * token, treating real available/producible stock as an empty phantom
 * bucket instead (and, on the debit side, would create+overdraw that
 * phantom bucket while the real composite one never gets debited).
 *
 * Only resolves when EXACTLY ONE candidate composite color contains the
 * token as one of its parts — a token shared by 2+ composite buckets is
 * genuinely ambiguous (which one should this consumption be attributed
 * to?) and is deliberately left unresolved rather than guessed.
 *
 * @param {string[]} candidateColors - lowercased, trimmed color strings live for this item (may include the exact token itself)
 * @param {string} tokenLower - lowercased, trimmed single-axis token being sought
 * @returns {string|null} the one matching composite color string, or null if there's an exact match already / nothing resolves
 */
/**
 * @private Splits a (possibly composite) pool color into its axis segments.
 * "Blue-White / Black / Grey" -> ["Blue-White", "Black", "Grey"]; a plain
 * single-axis color yields a one-element array.
 */
function _colorSegments(color) {
  return String(color || '').split(COLOR_COMBO_DELIMITER).map(s => s.trim()).filter(Boolean);
}

/**
 * @private Order-independent identity for a composite color: its segments
 * lowercased and sorted, joined with a delimiter that can't occur in a color
 * name. "Blue-White / Black / Grey" and "Blue-White / Grey / Black" describe
 * the SAME physical unit — one frame color, one rim color, one mudguard
 * color — so they must resolve to one bucket. Used to catch (and heal)
 * historical rows credited before _composeLotColorKey imposed a canonical
 * order, and by verifyProductionColorChain to report the split.
 */
function _colorOrderKey(color) {
  return _colorSegments(color).map(s => s.toLowerCase()).sort().join('\u0000');
}

/**
 * @private Builds a lot's composite bucket color in a CANONICAL, repeatable
 * segment order, so the same real combination always keys the same bucket.
 *
 * EVERY axis takes the position THIS PROCESS'S OWN RECIPE gives it — the
 * primary included (see getAxisOrderByProcess / computeColorAxesForProcess).
 * A POOL recipe row is the association with the upstream process that
 * produces it, so recipe row order is exactly "this process's inputs, in
 * the sequence the operator arranged them", and it is also the order the
 * Production checklist renders. So a recipe listing Fitted Rim above
 * Painted Frame credits "Black / Blue-White", and the operator can change
 * that reading by reordering the recipe.
 *
 * The primary axis is NOT anchored first. It was until 2026-07-29, which
 * meant a lot of frames-on-black-rims read "Blue-White / Black" no matter
 * where the recipe put the rim — the quantity-bearing axis is not
 * necessarily the one you name first.
 *
 * Order previously came straight from the Color Breakdown array, i.e. from
 * checklist DOM order, which followed Warehouse Pool sheet row order — and
 * that is itself rebuilt on every recalculation. Two lots of the very same
 * product could therefore be credited as "Blue-White / Black / Grey" and
 * "Blue-White / Grey / Black" and have their stock split across two buckets.
 * Needs 3+ axes (2+ independent ones) to bite.
 *
 * A primary color that is itself a composite (inherited from upstream)
 * stays intact as one unit — only its position among this stage's axes is
 * decided here, never its internal order.
 *
 * @param {{color:string, axisKey:string}} primaryEntry This lot's primary-axis entry.
 * @param {Array<{color:string, axisKey:string}>} independentEntries
 * @param {Object} [axisOrder] { [axisKeyLower]: position } for this lot's own
 *   process — see getAxisOrderByProcess. An axis missing from it (renamed,
 *   removed, or a legacy entry with no axisKey at all) sorts after every
 *   known one, then by axis key and color, so the result is always fully
 *   determined even when the recipe can no longer explain an entry.
 */
function _composeLotColorKey(primaryEntry, independentEntries, axisOrder) {
  const order = axisOrder || {};
  const positionOf = (e) => {
    const k = String((e && e.axisKey) || '').trim().toLowerCase();
    return (k && Object.prototype.hasOwnProperty.call(order, k)) ? order[k] : Number.MAX_SAFE_INTEGER;
  };
  const ordered = [primaryEntry].concat(independentEntries || [])
    .filter(e => e && String(e.color || '').trim())
    .sort((a, b) => {
      const pa = positionOf(a);
      const pb = positionOf(b);
      if (pa !== pb) return pa - pb;
      const ka = String(a.axisKey || '').trim().toLowerCase();
      const kb = String(b.axisKey || '').trim().toLowerCase();
      if (ka !== kb) return ka < kb ? -1 : 1;
      // Same (or absent) axis key — fall back to the color itself so the
      // result is still fully determined rather than input-order dependent.
      const ca = String(a.color || '').trim().toLowerCase();
      const cb = String(b.color || '').trim().toLowerCase();
      return ca < cb ? -1 : (ca > cb ? 1 : 0);
    });
  return ordered.map(e => String(e.color || '').trim()).join(COLOR_COMBO_DELIMITER);
}

function _resolveCompositeColorToken(candidateColors, tokenLower) {
  if (!tokenLower || candidateColors.indexOf(tokenLower) !== -1) return null;
  const matches = new Set(
    candidateColors.filter(c =>
      c.indexOf(COLOR_COMBO_DELIMITER) !== -1 &&
      c.split(COLOR_COMBO_DELIMITER).some(t => t.trim() === tokenLower)
    )
  );
  return matches.size === 1 ? Array.from(matches)[0] : null;
}

// ── Warehouse Pool Opening Balances ─────────────────────────────────────
// Manual seed entries for stock that already existed before this app went
// live (or a one-off correction). Unlike the Warehouse Pool sheet itself —
// fully wiped and rebuilt by recalculateWarehousePool() on every call —
// this sheet is its own persistent source of truth, read once per rebuild
// and folded into the matching bucket's Produced Qty. Shares
// WAREHOUSE_LOCK_TIMEOUT_MS with the pool sheet itself since both take the
// same document lock.

/**
 * Initializes the Warehouse Pool Opening sheet with correct headers.
 */
function initWarehousePoolOpeningSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
    if (!sheet) {
      sheet = ss.insertSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
    }

    const headers = ['Output Item Name', 'Process ID', 'Product Tag', 'Color', 'Qty', 'Date', 'Remarks'];

    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Warehouse Pool Opening sheet initialized successfully.');
  } catch (error) {
    Log.error('[initWarehousePoolOpeningSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize Warehouse Pool Opening sheet: ' + error.message);
  }
}

/**
 * Lists every opening-balance entry, with the originating Process Name
 * resolved for display (the sheet itself only stores Process ID).
 */
function getWarehousePoolOpeningData() {
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
    } catch (e) {
      initWarehousePoolOpeningSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, []);

    const allProcesses = typeof _getAllProcessRecords === 'function' ? _getAllProcessRecords() : [];
    const processNameById = {};
    allProcesses.forEach(p => { processNameById[p.processId.toLowerCase()] = p.processName; });

    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const records = data
      .map((row, i) => {
        const processId = String(row[WAREHOUSE_POOL_OPENING_COL.PROCESS_ID - 1] || '').trim();
        const rawDate = row[WAREHOUSE_POOL_OPENING_COL.DATE - 1];
        return {
          rowIdx: i + 2,
          outputItemName: String(row[WAREHOUSE_POOL_OPENING_COL.OUTPUT_ITEM_NAME - 1] || '').trim(),
          processId: processId,
          processName: processNameById[processId.toLowerCase()] || processId,
          productTag: String(row[WAREHOUSE_POOL_OPENING_COL.PRODUCT_TAG - 1] || '').trim(),
          color: String(row[WAREHOUSE_POOL_OPENING_COL.COLOR - 1] || '').trim(),
          qty: Number(row[WAREHOUSE_POOL_OPENING_COL.QTY - 1]) || 0,
          date: rawDate instanceof Date ? toSafeDateString(rawDate) : String(rawDate || ''),
          dateRaw: rawDate instanceof Date ? rawDate.toISOString() : null,
          remarks: String(row[WAREHOUSE_POOL_OPENING_COL.REMARKS - 1] || '').trim()
        };
      })
      .filter(r => r.outputItemName);

    records.sort((a, b) => {
      const dateA = a.dateRaw ? new Date(a.dateRaw) : new Date(0);
      const dateB = b.dateRaw ? new Date(b.dateRaw) : new Date(0);
      if (dateB - dateA !== 0) return dateB - dateA;
      return b.rowIdx - a.rowIdx;
    });

    return buildResponse(true, records);
  } catch (error) {
    Log.error('[getWarehousePoolOpeningData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load Warehouse Pool opening balances: ' + error.message);
  }
}

/**
 * Records a new opening-balance entry for a Warehouse Pool bucket. The
 * Output Item Name is always derived from the selected Process (never
 * free-typed), so the seeded bucket always lines up with a real process's
 * output. Product Tag and Color are optional, same as a Production lot.
 * Negative quantities are allowed (zero is rejected as a no-op) since this
 * form doubles as a one-off correction — e.g. clawing back a bucket that
 * went negative from over-issued production.
 */
function saveWarehousePoolOpening(formData) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(WAREHOUSE_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
    } catch (e) {
      initWarehousePoolOpeningSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
    }

    const processId = sanitizeString(formData.processId || '', 'processId');
    if (!processId) {
      return buildResponse(false, null, 'A Process is required.');
    }

    const allProcesses = typeof _getAllProcessRecords === 'function' ? _getAllProcessRecords() : [];
    const process = allProcesses.find(p => p.processId.toLowerCase() === processId.toLowerCase());
    if (!process) {
      return buildResponse(false, null, 'Selected Process was not found.');
    }
    if (!process.outputItemName) {
      return buildResponse(false, null, `Process "${process.processName}" has no Output Item Name configured.`);
    }

    const productTag = process.isFinalStage ? sanitizeString(formData.productTag || '', 'productTag') : '';
    const color = sanitizeString(formData.color || '', 'color');

    const qty = validateNumber(formData.qty, -10000000, 10000000);
    if (qty === 0) {
      return buildResponse(false, null, 'Opening Quantity cannot be zero.');
    }

    const dateObj = toSafeDateObject(formData.date) || new Date();

    const remarks = sanitizeString(formData.remarks || '', 'remarks');

    sheet.appendRow([process.outputItemName, process.processId, productTag, color, qty, dateObj, remarks]);

    recalculateWarehousePool();
    SpreadsheetApp.flush();

    const logMsg = `Opening stock added: ${qty} x "${process.outputItemName}"${color ? `, Color: ${color}` : ''}${productTag ? ` (tagged: ${productTag})` : ''}.`;
    logAction('CREATE', APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING, process.processId, logMsg, 'SUCCESS');

    return buildResponse(true, null, 'Opening stock recorded successfully.');
  } catch (error) {
    Log.error('[saveWarehousePoolOpening] Error:', error.message);
    logAction('ERROR', 'saveWarehousePoolOpening', formData.processId || 'NEW', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to record opening stock: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * @private
 * Reads the live Warehouse Pool sheet for the exact bucket matching
 * outputItemName + productTag + color, returning its current Available Qty
 * (0 if the bucket doesn't exist yet — e.g. a brand-new opening entry).
 */
function _getWarehousePoolBucketAvailableQty(outputItemName, productTag, color) {
  let sheet;
  try {
    sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
  } catch (e) {
    return 0;
  }
  ensureWarehousePoolColorColumn(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const targetKey = _poolKey(outputItemName, productTag, color);
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowKey = _poolKey(
      row[WAREHOUSE_POOL_COL.OUTPUT_ITEM_NAME - 1],
      row[WAREHOUSE_POOL_COL.PRODUCT_TAG - 1],
      row[WAREHOUSE_POOL_COL.COLOR - 1]
    );
    if (rowKey === targetKey) {
      return Number(row[WAREHOUSE_POOL_COL.AVAILABLE_QTY - 1]) || 0;
    }
  }
  return 0;
}

/**
 * Manually corrects a Warehouse Pool bucket's Available Qty (e.g. to seed
 * opening stock for a bucket that already has Production history, or to fix
 * a physical-recount discrepancy), by appending a delta row to the Warehouse
 * Pool Opening sheet — the same durable source recalculateWarehousePool()
 * already folds into Produced Qty — and rebuilding. A reason is required and
 * the adjustment is recorded in the audit Logs sheet so corrections stay
 * traceable, mirroring adjustStockManually() for raw-material Stock. Negative
 * values are allowed — over-issued production can legitimately leave a
 * bucket negative until the user reviews and corrects it.
 * @param {string} outputItemName
 * @param {string} processId - Originating Process (required so the
 *   correction row lines up with a real process's output, same as the
 *   opening-stock form).
 * @param {string} productTag
 * @param {string} color
 * @param {number} newAvailableQty - The corrected Available Qty value.
 * @param {string} reason - Why the adjustment was made (required).
 */
function adjustWarehousePoolManually(outputItemName, processId, productTag, color, newAvailableQty, reason) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(WAREHOUSE_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const itemName = sanitizeString(outputItemName || '', 'outputItemName');
    if (!itemName) {
      return buildResponse(false, null, 'Output Item Name is required.');
    }

    const procId = sanitizeString(processId || '', 'processId');
    if (!procId) {
      return buildResponse(false, null, 'Originating Process is required.');
    }

    const allProcesses = typeof _getAllProcessRecords === 'function' ? _getAllProcessRecords() : [];
    if (!allProcesses.some(p => p.processId.toLowerCase() === procId.toLowerCase())) {
      return buildResponse(false, null, 'Originating Process was not found.');
    }

    const newQty = Number(newAvailableQty);
    if (isNaN(newQty)) {
      return buildResponse(false, null, 'Corrected quantity must be a valid number.');
    }

    const reasonText = sanitizeString(reason || '', 'reason');
    if (!reasonText) {
      return buildResponse(false, null, 'A reason is required for manual stock corrections.');
    }

    const tag = sanitizeString(productTag || '', 'productTag');
    const colorVal = sanitizeString(color || '', 'color');

    const oldQty = _getWarehousePoolBucketAvailableQty(itemName, tag, colorVal);
    const delta = newQty - oldQty;
    if (delta === 0) {
      // oldQty here is read fresh from the sheet, so it's the authoritative
      // current value even if the caller's own on-screen number (e.g. a
      // stale client-side cache after another user's change) was different.
      // Surfacing it lets the UI reconcile its display instead of being
      // stuck showing a stale value that will keep rejecting this same edit.
      // buildResponse() forces data to null on failure, so this bypasses it
      // and returns the {success,data,message} shape directly.
      return { success: false, data: { oldAvailableQty: oldQty, newAvailableQty: oldQty }, message: 'New quantity is the same as the current value — nothing to adjust.' };
    }

    let openingSheet;
    try {
      openingSheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
    } catch (e) {
      initWarehousePoolOpeningSheet();
      openingSheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
    }

    openingSheet.appendRow([itemName, procId, tag, colorVal, delta, new Date(), `Correction: ${reasonText}`]);

    recalculateWarehousePool();
    SpreadsheetApp.flush();

    // Each field is URI-encoded before joining so a literal "|" inside an
    // item name, tag, or color can't be mistaken for the delimiter and
    // corrupt the 3-way split in getWarehousePoolAdjustmentHistory().
    const recordId = [itemName, tag, colorVal].map(encodeURIComponent).join('|');
    const details = `Old: ${oldQty}, New: ${newQty}. Reason: ${reasonText}`;
    logAction('ADJUST', APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING, recordId, details, 'SUCCESS');

    return buildResponse(true, { oldAvailableQty: oldQty, newAvailableQty: newQty }, 'Warehouse Pool stock adjusted successfully.');
  } catch (error) {
    Log.error('[adjustWarehousePoolManually] Error:', error.message);
    logAction('ERROR', 'adjustWarehousePoolManually', String(outputItemName), error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to adjust Warehouse Pool stock: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Returns the history of manual Warehouse Pool corrections ('ADJUST'),
 * sourced from the audit Logs sheet (see adjustWarehousePoolManually). Used
 * by the Warehouse Pool Ledger view to surface these alongside Production
 * credits/debits and Dispatch debits.
 */
function getWarehousePoolAdjustmentHistory() {
  try {
    const logsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSheetName('LOGS'));
    if (!logsSheet) return buildResponse(true, []);

    const lastRow = logsSheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, []);

    const data = logsSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const records = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const action = String(row[LOGS_COL.ACTION - 1] || '').trim().toUpperCase();
      const sheetName = String(row[LOGS_COL.SHEET - 1] || '').trim();
      if (action !== 'ADJUST' || sheetName !== APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING) continue;

      const recordId = String(row[LOGS_COL.RECORD_ID - 1] || '');
      const idParts = recordId.split('|');
      if (idParts.length !== 3) continue;

      // recordId parts are URI-encoded (see adjustWarehousePoolManually) so a
      // literal "|" inside a value can't be mistaken for the delimiter.
      // Older log rows predate the encoding and decode back to themselves
      // unchanged as long as they contain no "%" sequences.
      let decodedParts;
      try {
        decodedParts = idParts.map(decodeURIComponent);
      } catch (e) {
        decodedParts = idParts;
      }

      const details = String(row[LOGS_COL.DETAILS - 1] || '');
      const detailsMatch = details.match(/Old:\s*([\d.\-]+),\s*New:\s*([\d.\-]+)\.\s*Reason:\s*(.*)$/);
      if (!detailsMatch) continue;

      records.push({
        date: row[LOGS_COL.TIMESTAMP - 1],
        outputItemName: decodedParts[0],
        productTag: decodedParts[1],
        color: decodedParts[2],
        oldValue: Number(detailsMatch[1]),
        newValue: Number(detailsMatch[2]),
        reason: detailsMatch[3],
        user: row[LOGS_COL.USER - 1]
      });
    }

    return buildResponse(true, records);
  } catch (error) {
    Log.error('[getWarehousePoolAdjustmentHistory] Error:', error.message);
    return buildResponse(false, null, 'Failed to load Warehouse Pool adjustment history: ' + error.message);
  }
}

/**
 * Deletes an opening-balance entry and rebuilds the pool.
 */
function deleteWarehousePoolOpening(rowIdx, expectedOutputItemName, expectedQty) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(WAREHOUSE_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
    if (!sheet) throw new Error('Warehouse Pool Opening sheet not found.');

    const targetRow = parseInt(rowIdx, 10);
    if (isNaN(targetRow) || targetRow < 2 || targetRow > sheet.getLastRow()) {
      return buildResponse(false, null, 'Invalid opening stock entry selected for deletion.');
    }

    const rowVals = sheet.getRange(targetRow, WAREHOUSE_POOL_OPENING_COL.OUTPUT_ITEM_NAME, 1,
      WAREHOUSE_POOL_OPENING_COL.QTY - WAREHOUSE_POOL_OPENING_COL.OUTPUT_ITEM_NAME + 1).getValues()[0];
    const outputItemName = String(rowVals[0]).trim();
    const qty = Number(rowVals[WAREHOUSE_POOL_OPENING_COL.QTY - WAREHOUSE_POOL_OPENING_COL.OUTPUT_ITEM_NAME]) || 0;

    // Safety check to ensure we do not delete a shifted/modified row.
    if (expectedOutputItemName !== undefined && expectedQty !== undefined) {
      if (outputItemName.toLowerCase() !== String(expectedOutputItemName || '').trim().toLowerCase() ||
          Math.abs(qty - Number(expectedQty)) > 0.0001) {
        return buildResponse(false, null, 'Data mismatch: The entry has been modified or shifted. Please refresh.');
      }
    }

    sheet.deleteRow(targetRow);
    recalculateWarehousePool();
    SpreadsheetApp.flush();

    logAction('DELETE', APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING, outputItemName, `Opening stock entry for "${outputItemName}" deleted.`, 'SUCCESS');

    return buildResponse(true, null, 'Opening stock entry deleted successfully.');
  } catch (error) {
    Log.error('[deleteWarehousePoolOpening] Error:', error.message);
    logAction('ERROR', 'deleteWarehousePoolOpening', String(rowIdx), error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete opening stock entry: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * @private
 * Reads the Warehouse Pool Opening sheet, returning each row in the same
 * shape recalculateWarehousePool()'s getBucket() expects.
 */
function _getWarehousePoolOpeningRows() {
  const rows = [];
  let sheet;
  try {
    sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL_OPENING);
  } catch (e) {
    return rows; // Opening sheet not initialized yet — nothing to seed
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return rows;

  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  data.forEach(row => {
    const outputItemName = String(row[WAREHOUSE_POOL_OPENING_COL.OUTPUT_ITEM_NAME - 1] || '').trim();
    if (!outputItemName) return;
    // Zero is skipped (no-op); negative IS kept — a downward manual
    // correction (see adjustWarehousePoolManually) appends a negative delta
    // row here, and it must survive this read or the correction silently
    // reverts on the next rebuild.
    const qty = Number(row[WAREHOUSE_POOL_OPENING_COL.QTY - 1]) || 0;
    if (qty === 0) return;

    rows.push({
      outputItemName: outputItemName,
      processId: String(row[WAREHOUSE_POOL_OPENING_COL.PROCESS_ID - 1] || '').trim(),
      productTag: String(row[WAREHOUSE_POOL_OPENING_COL.PRODUCT_TAG - 1] || '').trim(),
      color: String(row[WAREHOUSE_POOL_OPENING_COL.COLOR - 1] || '').trim(),
      qty: qty
    });
  });

  return rows;
}

/**
 * Recalculates the entire Warehouse Pool sheet from scratch based on
 * manually-recorded Opening Balances + Completed Production lots (credits)
 * and downstream POOL-sourced component consumption + finished-stage
 * Dispatch consumption (debits). Mirrors recalculateStock()'s "always
 * rebuild from source data" approach.
 */
function recalculateWarehousePool() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(WAREHOUSE_LOCK_TIMEOUT_MS)) {
    Log.error('[recalculateWarehousePool] Could not acquire lock.');
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    let poolSheet;
    try {
      poolSheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
    } catch (e) {
      initWarehousePoolSheet();
      poolSheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
    }
    ensureWarehousePoolColorColumn(poolSheet);

    let prodSheet;
    try {
      prodSheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
    } catch (e) {
      prodSheet = null;
    }

    const buckets = {}; // key -> { outputItemName, processId, productTag, color, producedQty, consumedQty }

    function getBucket(outputItemName, processId, productTag, color) {
      const key = _poolKey(outputItemName, productTag, color);
      if (!buckets[key]) {
        buckets[key] = {
          outputItemName: String(outputItemName || '').trim(),
          processId: String(processId || '').trim(),
          productTag: String(productTag || '').trim(),
          color: String(color || '').trim(),
          producedQty: 0,
          consumedQty: 0
        };
      }
      return buckets[key];
    }

    // Pass 0: seed buckets from manually-recorded Opening Balances (stock
    // that already existed before this app went live). This sheet is never
    // wiped by this rebuild, so it's the one durable source of "stock that
    // didn't come from a Production lot".
    _getWarehousePoolOpeningRows().forEach(r => {
      getBucket(r.outputItemName, r.processId, r.productTag, r.color).producedQty += r.qty;
    });

    if (prodSheet) {
      ensureProductionExtraColumns(prodSheet);
      ensureProductionProcessColumns(prodSheet);
      ensureProductionContractorColumns(prodSheet);
      ensureProductionWarehouseColumns(prodSheet);
      ensureProductionColorColumn(prodSheet);
      ensureProductionColorBreakdownColumn(prodSheet);

      const lastRow = prodSheet.getLastRow();
      if (lastRow >= 2) {
        const numCols = PRODUCTION_COL.COLOR_BREAKDOWN;
        const data = prodSheet.getRange(2, 1, lastRow - 1, numCols).getValues();

        // Axis order for every process, read once — the composite color a
        // lot is credited under lists its axes in that process's own recipe
        // order (see _composeLotColorKey / getAxisOrderByProcess), which is
        // also the order the operator saw them on the checklist.
        const axisOrderByProcess = typeof getAxisOrderByProcess === 'function'
          ? getAxisOrderByProcess()
          : {};

        // Pass 1: credit every Completed lot's own output to its pool
        // bucket(s). A color-agnostic lot credits the single blank-color
        // bucket. A multi-color lot's Color Breakdown entries (one per
        // checked color across every axis — see getCheckedColorQtys,
        // Script_Production.html) are combined into composite buckets
        // whenever the pairing is unambiguous: every OTHER
        // genuinely-independent axis (its color doesn't _colorNamesMatch
        // any primary color — e.g. a Rim Color sharing no name segment with
        // any Frame color, see module_process.js) contributes at most ONE
        // checked entry of its own. Any number of independent axes can
        // combine this way — not just a single extra one — as long as each
        // DISTINCT axis (its own axisKey) is only represented once; two
        // axes each contributing exactly one color (e.g. Frame=Blue-White,
        // Rim=BCP, Mudguard=Black) is just as unambiguous as one, and
        // combines into one 3-way "Blue-White / BCP / Black" bucket the
        // same way two would combine into "Blue-White / BCP". A redundant
        // axis (e.g. Mudguard Color, whose checked value DOES name-match
        // a primary — the same batch described a second way, per the
        // exact heuristic the Production checklist itself uses to
        // auto-sync such rows) is excluded from this combination entirely
        // — this is what turns two independent credits (10 under
        // "Red-White", 10 under "Black") into one real "Red-White / Black"
        // bucket.
        //
        // The PRIMARY axis (countsTowardTotal !== false) may hold any
        // number of checked colors — one composite bucket is emitted per
        // primary color, carrying that color's own qty, so a lot producing
        // 10 each of Blue/Pink/Purple/Red-White frames on Black rims yields
        // four "<frame> / Black" buckets totalling the lot's 40 units.
        // Until 2026-07-29 combining required EXACTLY ONE primary entry, so
        // any real multi-color lot fell back to per-entry crediting and
        // fragmented into loose half-colors (Blue-White AND Blue AND Black
        // as if each were a complete producible output), inflating the
        // credited total by one full lot per redundant axis and leaving the
        // downstream process's checklist unable to see the composite it was
        // actually meant to consume.
        //
        // Anything less clean-cut (no counted entry at all, or any one
        // axis contributing 2+ entries with no stored cross-axis pairing to
        // tell which goes with which) falls back to crediting every entry
        // under its own single color, exactly as before this combining
        // logic existed — never guessing at a quantity attribution.
        data.forEach(row => {
          const status = String(row[PRODUCTION_COL.STATUS - 1] || '').trim().toLowerCase();
          if (status !== 'completed') return;

          const processId = String(row[PRODUCTION_COL.PROCESS_ID - 1] || '').trim();
          const outputItemName = String(row[PRODUCTION_COL.OUTPUT_ITEM_NAME - 1] || '').trim();
          if (!outputItemName) return;

          const productTag = String(row[PRODUCTION_COL.PRODUCT_ID - 1] || '').trim();

          let colorBreakdown = [];
          const colorBreakdownRaw = String(row[PRODUCTION_COL.COLOR_BREAKDOWN - 1] || '').trim();
          if (colorBreakdownRaw) {
            try {
              const parsed = JSON.parse(colorBreakdownRaw);
              if (Array.isArray(parsed)) colorBreakdown = parsed;
            } catch (e) { /* ignore malformed data */ }
          }

          if (colorBreakdown.length > 0) {
            // Zero/negative kept — a negative per-color qty is a
            // correction/reversal lot that credits this bucket back down
            // (see saveProduction in module_production.js), mirroring the
            // flat (non-color) path below which never filtered by sign.
            const creditColor = (color, qty) => {
              if (!color) return;
              getBucket(outputItemName, processId, productTag, color).producedQty += (Number(qty) || 0);
            };

            const primaryEntries = colorBreakdown.filter(e => e && e.countsTowardTotal !== false && String(e.color || '').trim());
            const otherEntries = colorBreakdown.filter(e => e && e.countsTowardTotal === false && String(e.color || '').trim());

            let combined = false;
            if (primaryEntries.length >= 1) {
              // A mirror axis (e.g. Mudguard) describes the same batch a
              // second way, so it must not become its own segment. It is
              // recognised by _colorNamesMatch against ANY of the primary
              // colors — not just one — because across a multi-color lot it
              // contributes one entry per primary color (Blue against
              // Blue-White, Red against Red-White, ...).
              //
              // The exception is an INHERITED-SEGMENT COLLISION. When the
              // primary color is itself a composite carried down from
              // upstream, its segments include other processes' axes, and a
              // downstream axis cannot be a mirror of one of those — it just
              // happens to share the name. Such an entry matches a whole
              // segment EXACTLY (seat "Black" against a frame credited
              // "Black / Blue-White"), whereas a real mirror is a variant of
              // the primary color rather than one of its segments verbatim
              // ("Blue" against "Blue-White"). Only composite primaries get
              // this exception, so a plain single-axis primary keeps exactly
              // the behavior it always had.
              const primaryColors = primaryEntries.map(e => String(e.color || '').trim());
              const inheritedSegmentsLower = new Set();
              primaryColors.forEach(pc => {
                const segs = _colorSegments(pc);
                if (segs.length < 2) return; // not a chained/composite primary
                segs.forEach(s => inheritedSegmentsLower.add(s.toLowerCase()));
              });
              const independent = otherEntries.filter(e => {
                const colorLower = String(e.color || '').trim().toLowerCase();
                if (inheritedSegmentsLower.has(colorLower)) return true; // collision, keep as its own axis
                return !primaryColors.some(pc => _colorNamesMatch(pc, e.color));
              });

              // Each distinct axis among the independent entries must
              // contribute exactly one. Entries that DO carry a real
              // axisKey are grouped by it, so two DIFFERENT axes (e.g.
              // Mudguard + Rim) each contributing one entry combine safely
              // no matter how many total independent entries that adds up
              // to. An entry with NO axisKey at all (legacy data, or a
              // custom/free-form color with no real axis structure) has no
              // grouping info to disambiguate by at all — unlike a missing
              // axisKey meaning "this entry's own unique axis" (which would
              // wrongly treat 2 independent unstructured entries as always
              // safe to combine), every blank-axisKey entry shares ONE
              // pooled key, so a SINGLE such entry still combines (matches
              // the original 1-independent-entry case) but 2+ of them
              // collide and correctly fall back to the old per-entry
              // crediting — the exact "no stored cross-axis pairing to
              // tell which goes with which" case this whole function's
              // opening comment describes.
              const axisCounts = new Map();
              independent.forEach(e => {
                const key = String(e.axisKey || '').trim().toLowerCase() || '__no_axis_key__';
                axisCounts.set(key, (axisCounts.get(key) || 0) + 1);
              });
              const ambiguous = Array.from(axisCounts.values()).some(c => c > 1);

              if (!ambiguous) {
                // One composite bucket PER primary color, each carrying its
                // own primary qty. An independent axis holding a single
                // color for the whole lot (e.g. Rim = Black on all 40 units)
                // pairs with every primary color — which color goes with
                // which is not in question when that axis only has one.
                // Segment order is canonical (see _composeLotColorKey), not
                // Color Breakdown array order — otherwise two lots of the
                // same product credit two differently-ordered buckets and
                // split its stock.
                const axisOrder = axisOrderByProcess[processId.toLowerCase()];
                primaryEntries.forEach(pe => {
                  creditColor(_composeLotColorKey(pe, independent, axisOrder), pe.qty);
                });
                combined = true;
              }
            }

            if (!combined) {
              colorBreakdown.forEach(entry => creditColor(String(entry.color || '').trim(), entry.qty));
            }
          } else {
            const qty = Number(row[PRODUCTION_COL.QTY - 1]) || 0;
            getBucket(outputItemName, processId, productTag, '').producedQty += qty;
          }
        });

        // Pass 2: debit POOL-sourced components consumed by Completed lots
        // from the (untagged, intermediate) bucket of the upstream item. A
        // component scoped to a specific color (colorGroup other than
        // COMMON) debits that color's bucket specifically; a COMMON
        // component debits the blank-color bucket (the right bucket for an
        // upstream process that isn't itself multi-color).
        //
        // Only built if at least one component actually carries a non-blank
        // Unit (see PROCESS_COMPONENTS_COL.UNIT) — most recipes still have
        // none, so this stays a no-op cost in the common case. Mirrors
        // module_stock.js#_getBilledAndConsumedQtyMaps's identical handling
        // for ITEM-sourced components — a POOL-sourced row's Unit was being
        // silently ignored here, understating pool consumption by whatever
        // that row's conversion factor is (e.g. a "Dozen" row debiting as if
        // it were 1 Pcs).
        let poolItemUnitMap = null;
        let poolUnitsMap = null;
        data.forEach(row => {
          const status = String(row[PRODUCTION_COL.STATUS - 1] || '').trim().toLowerCase();
          if (status !== 'completed') return;

          const rawComponents = String(row[PRODUCTION_COL.COMPONENTS_CONSUMED - 1] || '').trim();
          if (!rawComponents) return;

          let components = [];
          try {
            const parsed = JSON.parse(rawComponents);
            if (Array.isArray(parsed)) components = parsed;
          } catch (e) {
            return;
          }

          components.forEach(comp => {
            const sourceType = String(comp.sourceType || '').trim().toUpperCase();
            if (sourceType !== COMPONENT_SOURCE_TYPES.POOL) return;
            const itemName = String(comp.itemName || '').trim();
            if (!itemName) return;
            let qty = Number(comp.qty) || 0;

            // Blank unit means "already in the pool item's Base Unit" —
            // preserves old behavior exactly for every pre-existing recipe row.
            const unit = String(comp.unit || '').trim();
            if (unit && typeof convertQtyToBaseUnit === 'function') {
              if (!poolItemUnitMap) poolItemUnitMap = typeof _getItemUnitInfoMap === 'function' ? _getItemUnitInfoMap() : {};
              if (!poolUnitsMap) poolUnitsMap = typeof _getUnitsMap === 'function' ? _getUnitsMap() : {};
              const unitInfo = typeof _lookupItemUnitInfo === 'function'
                ? _lookupItemUnitInfo(poolItemUnitMap, itemName, '')
                : { baseUnit: 'Pcs', purchaseUnit: 'Pcs', weightPerBaseUnit: 0 };
              try {
                qty = convertQtyToBaseUnit(qty, unit, unitInfo, poolUnitsMap);
              } catch (e) {
                // Unconvertible (e.g. no Weight-per-Base-Unit set yet) — fall
                // back to the as-entered qty rather than blocking the whole
                // Warehouse Pool recalculation over one bad recipe row.
              }
            }

            const colorGroup = String(comp.colorGroup || '').trim();
            let color = colorGroup && !isCommonColorGroup(colorGroup) ? colorGroup : '';

            // See _resolveCompositeColorToken — a manually-configured single
            // -token Color Sub-Group can legitimately refer to one part of a
            // composite bucket credited under Pass 1 above; resolve it to
            // that bucket's real key when unambiguous, rather than debiting
            // a phantom single-token bucket that was never credited.
            if (color && !buckets[_poolKey(itemName, '', color.toLowerCase())]) {
              const itemNameLower = itemName.toLowerCase();
              const candidates = Object.values(buckets)
                .filter(b => b.outputItemName.toLowerCase() === itemNameLower && !b.productTag);

              // A consumption recorded before _composeLotColorKey imposed a
              // canonical segment order names the same combination in a
              // different order ("Blue-White / Grey / Black" for what is now
              // credited as "Blue-White / Black / Grey"). Match on the
              // order-independent identity first, so historical rows debit
              // the real bucket instead of opening a phantom negative one.
              // Only when exactly one bucket carries that segment set —
              // otherwise there is nothing to disambiguate with.
              const wantOrderKey = _colorOrderKey(color);
              const orderMatches = candidates.filter(b => _colorOrderKey(b.color) === wantOrderKey);
              if (orderMatches.length === 1) {
                color = orderMatches[0].color;
              } else {
                const resolved = _resolveCompositeColorToken(
                  candidates.map(b => b.color.toLowerCase()), color.toLowerCase());
                if (resolved) color = resolved;
              }
            }

            const bucket = getBucket(itemName, '', '', color);
            bucket.consumedQty += qty;
          });
        });
      }
    }

    // Pass 3: debit finished-goods buckets by Dispatch quantity. A Product-
    // tagged bucket is matched by its tag; an untagged final-stage bucket
    // has no tag to match, so Dispatch's "Product ID" for that lot is the
    // Output Item Name itself (see _computeReadyToDispatchMap in
    // module_dispatch.js) — fall back to matching on that, restricted to
    // final-stage buckets so an untagged intermediate-WIP bucket sharing the
    // same Output Item Name from a non-final process is never touched.
    let dispatchSheet;
    try {
      dispatchSheet = getSheet(APP_CONFIG.SHEETS.DISPATCH);
    } catch (e) {
      dispatchSheet = null;
    }

    if (dispatchSheet) {
      const dLastRow = dispatchSheet.getLastRow();
      if (dLastRow >= 2) {
        const dData = dispatchSheet.getRange(2, 1, dLastRow - 1, DISPATCH_COL.QTY).getValues();
        const dispatchQtyByKey = {}; // productId(tag or outputItemName)Lower -> total dispatched qty
        dData.forEach(row => {
          const productId = String(row[DISPATCH_COL.PRODUCT_ID - 1] || '').trim();
          if (!productId) return;
          const qty = Number(row[DISPATCH_COL.QTY - 1]) || 0;
          const key = productId.toLowerCase();
          dispatchQtyByKey[key] = (dispatchQtyByKey[key] || 0) + qty;
        });

        const finalStageIds = new Set(
          (typeof _getAllProcessRecords === 'function' ? _getAllProcessRecords() : [])
            .filter(p => p.isFinalStage)
            .map(p => p.processId.toLowerCase())
        );

        // Dispatch carries no color of its own, so a Product Tag (or
        // untagged Output Item Name) credited across multiple color buckets
        // (a multi-color final-stage lot) can't be debited by color —
        // greedily drain whichever color buckets have stock first, dumping
        // any leftover (over-dispatch beyond total availability) on the
        // first bucket so the total consumedQty across all matching buckets
        // still equals the total dispatched qty.
        Object.keys(dispatchQtyByKey).forEach(key => {
          let remaining = dispatchQtyByKey[key];
          let matchingBuckets = Object.values(buckets).filter(b => b.productTag && b.productTag.toLowerCase() === key);
          if (matchingBuckets.length === 0) {
            matchingBuckets = Object.values(buckets).filter(b =>
              !b.productTag &&
              b.outputItemName.toLowerCase() === key &&
              b.processId && finalStageIds.has(b.processId.toLowerCase())
            );
          }
          if (matchingBuckets.length === 0) return;

          matchingBuckets.forEach(bucket => {
            if (remaining <= 0) return;
            const available = Math.max(bucket.producedQty - bucket.consumedQty, 0);
            const take = Math.min(remaining, available);
            bucket.consumedQty += take;
            remaining -= take;
          });

          if (remaining > 0) {
            matchingBuckets[0].consumedQty += remaining;
          }
        });
      }
    }

    const rows = Object.values(buckets)
      .filter(b => b.outputItemName)
      .map(b => [b.outputItemName, b.processId, b.productTag, b.producedQty, b.consumedQty, b.producedQty - b.consumedQty, b.color]);

    // Rewrite the sheet from scratch (small dataset — process count is tiny).
    const lastRow = poolSheet.getLastRow();
    if (lastRow >= 2) {
      poolSheet.getRange(2, 1, lastRow - 1, 7).clearContent();
    }
    if (rows.length > 0) {
      poolSheet.getRange(2, 1, rows.length, 7).setValues(rows);
    }

    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Warehouse Pool recalculated.');
  } catch (error) {
    Log.error('[recalculateWarehousePool] Error:', error.message);
    logAction('ERROR', 'recalculateWarehousePool', 'N/A', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to recalculate Warehouse Pool: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Public read endpoint for the Warehouse Pool view.
 */
function getWarehousePoolData() {
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
    } catch (e) {
      initWarehousePoolSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
    }
    ensureWarehousePoolColorColumn(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, []);

    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const records = data
      .map((row, i) => ({
        rowIdx: i + 2,
        outputItemName: String(row[WAREHOUSE_POOL_COL.OUTPUT_ITEM_NAME - 1] || '').trim(),
        processId: String(row[WAREHOUSE_POOL_COL.PROCESS_ID - 1] || '').trim(),
        productTag: String(row[WAREHOUSE_POOL_COL.PRODUCT_TAG - 1] || '').trim(),
        producedQty: Number(row[WAREHOUSE_POOL_COL.PRODUCED_QTY - 1]) || 0,
        consumedQty: Number(row[WAREHOUSE_POOL_COL.CONSUMED_QTY - 1]) || 0,
        availableQty: Number(row[WAREHOUSE_POOL_COL.AVAILABLE_QTY - 1]) || 0,
        color: String(row[WAREHOUSE_POOL_COL.COLOR - 1] || '').trim()
      }))
      .filter(r => r.outputItemName);

    records.sort((a, b) => a.outputItemName.localeCompare(b.outputItemName) || a.color.localeCompare(b.color));

    return buildResponse(true, records);
  } catch (error) {
    Log.error('[getWarehousePoolData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load Warehouse Pool data: ' + error.message);
  }
}

/**
 * Returns a lowercased-item-name -> availability map for every untagged
 * (intermediate WIP) Warehouse Pool bucket, built from a single batch read
 * of the sheet. Use this instead of calling getPoolAvailableQty() in a loop
 * — each call to that function re-reads the whole sheet, so looking up N
 * items one at a time costs N full-sheet reads instead of 1.
 *
 * Each entry's `total` sums every color bucket for that item (the qty a
 * COMMON-scoped component may draw, since it doesn't care which color);
 * `byColor[colorLower]` is that one color's own bucket (the qty a
 * component scoped to a specific Color Master name may draw), and
 * `byColor['']` is the blank/color-agnostic bucket.
 * @returns {Object} { [itemNameLower]: { total: number, byColor: { [colorLower]: number } } }
 */
function getPoolAvailableQtyMap() {
  const map = {};
  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.WAREHOUSE_POOL);
    ensureWarehousePoolColorColumn(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return map;

    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    data.forEach(row => {
      const itemName = String(row[WAREHOUSE_POOL_COL.OUTPUT_ITEM_NAME - 1] || '').trim().toLowerCase();
      const productTag = String(row[WAREHOUSE_POOL_COL.PRODUCT_TAG - 1] || '').trim();
      if (!itemName || productTag) return;
      const color = String(row[WAREHOUSE_POOL_COL.COLOR - 1] || '').trim().toLowerCase();
      const availableQty = Number(row[WAREHOUSE_POOL_COL.AVAILABLE_QTY - 1]) || 0;

      if (!map[itemName]) map[itemName] = { total: 0, byColor: {} };
      map[itemName].total += availableQty;
      map[itemName].byColor[color] = (map[itemName].byColor[color] || 0) + availableQty;
    });

    // See _resolveCompositeColorToken — mirror the same single-token ->
    // composite-bucket fallback here so a pre-save availability check
    // (_validatePoolAvailability in module_production.js, which reads
    // byColor[colorGroup]) agrees with what recalculateWarehousePool's
    // Pass 2 will actually debit once the lot completes, instead of
    // reporting a token-scoped need as unavailable when real (composite-
    // keyed) stock for it exists.
    Object.values(map).forEach(entry => {
      const canonicalColors = Object.keys(entry.byColor);
      const tokenSources = new Map(); // token -> Set of composite source colors
      canonicalColors.forEach(c => {
        if (c.indexOf(COLOR_COMBO_DELIMITER) === -1) return;
        c.split(COLOR_COMBO_DELIMITER).forEach(t => {
          const token = t.trim();
          if (!token) return;
          if (!tokenSources.has(token)) tokenSources.set(token, new Set());
          tokenSources.get(token).add(c);
        });
      });
      tokenSources.forEach((sources, token) => {
        if (entry.byColor.hasOwnProperty(token) || sources.size !== 1) return;
        entry.byColor[token] = entry.byColor[Array.from(sources)[0]];
      });
    });

    return map;
  } catch (e) {
    return map;
  }
}

/**
 * Warns (never blocks) when removing a Completed lot's own credit to the
 * Warehouse Pool — via un-completing its status or deleting it outright —
 * would leave a bucket negative, i.e. a downstream lot already consumed
 * this credit. Mirrors the informational-only pattern _validatePoolAvailability
 * already uses for the opposite direction (a lot's own POOL-sourced
 * consumption); this is intentionally non-blocking, matching the "allow
 * negative pool/stock so operations aren't blocked" exception.
 * @param {string} outputItemName
 * @param {Array<{color:string, qty:number}>|null} colorBreakdown - null/empty for a flat (non-color) lot
 * @param {number} flatQty - used when colorBreakdown is empty
 * @returns {string|null} warning message, or null if nothing would go negative
 */
function _checkPoolCreditRemovalWarning(outputItemName, colorBreakdown, flatQty) {
  const name = String(outputItemName || '').trim();
  if (!name) return null;

  const poolMap = getPoolAvailableQtyMap();
  const entry = poolMap[name.toLowerCase()];
  if (!entry) return null;

  const credits = (colorBreakdown && colorBreakdown.length > 0)
    ? colorBreakdown.map(e => ({ color: String(e.color || '').trim(), qty: Number(e.qty) || 0 }))
    : [{ color: '', qty: Number(flatQty) || 0 }];

  const shortfalls = [];
  credits.forEach(c => {
    if (c.qty <= 0) return;
    const currentAvailable = entry.byColor[c.color.toLowerCase()] || 0;
    const wouldBe = currentAvailable - c.qty;
    if (wouldBe < -0.0001) {
      shortfalls.push(`"${name}"${c.color ? ' (' + c.color + ')' : ''}: ${wouldBe.toFixed(2)}`);
    }
  });

  if (shortfalls.length === 0) return null;
  return `Warning: this leaves the Warehouse Pool negative for ${shortfalls.join(', ')} — a downstream lot already consumed this credit. The pool balance will show negative until corrected.`;
}

/**
 * Returns the current total available qty (summed across every color
 * bucket) for an untagged (intermediate WIP) Warehouse Pool item.
 *
 * Looking up a single item is fine, but if you need availability for
 * several items (e.g. every POOL-sourced component in a recipe), call
 * getPoolAvailableQtyMap() once instead — each call here is a full-sheet
 * read.
 * @param {string} outputItemName
 * @returns {number}
 */
function getPoolAvailableQty(outputItemName) {
  const target = String(outputItemName || '').trim().toLowerCase();
  if (!target) return 0;
  const entry = getPoolAvailableQtyMap()[target];
  return entry ? entry.total : 0;
}

// ─────────────────────────────────────────────────────────────────────────
// PRODUCTION COLOR-CHAIN VERIFICATION
// ─────────────────────────────────────────────────────────────────────────

/**
 * Read-only audit of the color identity every Production lot carries
 * through the process chain. Writes nothing — run it any time to answer
 * "is my composite color data self-consistent, and is any product's stock
 * silently split across two buckets?".
 *
 * A composite pool color (see COLOR_COMBO_DELIMITER, and Pass 1 of
 * recalculateWarehousePool) is a chain: its first segment is the producing
 * lot's primary-axis color — itself possibly a composite inherited from
 * upstream — followed by one segment per independent axis that lot combined
 * in. This walks that structure back and reports where it fails to hold up:
 *
 *  - order-split       Two or more live buckets for the SAME item whose
 *                      segments are the same set in a different order (e.g.
 *                      "Blue-White / Black / Grey" vs "Blue-White / Grey /
 *                      Black"). One real product, stock halved across two
 *                      rows. Since _composeLotColorKey imposed a canonical
 *                      order this can only be historical data, and a
 *                      recalculation heals it — this finding names what
 *                      will merge.
 *  - unknown-axis      A lot's Color Breakdown entry carries an axisKey that
 *                      is not an axis of its own process (renamed/removed
 *                      axis, or a lot moved between processes).
 *  - color-off-axis    An entry's color is not one its process's axes
 *                      actually offer, and it is not flagged isCustom —
 *                      usually a Color Master rename whose cascade did not
 *                      reach this lot.
 *  - orphan-segment    A live composite bucket contains a segment that no
 *                      axis of its producing process can account for.
 *  - unresolved-debit  A lot consumes a POOL component under a color that
 *                      matches no bucket of that item, so the consumption
 *                      opens a phantom negative bucket instead of debiting
 *                      real stock.
 *
 * Cost: one Production read plus one recipe read per distinct process, so
 * it is safe to run against a full sheet. Findings are returned AND written
 * to the execution log, so this is equally usable from the Apps Script
 * editor (see _runVerifyColorChain) and from a UI caller later.
 *
 * @param {Object} [options]
 * @param {number} [options.limit] Max findings RETURNED per category (all are counted). Default 50.
 * @returns {Object} buildResponse with { lotsChecked, bucketsChecked, countsByType, findings }
 */
function verifyProductionColorChain(options) {
  try {
    const limit = (options && Number(options.limit)) || 50;
    const findings = [];
    const countsByType = {};
    function addFinding(type, message, detail) {
      countsByType[type] = (countsByType[type] || 0) + 1;
      if (countsByType[type] <= limit) {
        findings.push({ type: type, message: message, detail: detail || {} });
      }
    }

    const poolRows = (getWarehousePoolData().data || []);

    // ── order-split: same item, same segment set, different segment order.
    const byItemAndOrderKey = new Map();
    poolRows.forEach(r => {
      if (!r.color || _colorSegments(r.color).length < 2) return;
      const key = r.outputItemName.trim().toLowerCase() + '||' + _colorOrderKey(r.color);
      if (!byItemAndOrderKey.has(key)) byItemAndOrderKey.set(key, []);
      byItemAndOrderKey.get(key).push(r);
    });
    byItemAndOrderKey.forEach(rows => {
      const distinct = Array.from(new Set(rows.map(r => r.color)));
      if (distinct.length < 2) return;
      addFinding('order-split',
        '"' + rows[0].outputItemName + '" has ' + distinct.length +
        ' buckets that are the same color combination in a different order: ' +
        distinct.map(c => '"' + c + '"').join(' vs ') +
        '. Their stock is split; a Warehouse Pool recalculation merges them.',
        { outputItemName: rows[0].outputItemName, colors: distinct,
          qtys: rows.map(r => ({ color: r.color, available: r.availableQty })) });
    });

    // ── per-lot checks against each process's real axes.
    const sheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
    const lastRow = sheet.getLastRow();
    const colorLinks = typeof _getAllProcessColorLinks === 'function' ? _getAllProcessColorLinks() : [];
    const axesCache = new Map(); // processIdLower -> { axes, byKey, allColorsLower }
    function axesFor(processId) {
      const key = String(processId || '').trim().toLowerCase();
      if (!axesCache.has(key)) {
        let axes = [];
        try {
          const comps = (getProcessComponentsData(processId).data || []);
          axes = computeColorAxesForProcess(processId, comps, poolRows, colorLinks) || [];
        } catch (e) {
          axes = [];
        }
        const byKey = new Map();
        const allColorsLower = new Set();
        axes.forEach(a => {
          byKey.set(String(a.key || '').toLowerCase(), a);
          (a.colors || []).forEach(c => allColorsLower.add(String(c).trim().toLowerCase()));
        });
        axesCache.set(key, { axes: axes, byKey: byKey, allColorsLower: allColorsLower });
      }
      return axesCache.get(key);
    }

    let lotsChecked = 0;
    let prodData = [];
    if (lastRow >= 2) {
      prodData = sheet.getRange(2, 1, lastRow - 1, PRODUCTION_COL.COLOR_BREAKDOWN).getValues();
      prodData.forEach((row, i) => {
        const rowIdx = i + 2;
        const status = String(row[PRODUCTION_COL.STATUS - 1] || '').trim().toLowerCase();
        if (status !== 'completed') return;
        const processId = String(row[PRODUCTION_COL.PROCESS_ID - 1] || '').trim();
        const lotNumber = String(row[PRODUCTION_COL.LOT_NUMBER - 1] || '').trim() || ('row ' + rowIdx);

        let breakdown = [];
        try {
          const parsed = JSON.parse(String(row[PRODUCTION_COL.COLOR_BREAKDOWN - 1] || '').trim() || '[]');
          if (Array.isArray(parsed)) breakdown = parsed;
        } catch (e) {
          return;
        }
        if (breakdown.length === 0) return;
        lotsChecked++;

        const info = axesFor(processId);
        if (info.axes.length === 0) return; // no color dimension to check against

        breakdown.forEach(e => {
          const color = String((e && e.color) || '').trim();
          if (!color) return;
          const axisKey = String((e && e.axisKey) || '').trim().toLowerCase();
          // A blank axisKey is legacy data, not an error — those lots simply
          // predate axis tagging.
          if (axisKey && axisKey !== 'other' && axisKey !== 'custom' && !info.byKey.has(axisKey)) {
            addFinding('unknown-axis',
              'Lot ' + lotNumber + ' records color "' + color + '" under axis "' + axisKey +
              '", which its process no longer has.',
              { rowIdx: rowIdx, lotNumber: lotNumber, processId: processId, color: color, axisKey: axisKey });
            return;
          }
          if (e && e.isCustom) return; // operator-typed one-off, exempt by design
          if (!info.allColorsLower.has(color.toLowerCase())) {
            addFinding('color-off-axis',
              'Lot ' + lotNumber + ' records color "' + color +
              '", which none of its process\'s axes offer.',
              { rowIdx: rowIdx, lotNumber: lotNumber, processId: processId, color: color, axisKey: axisKey });
          }
        });
      });
    }

    // ── orphan segments in live composite buckets.
    let bucketsChecked = 0;
    poolRows.forEach(r => {
      const segments = _colorSegments(r.color);
      if (segments.length < 2) return;
      bucketsChecked++;
      const info = axesFor(r.processId);
      if (info.axes.length === 0) return;
      segments.forEach(seg => {
        if (!info.allColorsLower.has(seg.toLowerCase())) {
          addFinding('orphan-segment',
            'Bucket "' + r.color + '" of "' + r.outputItemName + '" contains segment "' + seg +
            '", which no axis of its producing process accounts for.',
            { outputItemName: r.outputItemName, color: r.color, segment: seg, processId: r.processId });
        }
      });
    });

    // ── phantom buckets: a color that was only ever CONSUMED, never
    // produced. Detected from the rebuilt pool itself rather than by
    // re-deriving Pass 2's lookup, because an unresolved consumption
    // creates the very bucket it failed to find — checking "does a bucket
    // with this name exist" would therefore always say yes and hide the
    // problem.
    //
    // Deliberately NOT a negative-stock check: a bucket that was genuinely
    // produced and has gone negative is a real physical counting signal and
    // is left alone. Only producedQty === 0 with consumption against it is
    // reported, which means the color was never produced under that name at
    // all — a reference problem (a rename that missed, or a recipe Color
    // Sub-Group naming a combination that does not exist), not a count.
    const consumersByItemColor = new Map(); // itemLower||colorLower -> [lotNumber]
    prodData.forEach((row, i) => {
      const status = String(row[PRODUCTION_COL.STATUS - 1] || '').trim().toLowerCase();
      if (status !== 'completed') return;
      const lotNumber = String(row[PRODUCTION_COL.LOT_NUMBER - 1] || '').trim() || ('row ' + (i + 2));
      let comps = [];
      try {
        const parsed = JSON.parse(String(row[PRODUCTION_COL.COMPONENTS_CONSUMED - 1] || '').trim() || '[]');
        if (Array.isArray(parsed)) comps = parsed;
      } catch (e) {
        return;
      }
      comps.forEach(c => {
        if (String((c && c.sourceType) || '').toUpperCase() !== COMPONENT_SOURCE_TYPES.POOL) return;
        const itemName = String((c && c.itemName) || '').trim();
        const cg = String((c && c.colorGroup) || '').trim();
        if (!itemName || !cg || isCommonColorGroup(cg)) return;
        const k = itemName.toLowerCase() + '||' + cg.toLowerCase();
        if (!consumersByItemColor.has(k)) consumersByItemColor.set(k, []);
        consumersByItemColor.get(k).push(lotNumber);
      });
    });
    poolRows.forEach(r => {
      if (!r.color) return;
      if (!(Number(r.producedQty) === 0 && Number(r.consumedQty) > 0)) return;
      const lots = consumersByItemColor.get(r.outputItemName.trim().toLowerCase() + '||' + r.color.toLowerCase()) || [];
      addFinding('unresolved-debit',
        '"' + r.outputItemName + '" is consumed in color "' + r.color + '" (' + r.consumedQty +
        ') but was never produced in it' +
        (lots.length ? ' — see lot(s) ' + lots.slice(0, 5).join(', ') : '') +
        '. That consumption is drawing on a color that does not exist rather than on real stock.',
        { outputItemName: r.outputItemName, color: r.color, consumedQty: r.consumedQty, lotNumbers: lots });
    });

    const total = Object.keys(countsByType).reduce((s, k) => s + countsByType[k], 0);
    Logger.log('[VERIFY COLOR CHAIN] lots checked: ' + lotsChecked +
      ', composite buckets checked: ' + bucketsChecked + ', findings: ' + total);
    if (total === 0) {
      Logger.log('[VERIFY COLOR CHAIN] No problems found - every lot color chain is consistent.');
    } else {
      Object.keys(countsByType).sort().forEach(t => Logger.log('  ' + t + ': ' + countsByType[t]));
      findings.forEach(f => Logger.log('  [' + f.type + '] ' + f.message));
      if (total > findings.length) {
        Logger.log('  ... ' + (total - findings.length) + ' more (raise options.limit to see them).');
      }
    }

    return buildResponse(true, {
      lotsChecked: lotsChecked,
      bucketsChecked: bucketsChecked,
      countsByType: countsByType,
      findings: findings
    }, total === 0 ? 'No color-chain problems found.' : total + ' color-chain finding(s).');
  } catch (error) {
    Log.error('[verifyProductionColorChain] Error:', error.message);
    return buildResponse(false, null, 'Failed to verify production color chain: ' + error.message);
  }
}

/**
 * One-click entry point for the Apps Script editor: pick
 * _runVerifyColorChain from the function dropdown and Run, then read the
 * execution log. Writes nothing.
 */
function _runVerifyColorChain() {
  verifyProductionColorChain();
}
