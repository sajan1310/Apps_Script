/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_contractors.gs — CONTRACTORS & CONTRACTOR RATE CARD MODULE
 *
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Master data for independent contractors who run production processes
 * (paid per unit at a negotiated rate) and, via the hardcoded
 * LOGISTICS_PROCESS_NAME virtual category, logistics/dispatch contractors
 * too. Production lots (module_production.js) and Dispatch records
 * (module_dispatch.js) snapshot a rate from here at save time; BOM cost
 * lines (module_bom.js) can pull a rate from here as a starting point for
 * their own (possibly different) negotiated cost line.
 *
 * Sheet Layout (Contractors):
 * ───────────────────────────────────────────────────────────────────────────
 * Col A (1):   Contractor Name (unique identifier)
 * Col B (2):   Contact Number
 * Col C (3):   Address
 * Col D (4):   GST/PAN
 * Col E (5):   Remarks
 *
 * Sheet Layout (Contractor Rates — the rate card):
 * ───────────────────────────────────────────────────────────────────────────
 * Col A (1):   Contractor Name
 * Col B (2):   Process Name (free string — active Process Master name, or
 *              the LOGISTICS_PROCESS_NAME virtual category)
 * Col C (3):   Rate Per Unit
 * Col D (4):   Remarks
 * ═══════════════════════════════════════════════════════════════════════════
 */

const CONTRACTOR_LOCK_TIMEOUT_MS = 15000;

// ── Contractors master ──────────────────────────────────────────────────

function initContractorsSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.CONTRACTORS);
    if (!sheet) {
      sheet = ss.insertSheet(APP_CONFIG.SHEETS.CONTRACTORS);
    }
    sheet.getRange(1, 1, 1, 5)
      .setValues([['Contractor Name', 'Contact Number', 'Address', 'GST/PAN', 'Remarks']])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Contractors sheet initialized successfully.');
  } catch (error) {
    Log.error('[initContractorsSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize Contractors sheet: ' + error.message);
  }
}

function getContractorsData() {
  return getCachedListResponse(MASTER_DATA_CACHE_KEYS.CONTRACTORS, () => {
    try {
      let sheet;
      try {
        sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTORS);
      } catch (e) {
        initContractorsSheet();
        sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTORS);
      }

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return buildResponse(true, []);

      const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      const contractors = data.map(r => ({
        name: String(r[CONTRACTORS_COL.CONTRACTOR_NAME - 1] || '').trim(),
        contact: String(r[CONTRACTORS_COL.CONTACT - 1] || '').trim(),
        address: String(r[CONTRACTORS_COL.ADDRESS - 1] || '').trim(),
        gstPan: String(r[CONTRACTORS_COL.GST_PAN - 1] || '').trim(),
        remarks: String(r[CONTRACTORS_COL.REMARKS - 1] || '').trim()
      })).filter(c => c.name !== '');

      contractors.sort((a, b) => a.name.localeCompare(b.name));

      return buildResponse(true, contractors);
    } catch (error) {
      Log.error('[getContractorsData] Error:', error.message);
      return buildResponse(false, null, 'Failed to load contractors: ' + error.message);
    }
  });
}

function saveContractor(formData) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CONTRACTOR_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTORS);
    } catch (e) {
      initContractorsSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTORS);
    }

    const newName = sanitizeString(formData.contractorName, 'contractorName');
    if (!newName) {
      return buildResponse(false, null, 'Contractor name cannot be empty.');
    }

    const isEdit = !!formData.originalContractorName;
    const originalName = isEdit ? sanitizeString(formData.originalContractorName, 'originalName') : newName;

    const data = sheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === originalName.toLowerCase()) {
        targetRow = i + 1;
        break;
      }
    }

    if (isEdit && targetRow === -1) {
      return buildResponse(false, null, 'Original contractor not found in sheet.');
    }
    if (!isEdit && targetRow !== -1) {
      return buildResponse(false, null, 'A contractor with this name already exists.');
    }

    if (isEdit && newName.toLowerCase() !== originalName.toLowerCase()) {
      const dupeCheck = data.findIndex((r, idx) => idx > 0 && String(r[0]).trim().toLowerCase() === newName.toLowerCase());
      if (dupeCheck !== -1) {
        return buildResponse(false, null, 'Another contractor with this new name already exists.');
      }
    }

    const rowData = [
      newName,
      sanitizeString(formData.contact || '', 'contact'),
      sanitizeString(formData.address || '', 'address'),
      sanitizeString(formData.gstPan || '', 'gstPan'),
      sanitizeString(formData.remarks || '', 'remarks')
    ];

    if (isEdit) {
      sheet.getRange(targetRow, 1, 1, 5).setValues([rowData]);

      // Keep every de-normalized snapshot of this contractor's name in sync
      // if they were renamed (rate card, Production lots, Dispatch logistics,
      // and Contractor Payments all store the name as a plain string, not a
      // foreign key). Plain string compare (not case-insensitive) so a
      // casing-only edit still cascades — otherwise this row is the only
      // place that picks up the new casing.
      if (newName !== originalName) {
        _renameContractorEverywhere(originalName, newName);
      }
    } else {
      sheet.appendRow(rowData);
    }

    SpreadsheetApp.flush();
    invalidateListCache(MASTER_DATA_CACHE_KEYS.CONTRACTORS);

    const action = isEdit ? 'UPDATE_CONTRACTOR' : 'CREATE_CONTRACTOR';
    const msg = isEdit ? 'Contractor profile updated successfully.' : 'New contractor registered.';
    logAction(action, APP_CONFIG.SHEETS.CONTRACTORS, newName, JSON.stringify(formData), 'SUCCESS');

    return buildResponse(true, { name: newName }, msg);
  } catch (error) {
    Log.error('[saveContractor] Error:', error.message);
    logAction('SAVE_CONTRACTOR_ERROR', APP_CONFIG.SHEETS.CONTRACTORS, formData ? formData.contractorName : 'unknown', error.message, 'ERROR');
    return buildResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

function deleteContractor(contractorName) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CONTRACTOR_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System busy.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTORS);
    if (!sheet) throw new Error('Contractors sheet not found.');

    const targetName = String(contractorName || '').trim().toLowerCase();
    const data = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 1).getValues();
    let targetRow = -1;

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === targetName) {
        targetRow = i + 2;
        break;
      }
    }

    if (targetRow === -1) throw new Error('Contractor not found.');

    const inUseMsg = _contractorInUseMessage(contractorName);
    if (inUseMsg) return buildResponse(false, null, inUseMsg);

    sheet.deleteRow(targetRow);

    // Also remove this contractor's rate card rows
    try {
      const ratesSheet = getSheet(APP_CONFIG.SHEETS.CONTRACTOR_RATES);
      deleteRowsById(contractorName, ratesSheet, 2, CONTRACTOR_RATES_COL.CONTRACTOR_NAME);
    } catch (e) {
      // Rate card sheet doesn't exist yet, nothing to clean up
    }

    SpreadsheetApp.flush();
    invalidateListCache(MASTER_DATA_CACHE_KEYS.CONTRACTORS);

    logAction('DELETE_CONTRACTOR', APP_CONFIG.SHEETS.CONTRACTORS, contractorName, 'Contractor deleted', 'SUCCESS');

    return buildResponse(true, null, `Contractor "${contractorName}" deleted.`);
  } catch (error) {
    logAction('DELETE_CONTRACTOR_ERROR', APP_CONFIG.SHEETS.CONTRACTORS, contractorName, error.message, 'ERROR');
    return buildResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

function deleteContractorsBulk(contractorNames) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CONTRACTOR_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System busy.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTORS);
    if (!sheet) throw new Error('Contractors sheet not found.');

    const requestedNames = (contractorNames || []).map(n => String(n || '').trim()).filter(Boolean);
    if (requestedNames.length === 0) {
      return buildResponse(true, null, 'No contractors selected.');
    }

    // Skip any contractor with Production/Dispatch/Payment history — deleting
    // them would orphan that history under a name no longer in the master
    // list, silently hiding their balance due (see _contractorInUseMessage).
    // Each referencing sheet is read ONCE here into a name Set, instead of
    // _contractorInUseMessage() re-reading all 3 sheets from scratch for
    // EVERY requested name (O(K x 3N) instead of O(3N + K)) — same pattern
    // as deleteClientOrdersBulk. The bulk path only needs the yes/no
    // in-use verdict, not _contractorInUseMessage's per-sheet reason text.
    const inUseNames = new Set();
    const addNamesFromSheet = (sheetName, col) => {
      let sh;
      try { sh = getSheet(sheetName); } catch (e) { return; }
      const lastRow = sh.getLastRow();
      if (lastRow < 2) return;
      sh.getRange(2, col, lastRow - 1, 1).getValues().forEach(row => {
        const n = String(row[0] || '').trim().toLowerCase();
        if (n) inUseNames.add(n);
      });
    };
    addNamesFromSheet(APP_CONFIG.SHEETS.PRODUCTION, PRODUCTION_COL.ASSIGNED_TO);
    addNamesFromSheet(APP_CONFIG.SHEETS.DISPATCH, DISPATCH_COL.LOGISTICS_CONTRACTOR);
    addNamesFromSheet(APP_CONFIG.SHEETS.CONTRACTOR_PAYMENTS, CONTRACTOR_PAYMENTS_COL.CONTRACTOR_NAME);

    const skipped = requestedNames.filter(n => inUseNames.has(n.toLowerCase()));
    const deletableNames = requestedNames.filter(n => !inUseNames.has(n.toLowerCase()));
    const targetSet = new Set(deletableNames.map(n => n.toLowerCase()));

    const lastRow = sheet.getLastRow();
    if (lastRow < 2 || targetSet.size === 0) {
      const msg = skipped.length
        ? `No contractors deleted — all selected have Production, Dispatch, or payment history: ${skipped.join(', ')}.`
        : 'No contractors to delete.';
      return buildResponse(skipped.length === 0, null, msg);
    }

    const { rowsDeleted } = rewriteSheetExcludingRows(sheet, 2, row => {
      const name = String(row[CONTRACTORS_COL.CONTRACTOR_NAME - 1] || '').trim().toLowerCase();
      return targetSet.has(name);
    });

    try {
      const ratesSheet = getSheet(APP_CONFIG.SHEETS.CONTRACTOR_RATES);
      if (ratesSheet) {
        rewriteSheetExcludingRows(ratesSheet, 2, row => {
          const name = String(row[CONTRACTOR_RATES_COL.CONTRACTOR_NAME - 1] || '').trim().toLowerCase();
          return targetSet.has(name);
        });
      }
    } catch (e) {
      // Rate card sheet doesn't exist yet, nothing to clean up
    }

    SpreadsheetApp.flush();
    invalidateListCache(MASTER_DATA_CACHE_KEYS.CONTRACTORS);

    let msg = `Deleted ${rowsDeleted} contractor(s).`;
    if (skipped.length) {
      msg += ` Skipped (has history): ${skipped.join(', ')}.`;
    }
    logAction('BULK_DELETE', APP_CONFIG.SHEETS.CONTRACTORS, 'multiple', msg, 'SUCCESS');

    return buildResponse(true, null, msg);
  } catch (error) {
    logAction('ERROR', 'deleteContractorsBulk', 'multiple', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete contractors: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

// ── Contractor Rate Card ────────────────────────────────────────────────

function initContractorRatesSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.CONTRACTOR_RATES);
    if (!sheet) {
      sheet = ss.insertSheet(APP_CONFIG.SHEETS.CONTRACTOR_RATES);
    }
    sheet.getRange(1, 1, 1, 4)
      .setValues([['Contractor Name', 'Process Name', 'Rate Per Unit', 'Remarks']])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Contractor Rates sheet initialized successfully.');
  } catch (error) {
    Log.error('[initContractorRatesSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize Contractor Rates sheet: ' + error.message);
  }
}

/**
 * Returns the full rate card, optionally filtered to one contractor.
 * @param {string} [contractorName]
 */
function getContractorRatesData(contractorName) {
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTOR_RATES);
    } catch (e) {
      initContractorRatesSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTOR_RATES);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, []);

    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const filterName = contractorName ? String(contractorName).trim().toLowerCase() : null;

    const rates = data.map((r, i) => ({
      rowIdx: i + 2,
      contractorName: String(r[CONTRACTOR_RATES_COL.CONTRACTOR_NAME - 1] || '').trim(),
      processName: String(r[CONTRACTOR_RATES_COL.PROCESS_NAME - 1] || '').trim(),
      ratePerUnit: Number(r[CONTRACTOR_RATES_COL.RATE_PER_UNIT - 1]) || 0,
      remarks: String(r[CONTRACTOR_RATES_COL.REMARKS - 1] || '').trim()
    })).filter(r => r.contractorName && r.processName &&
      (!filterName || r.contractorName.toLowerCase() === filterName));

    rates.sort((a, b) => a.contractorName.localeCompare(b.contractorName) || a.processName.localeCompare(b.processName));

    return buildResponse(true, rates);
  } catch (error) {
    Log.error('[getContractorRatesData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load contractor rates: ' + error.message);
  }
}

/**
 * Internal lookup: rate per unit for a (contractor, process) pair, or 0 if
 * no rate card entry exists. Case-insensitive match on both fields.
 * @private
 */
function _getContractorRate(contractorName, processName) {
  const cName = String(contractorName || '').trim().toLowerCase();
  const pName = String(processName || '').trim().toLowerCase();
  if (!cName || !pName) return 0;

  let sheet;
  try {
    sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTOR_RATES);
  } catch (e) {
    return 0;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  for (let i = 0; i < data.length; i++) {
    const rowContractor = String(data[i][CONTRACTOR_RATES_COL.CONTRACTOR_NAME - 1] || '').trim().toLowerCase();
    const rowProcess = String(data[i][CONTRACTOR_RATES_COL.PROCESS_NAME - 1] || '').trim().toLowerCase();
    if (rowContractor === cName && rowProcess === pName) {
      return Number(data[i][CONTRACTOR_RATES_COL.RATE_PER_UNIT - 1]) || 0;
    }
  }
  return 0;
}

/**
 * Reads Dispatch rows with a recorded Logistics Contractor + Cost, i.e. the
 * "Dispatch / Logistics" virtual process category's payable entries —
 * mirrors how Completed Production lots are read for the same ledger.
 * Optionally filtered to a single contractor (case-insensitive).
 * @param {string} [contractorNameFilter]
 * @returns {Array<{date, dateRaw, contractorName, qty, payable, dispatchNumber, productName}>}
 * @private
 */
function _getDispatchLogisticsPayableRows(contractorNameFilter) {
  const rows = [];

  let dispatchSheet;
  try {
    dispatchSheet = getSheet(APP_CONFIG.SHEETS.DISPATCH);
  } catch (e) {
    return rows; // Dispatch sheet not initialized yet
  }

  if (typeof ensureDispatchLogisticsColumns === 'function') {
    ensureDispatchLogisticsColumns(dispatchSheet);
  }

  const lastRow = dispatchSheet.getLastRow();
  if (lastRow < 2) return rows;

  const filter = String(contractorNameFilter || '').trim().toLowerCase();
  const data = dispatchSheet.getRange(2, 1, lastRow - 1, DISPATCH_COL.LOGISTICS_COST).getValues();

  data.forEach(row => {
    const contractorName = String(row[DISPATCH_COL.LOGISTICS_CONTRACTOR - 1] || '').trim();
    if (!contractorName) return;
    if (filter && contractorName.toLowerCase() !== filter) return;

    const payable = Number(row[DISPATCH_COL.LOGISTICS_COST - 1]) || 0;
    if (payable <= 0) return; // No rate card match — nothing owed

    const rawDate = row[DISPATCH_COL.DISPATCH_DATE - 1];
    rows.push({
      date: rawDate instanceof Date ? toSafeDateString(rawDate) : String(rawDate || ''),
      dateRaw: rawDate instanceof Date ? rawDate.toISOString() : null,
      contractorName: contractorName,
      qty: Number(row[DISPATCH_COL.QTY - 1]) || 0,
      payable: payable,
      dispatchNumber: String(row[DISPATCH_COL.DISPATCH_NUMBER - 1] || '').trim(),
      productName: String(row[DISPATCH_COL.PRODUCT_NAME - 1] || '').trim()
    });
  });

  return rows;
}

/**
 * Public read endpoint mirroring _getContractorRate, for client-side rate
 * lookups (e.g. showing a live payable estimate before save).
 * @param {string} contractorName
 * @param {string} processName
 */
function getContractorRateForProcess(contractorName, processName) {
  try {
    return buildResponse(true, { ratePerUnit: _getContractorRate(contractorName, processName) });
  } catch (error) {
    Log.error('[getContractorRateForProcess] Error:', error.message);
    return buildResponse(false, null, 'Failed to look up contractor rate: ' + error.message);
  }
}

/**
 * Creates or updates a rate card entry. Upserts by (contractorName,
 * processName) — saving an existing pair replaces its rate rather than
 * duplicating the row.
 */
function saveContractorRate(formData) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CONTRACTOR_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTOR_RATES);
    } catch (e) {
      initContractorRatesSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTOR_RATES);
    }

    const contractorName = sanitizeString(formData.contractorName, 'contractorName');
    const processName = sanitizeString(formData.processName, 'processName');
    if (!contractorName || !processName) {
      return buildResponse(false, null, 'Contractor and Process are both required.');
    }

    const ratePerUnit = validateNumber(formData.ratePerUnit, 0, 10000000);
    const remarks = sanitizeString(formData.remarks || '', 'remarks').slice(0, APP_CONFIG.VALIDATION.MAX_REMARKS_LENGTH);

    // "Create if needed" — a contractor typed fresh into a Select2 dropdown
    // (Process modal, BOM cost row, Production Assigned To, Dispatch
    // logistics) should just work without a separate trip to the
    // Contractors tab to register them first.
    _ensureContractorExists(contractorName);

    const lastRow = sheet.getLastRow();
    let targetRow = -1;
    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (let i = 0; i < data.length; i++) {
        const rowContractor = String(data[i][0] || '').trim().toLowerCase();
        const rowProcess = String(data[i][1] || '').trim().toLowerCase();
        if (rowContractor === contractorName.toLowerCase() && rowProcess === processName.toLowerCase()) {
          targetRow = i + 2;
          break;
        }
      }
    }

    const rowData = [contractorName, processName, ratePerUnit, remarks];
    if (targetRow !== -1) {
      sheet.getRange(targetRow, 1, 1, 4).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    SpreadsheetApp.flush();

    const msg = `Rate card updated: ${contractorName} / ${processName} = ${ratePerUnit}/unit.`;
    logAction(targetRow !== -1 ? 'UPDATE' : 'CREATE', APP_CONFIG.SHEETS.CONTRACTOR_RATES, contractorName, msg, 'SUCCESS');

    return buildResponse(true, null, 'Rate saved successfully.');
  } catch (error) {
    Log.error('[saveContractorRate] Error:', error.message);
    logAction('ERROR', 'saveContractorRate', formData.contractorName || 'unknown', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to save rate: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

function deleteContractorRate(contractorName, processName) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CONTRACTOR_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTOR_RATES);
    if (!sheet) throw new Error('Contractor Rates sheet not found.');

    const cName = String(contractorName || '').trim().toLowerCase();
    const pName = String(processName || '').trim().toLowerCase();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, null, 'No rates to delete.');

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      const rowContractor = String(data[i][0] || '').trim().toLowerCase();
      const rowProcess = String(data[i][1] || '').trim().toLowerCase();
      if (rowContractor === cName && rowProcess === pName) {
        sheet.deleteRow(i + 2);
        SpreadsheetApp.flush();
        logAction('DELETE', APP_CONFIG.SHEETS.CONTRACTOR_RATES, contractorName, `Rate removed for ${processName}`, 'SUCCESS');
        return buildResponse(true, null, 'Rate deleted successfully.');
      }
    }

    return buildResponse(false, null, 'Rate card entry not found.');
  } catch (error) {
    Log.error('[deleteContractorRate] Error:', error.message);
    logAction('ERROR', 'deleteContractorRate', contractorName, error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete rate: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Returns a human-readable reason the contractor can't be deleted, or null
 * if they're clear. Deletion is blocked once Production, Dispatch, or
 * Contractor Payments hold a row under this name — those sheets aren't
 * cleaned up on delete (only the rate card is), so removing the master
 * record would orphan that history: getContractorLedgerData/
 * getContractorAccountLedger would still compute a balance for them, but
 * the Contractors tab would have nowhere to show it since it only lists
 * names from this master sheet. Same "blocked if referenced" pattern as
 * deleteProcess.
 * @private
 */
function _contractorInUseMessage(contractorName) {
  const name = String(contractorName || '').trim().toLowerCase();
  if (!name) return null;

  const hasMatch = (sheetName, col) => {
    let sheet;
    try {
      sheet = getSheet(sheetName);
    } catch (e) {
      return false;
    }
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;
    const values = sheet.getRange(2, col, lastRow - 1, 1).getValues();
    return values.some(row => String(row[0] || '').trim().toLowerCase() === name);
  };

  if (hasMatch(APP_CONFIG.SHEETS.PRODUCTION, PRODUCTION_COL.ASSIGNED_TO)) {
    return `Cannot delete "${contractorName}": referenced by Production lots.`;
  }
  if (hasMatch(APP_CONFIG.SHEETS.DISPATCH, DISPATCH_COL.LOGISTICS_CONTRACTOR)) {
    return `Cannot delete "${contractorName}": referenced by Dispatch logistics entries.`;
  }
  if (hasMatch(APP_CONFIG.SHEETS.CONTRACTOR_PAYMENTS, CONTRACTOR_PAYMENTS_COL.CONTRACTOR_NAME)) {
    return `Cannot delete "${contractorName}": has recorded payments.`;
  }
  return null;
}

/**
 * Ensures a contractor master record exists for the given name, appending
 * a minimal record (name only) if it doesn't. Runs under the caller's
 * existing document lock — does not acquire its own.
 * @private
 */
function _ensureContractorExists(contractorName) {
  const name = String(contractorName || '').trim();
  if (!name) return;

  let sheet;
  try {
    sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTORS);
  } catch (e) {
    initContractorsSheet();
    sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTORS);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const names = sheet.getRange(2, CONTRACTORS_COL.CONTRACTOR_NAME, lastRow - 1, 1).getValues();
    const exists = names.some(row => String(row[0] || '').trim().toLowerCase() === name.toLowerCase());
    if (exists) return;
  }

  sheet.appendRow([name, '', '', '', 'Auto-created from rate card entry']);
  invalidateListCache(MASTER_DATA_CACHE_KEYS.CONTRACTORS);
}

/**
 * Renames a contractor across every sheet that holds a de-normalized
 * snapshot of their name as a plain string (no foreign key):
 *  - Contractor Rates: Contractor Name
 *  - Production: Assigned To
 *  - Dispatch: Logistics Contractor
 *  - Contractor Payments: Contractor Name
 * Without this, a renamed contractor's existing Payable lots and recorded
 * Payments stay keyed under the old name, so getContractorLedgerData /
 * getContractorAccountLedger silently stop matching them to the contractor's
 * (now renamed) profile — their balance due effectively disappears from the
 * UI. Runs under the caller's existing document lock — does not acquire
 * its own.
 * @private
 */
function _renameContractorEverywhere(oldName, newName) {
  const tOldRaw = String(oldName || '').trim();
  const tOld = tOldRaw.toLowerCase();
  const tNew = String(newName || '').trim();
  if (!tOld || !tNew || tOldRaw === tNew) return;

  const renameInColumn = (sheetName, col) => {
    let sheet;
    try {
      sheet = getSheet(sheetName);
    } catch (e) {
      return; // Sheet doesn't exist yet, nothing to rename
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const range = sheet.getRange(2, col, lastRow - 1, 1);
    const values = range.getValues();
    let changed = false;

    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0] || '').trim().toLowerCase() === tOld) {
        values[i][0] = tNew;
        changed = true;
      }
    }

    if (changed) range.setValues(values);
  };

  renameInColumn(APP_CONFIG.SHEETS.CONTRACTOR_RATES, CONTRACTOR_RATES_COL.CONTRACTOR_NAME);
  renameInColumn(APP_CONFIG.SHEETS.PRODUCTION, PRODUCTION_COL.ASSIGNED_TO);
  renameInColumn(APP_CONFIG.SHEETS.DISPATCH, DISPATCH_COL.LOGISTICS_CONTRACTOR);
  renameInColumn(APP_CONFIG.SHEETS.CONTRACTOR_PAYMENTS, CONTRACTOR_PAYMENTS_COL.CONTRACTOR_NAME);
  // Display-only "sourced from" snapshot on a Product's Additional Cost
  // line (module_bom.js) — never used in a lookup or balance calc, but
  // still a de-normalized copy of the name this function otherwise
  // promises to fully propagate.
  renameInColumn(APP_CONFIG.SHEETS.BOM_COSTS, BOM_COSTS_COL.CONTRACTOR_NAME);
}

// ── Contractor Ledger (read-only, computed from Production) ────────────

/**
 * Aggregates Completed Production lots (by Assigned To) and Dispatch /
 * Logistics records (by Logistics Contractor) per contractor name, summing
 * Qty and Payable. Mirrors module_dispatch.js's _computeReadyToDispatchMap
 * "derive on read, no sheet of its own" approach.
 */
function getContractorLedgerData() {
  try {
    const map = {};

    let prodSheet;
    try {
      prodSheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
    } catch (e) {
      prodSheet = null;
    }

    if (prodSheet) {
      if (typeof ensureProductionContractorColumns === 'function') {
        ensureProductionContractorColumns(prodSheet);
      }

      const lastRow = prodSheet.getLastRow();
      if (lastRow >= 2) {
        const numCols = PRODUCTION_COL.CONTRACTOR_PAYABLE;
        const data = prodSheet.getRange(2, 1, lastRow - 1, numCols).getValues();

        const allProcesses = typeof _getAllProcessRecords === 'function' ? _getAllProcessRecords() : [];
        const processNameById = {};
        allProcesses.forEach(p => { processNameById[p.processId.toLowerCase()] = p.processName; });

        data.forEach(row => {
          const status = String(row[PRODUCTION_COL.STATUS - 1] || '').trim().toLowerCase();
          if (status !== 'completed') return;

          const contractorName = String(row[PRODUCTION_COL.ASSIGNED_TO - 1] || '').trim();
          if (!contractorName) return;

          const payable = Number(row[PRODUCTION_COL.CONTRACTOR_PAYABLE - 1]) || 0;
          if (payable <= 0) return; // Skip in-house labor with no rate card match

          const qty = Number(row[PRODUCTION_COL.QTY - 1]) || 0;
          const processId = String(row[PRODUCTION_COL.PROCESS_ID - 1] || '').trim().toLowerCase();
          const processName = processNameById[processId] || 'Uncategorized';

          const key = contractorName.toLowerCase();
          if (!map[key]) {
            map[key] = { contractorName: contractorName, lotCount: 0, totalQty: 0, totalPayable: 0, byProcess: {} };
          }

          map[key].lotCount += 1;
          map[key].totalQty += qty;
          map[key].totalPayable += payable;

          if (!map[key].byProcess[processName]) {
            map[key].byProcess[processName] = { processName, lotCount: 0, qty: 0, payable: 0 };
          }
          map[key].byProcess[processName].lotCount += 1;
          map[key].byProcess[processName].qty += qty;
          map[key].byProcess[processName].payable += payable;
        });
      }
    }

    // Fold in Dispatch / Logistics payable — same per-contractor, per-process
    // shape as a Production process bucket, so it shows up identically in
    // the ledger summary.
    _getDispatchLogisticsPayableRows().forEach(d => {
      const key = d.contractorName.toLowerCase();
      if (!map[key]) {
        map[key] = { contractorName: d.contractorName, lotCount: 0, totalQty: 0, totalPayable: 0, byProcess: {} };
      }

      map[key].lotCount += 1;
      map[key].totalQty += d.qty;
      map[key].totalPayable += d.payable;

      if (!map[key].byProcess[LOGISTICS_PROCESS_NAME]) {
        map[key].byProcess[LOGISTICS_PROCESS_NAME] = { processName: LOGISTICS_PROCESS_NAME, lotCount: 0, qty: 0, payable: 0 };
      }
      map[key].byProcess[LOGISTICS_PROCESS_NAME].lotCount += 1;
      map[key].byProcess[LOGISTICS_PROCESS_NAME].qty += d.qty;
      map[key].byProcess[LOGISTICS_PROCESS_NAME].payable += d.payable;
    });

    const paidMap = _getTotalPaidMap();

    Object.values(map).forEach(entry => {
      entry.byProcess = Object.values(entry.byProcess);
      const paidEntry = paidMap[entry.contractorName.toLowerCase()];
      entry.totalPaid = paidEntry ? paidEntry.amount : 0;
      entry.balanceDue = entry.totalPayable - entry.totalPaid;
    });

    // Contractors with payments recorded but no Payable lots (e.g. an
    // advance paid before any lot is completed) still need to show up.
    Object.keys(paidMap).forEach(key => {
      if (map[key]) return;
      map[key] = {
        contractorName: paidMap[key].contractorName,
        lotCount: 0,
        totalQty: 0,
        totalPayable: 0,
        byProcess: [],
        totalPaid: paidMap[key].amount,
        balanceDue: -paidMap[key].amount
      };
    });

    const records = Object.values(map).sort((a, b) => a.contractorName.localeCompare(b.contractorName));

    return buildResponse(true, records);
  } catch (error) {
    Log.error('[getContractorLedgerData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load contractor ledger: ' + error.message);
  }
}

/**
 * Reads the Contractor Payments sheet and sums amount paid per contractor.
 * @returns {Object} { [contractorNameLower]: { contractorName, amount } }
 * @private
 */
function _getTotalPaidMap() {
  const map = {};

  let sheet;
  try {
    sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTOR_PAYMENTS);
  } catch (e) {
    return map; // Payments sheet doesn't exist yet — nobody has been paid
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  data.forEach(row => {
    const contractorName = String(row[CONTRACTOR_PAYMENTS_COL.CONTRACTOR_NAME - 1] || '').trim();
    if (!contractorName) return;

    const amount = Number(row[CONTRACTOR_PAYMENTS_COL.AMOUNT - 1]) || 0;
    const key = contractorName.toLowerCase();
    if (!map[key]) map[key] = { contractorName: contractorName, amount: 0 };
    map[key].amount += amount;
  });

  return map;
}

// ── Contractor Payments (datewise cash record) ──────────────────────────

function initContractorPaymentsSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.CONTRACTOR_PAYMENTS);
    if (!sheet) {
      sheet = ss.insertSheet(APP_CONFIG.SHEETS.CONTRACTOR_PAYMENTS);
    }
    sheet.getRange(1, 1, 1, 5)
      .setValues([['Date', 'Contractor Name', 'Amount', 'Mode / Reference', 'Remarks']])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    SpreadsheetApp.flush();
    return buildResponse(true, null, 'Contractor Payments sheet initialized successfully.');
  } catch (error) {
    Log.error('[initContractorPaymentsSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize Contractor Payments sheet: ' + error.message);
  }
}

/**
 * Retrieves payment records, optionally filtered to one contractor,
 * sorted by date ascending then row order.
 * @param {string} [contractorName]
 */
function getContractorPaymentsData(contractorName) {
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTOR_PAYMENTS);
    } catch (e) {
      initContractorPaymentsSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTOR_PAYMENTS);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, []);

    const filterName = contractorName ? String(contractorName).trim().toLowerCase() : null;
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

    const payments = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rawDate = row[CONTRACTOR_PAYMENTS_COL.DATE - 1];
      const name = String(row[CONTRACTOR_PAYMENTS_COL.CONTRACTOR_NAME - 1] || '').trim();
      if (!name || (filterName && name.toLowerCase() !== filterName)) continue;

      payments.push({
        rowIdx: i + 2,
        date: rawDate instanceof Date ? toSafeDateString(rawDate) : String(rawDate || ''),
        dateRaw: rawDate instanceof Date ? rawDate.toISOString() : null,
        contractorName: name,
        amount: Number(row[CONTRACTOR_PAYMENTS_COL.AMOUNT - 1]) || 0,
        modeReference: String(row[CONTRACTOR_PAYMENTS_COL.MODE_REFERENCE - 1] || '').trim(),
        remarks: String(row[CONTRACTOR_PAYMENTS_COL.REMARKS - 1] || '').trim()
      });
    }

    payments.sort((a, b) => {
      const dateA = a.dateRaw ? new Date(a.dateRaw) : new Date(0);
      const dateB = b.dateRaw ? new Date(b.dateRaw) : new Date(0);
      return dateA - dateB || a.rowIdx - b.rowIdx;
    });

    return buildResponse(true, payments);
  } catch (error) {
    Log.error('[getContractorPaymentsData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load contractor payments: ' + error.message);
  }
}

/**
 * Records a payment made to a contractor.
 */
function recordContractorPayment(formData) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CONTRACTOR_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTOR_PAYMENTS);
    } catch (e) {
      initContractorPaymentsSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTOR_PAYMENTS);
    }

    const contractorName = sanitizeString(formData.contractorName, 'contractorName');
    if (!contractorName) {
      return buildResponse(false, null, 'Contractor is required.');
    }

    const amount = validateNumber(formData.amount, 0.01, 100000000);
    if (amount <= 0) {
      return buildResponse(false, null, 'Payment amount must be greater than zero.');
    }

    const dateObj = toSafeDateObject(formData.date) || new Date();

    const modeReference = sanitizeString(formData.modeReference || '', 'modeReference');
    const remarks = sanitizeString(formData.remarks || '', 'remarks').slice(0, APP_CONFIG.VALIDATION.MAX_REMARKS_LENGTH);

    _ensureContractorExists(contractorName);

    sheet.appendRow([dateObj, contractorName, amount, modeReference, remarks]);

    SpreadsheetApp.flush();

    const msg = `Payment of ${amount} recorded for ${contractorName} on ${toSafeDateString(dateObj)}.`;
    logAction('CREATE', APP_CONFIG.SHEETS.CONTRACTOR_PAYMENTS, contractorName, msg, 'SUCCESS');

    return buildResponse(true, null, 'Payment recorded successfully.');
  } catch (error) {
    Log.error('[recordContractorPayment] Error:', error.message);
    logAction('ERROR', 'recordContractorPayment', formData.contractorName || 'unknown', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to record payment: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes a payment record. Safety-checked against the expected contractor
 * + amount so a stale client reference can't delete the wrong row, same
 * pattern as deleteProduction/deleteDispatch.
 */
function deleteContractorPayment(rowIdx, expectedContractorName, expectedAmount) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(CONTRACTOR_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.CONTRACTOR_PAYMENTS);
    if (!sheet) throw new Error('Contractor Payments sheet not found.');

    const targetRow = parseInt(rowIdx, 10);
    if (isNaN(targetRow) || targetRow < 2 || targetRow > sheet.getLastRow()) {
      return buildResponse(false, null, 'Invalid payment record selected for deletion.');
    }

    const contractorName = String(sheet.getRange(targetRow, CONTRACTOR_PAYMENTS_COL.CONTRACTOR_NAME).getValue()).trim();
    const amount = Number(sheet.getRange(targetRow, CONTRACTOR_PAYMENTS_COL.AMOUNT).getValue()) || 0;

    if (expectedContractorName !== undefined && expectedAmount !== undefined) {
      if (contractorName.toLowerCase() !== String(expectedContractorName || '').trim().toLowerCase() ||
          Math.abs(amount - Number(expectedAmount)) > 0.0001) {
        return buildResponse(false, null, 'Data mismatch: The record has been modified or shifted. Please refresh.');
      }
    }

    sheet.deleteRow(targetRow);
    SpreadsheetApp.flush();

    const msg = `Payment of ${amount} for "${contractorName}" deleted.`;
    logAction('DELETE', APP_CONFIG.SHEETS.CONTRACTOR_PAYMENTS, contractorName, msg, 'SUCCESS');

    return buildResponse(true, null, 'Payment record deleted successfully.');
  } catch (error) {
    Log.error('[deleteContractorPayment] Error:', error.message);
    logAction('ERROR', 'deleteContractorPayment', String(rowIdx), error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete payment: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Builds the full datewise account ledger for one contractor: Payable
 * entries (from Completed Production lots and Dispatch / Logistics records)
 * interleaved with Payment entries (from Contractor Payments), sorted
 * oldest-first to compute a running balance, then returned newest-first
 * (matching every other ledger/table in this app) with each row's balance
 * already attached.
 * @param {string} contractorName
 */
function getContractorAccountLedger(contractorName) {
  try {
    const name = String(contractorName || '').trim();
    if (!name) return buildResponse(false, null, 'Contractor is required.');

    const entries = [];

    // Payable entries — one per Completed Production lot assigned to this contractor
    let prodSheet;
    try {
      prodSheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
    } catch (e) {
      prodSheet = null;
    }

    if (prodSheet) {
      if (typeof ensureProductionContractorColumns === 'function') {
        ensureProductionContractorColumns(prodSheet);
      }
      const lastRow = prodSheet.getLastRow();
      if (lastRow >= 2) {
        const numCols = PRODUCTION_COL.CONTRACTOR_PAYABLE;
        const data = prodSheet.getRange(2, 1, lastRow - 1, numCols).getValues();

        const allProcesses = typeof _getAllProcessRecords === 'function' ? _getAllProcessRecords() : [];
        const processNameById = {};
        allProcesses.forEach(p => { processNameById[p.processId.toLowerCase()] = p.processName; });

        data.forEach(row => {
          const status = String(row[PRODUCTION_COL.STATUS - 1] || '').trim().toLowerCase();
          if (status !== 'completed') return;

          const assignedTo = String(row[PRODUCTION_COL.ASSIGNED_TO - 1] || '').trim();
          if (assignedTo.toLowerCase() !== name.toLowerCase()) return;

          const payable = Number(row[PRODUCTION_COL.CONTRACTOR_PAYABLE - 1]) || 0;
          if (payable <= 0) return;

          const rawDate = row[PRODUCTION_COL.DATE - 1];
          const processId = String(row[PRODUCTION_COL.PROCESS_ID - 1] || '').trim().toLowerCase();
          const processName = processNameById[processId] || 'Uncategorized';
          const productName = String(row[PRODUCTION_COL.PRODUCT_NAME - 1] || '').trim();
          const lotNumber = String(row[PRODUCTION_COL.LOT_NUMBER - 1] || '').trim();

          entries.push({
            date: rawDate instanceof Date ? toSafeDateString(rawDate) : String(rawDate || ''),
            dateRaw: rawDate instanceof Date ? rawDate.toISOString() : null,
            type: 'Payable',
            ref: lotNumber || '-',
            description: `${processName} — ${productName}`,
            amount: payable
          });
        });
      }
    }

    // Payable entries — one per Dispatch row this contractor handled logistics for
    _getDispatchLogisticsPayableRows(name).forEach(d => {
      entries.push({
        date: d.date,
        dateRaw: d.dateRaw,
        type: 'Payable',
        ref: d.dispatchNumber || '-',
        description: `${LOGISTICS_PROCESS_NAME} — ${d.productName}`,
        amount: d.payable
      });
    });

    // Payment entries
    const paymentsRes = getContractorPaymentsData(name);
    if (paymentsRes.success) {
      paymentsRes.data.forEach(p => {
        entries.push({
          date: p.date,
          dateRaw: p.dateRaw,
          type: 'Payment',
          ref: p.modeReference || '-',
          description: p.remarks || '-',
          amount: -p.amount,
          rowIdx: p.rowIdx,
          rawAmount: p.amount
        });
      });
    }

    // Sort oldest-first to accumulate the running balance correctly
    entries.sort((a, b) => {
      const dateA = a.dateRaw ? new Date(a.dateRaw) : new Date(0);
      const dateB = b.dateRaw ? new Date(b.dateRaw) : new Date(0);
      return dateA - dateB;
    });

    let runningBalance = 0;
    let totalPayable = 0;
    let totalPaid = 0;
    entries.forEach(entry => {
      runningBalance += entry.amount;
      entry.balance = runningBalance;
      if (entry.type === 'Payable') totalPayable += entry.amount;
      else totalPaid += -entry.amount;
    });

    // Newest-first for display, matching every other ledger/table in this app
    entries.reverse();

    return buildResponse(true, {
      entries: entries,
      totalPayable: totalPayable,
      totalPaid: totalPaid,
      balanceDue: totalPayable - totalPaid
    });
  } catch (error) {
    Log.error('[getContractorAccountLedger] Error:', error.message);
    return buildResponse(false, null, 'Failed to load contractor account ledger: ' + error.message);
  }
}
