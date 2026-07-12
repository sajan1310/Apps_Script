/**
 * ============================================================================
 * VENDORS MODULE
 * ============================================================================
 * Architecture: Strict Master Data mapping (Profile only).
 * Ledger and pending order math is deferred to the client-side using
 * existing PO and Bill data sets for optimal UI performance.
 */

/**
 * Run this once manually from the GAS Editor to setup the sheet headers
 */
function initVendorsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.VENDORS);
  if (!sheet) {
    sheet = ss.insertSheet(APP_CONFIG.SHEETS.VENDORS);
  }
  sheet.getRange(1, 1, 1, 5).setValues([["Vendor Name", "Contact Number", "Address", "GSTIN", "Remarks"]]).setFontWeight("bold");
  SpreadsheetApp.flush();
}

/**
 * Retrieves all vendors from the database.
 * @returns {Object} JSON Response with vendors array.
 */
function getVendorsData() {
  return getCachedListResponse(MASTER_DATA_CACHE_KEYS.VENDORS, () => {
    try {
      const sheet = getSheet(APP_CONFIG.SHEETS.VENDORS);
      if (!sheet) return buildResponse(false, null, "Vendors sheet not found. Please run initVendorsSheet().");

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return buildResponse(true, []);

      const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      const vendors = data.map(r => ({
        name: String(r[0] || "").trim(),
        contact: String(r[1] || "").trim(),
        address: String(r[2] || "").trim(),
        gstin: String(r[3] || "").trim(),
        remarks: String(r[4] || "").trim()
      })).filter(v => v.name !== "");

      // Sort alphabetically by vendor name
      vendors.sort((a, b) => a.name.localeCompare(b.name));

      return buildResponse(true, vendors);
    } catch (error) {
      return buildResponse(false, null, "Failed to load vendors: " + error.message);
    }
  });
}

/**
 * Renames a vendor across every sheet that holds a de-normalized snapshot of
 * their name — PO Tracker, Bill Ledger, Return Ledger, Wastage Log, and a
 * Product BOM's "Selected Vendor Name" all store it as a plain string (no
 * foreign key); Items Master stores it inside the per-item vendor/rate JSON
 * (ITEMS_COL.VENDOR_DATA). Without this, a renamed vendor's PO/Bill/Return
 * history stays keyed under the old name, so App.Vendor.calculateLedgerAndPending
 * (which filters by exact name match) silently stops finding it — the
 * vendor's ledger/pending-orders effectively disappear from the UI even
 * though the rows still exist — and autoExtractFromPoOrBill's "does this
 * vendor already have a rate for this item" lookup misses too, forking a
 * duplicate vendor/rate entry under the new name instead of updating the
 * existing one. Mirrors module_contractors.js's _renameContractorEverywhere.
 * Runs under the caller's existing document lock — does not acquire its own.
 * @private
 */
function _renameVendorEverywhere(oldName, newName) {
  const tOldRaw = String(oldName || '').trim();
  const tOld = tOldRaw.toLowerCase();
  const tNew = String(newName || '').trim();
  if (!tOld || !tNew || tOldRaw === tNew) return;

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
      if (String(values[i][0] || '').trim().toLowerCase() === tOld) {
        values[i][0] = tNew;
        changed = true;
      }
    }

    if (changed) range.setValues(values);
  };

  // PO Tracker has 2 header rows (data starts row 3); every other sheet here
  // has 1 (data starts row 2) — see the *_SETTINGS blocks in config.js.
  renameInColumn(APP_CONFIG.SHEETS.PO, PO_COL.VENDOR, APP_CONFIG.PO_SETTINGS.DATA_START_ROW);
  renameInColumn(APP_CONFIG.SHEETS.BILL, BILL_COL.VENDOR, 2);
  renameInColumn(APP_CONFIG.SHEETS.RETURN, RETURN_COL.VENDOR, 2);
  renameInColumn(APP_CONFIG.SHEETS.WASTAGE, WASTAGE_COL.VENDOR, 2);
  renameInColumn(APP_CONFIG.SHEETS.BOM, BOM_COL.VENDOR, 2);

  // Items Master stores vendor/rate as a JSON array per item row
  // ([{vendor, rate}, ...]), not a single plain-string column — rewrite the
  // vendor name inside each entry instead of the whole cell.
  let itemsSheet;
  try {
    itemsSheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
  } catch (e) {
    itemsSheet = null;
  }
  if (itemsSheet) {
    const lastRow = itemsSheet.getLastRow();
    if (lastRow >= 2) {
      const range = itemsSheet.getRange(2, ITEMS_COL.VENDOR_DATA, lastRow - 1, 1);
      const values = range.getValues();
      let changed = false;
      for (let i = 0; i < values.length; i++) {
        const raw = String(values[i][0] || '').trim();
        if (!raw) continue;
        let parsed;
        try {
          parsed = JSON.parse(raw);
        } catch (e) {
          continue;
        }
        if (!Array.isArray(parsed)) continue;
        let rowChanged = false;
        parsed.forEach(v => {
          if (String((v && v.vendor) || '').trim().toLowerCase() === tOld) {
            v.vendor = tNew;
            rowChanged = true;
          }
        });
        if (rowChanged) {
          values[i][0] = JSON.stringify(parsed);
          changed = true;
        }
      }
      if (changed) range.setValues(values);
    }
  }

  // Items Master's vendor/rate JSON was potentially rewritten above.
  invalidateListCache(MASTER_DATA_CACHE_KEYS.ITEMS);
}

/**
 * Creates or updates a vendor record.
 * @param {Object} formData The vendor data object
 * @returns {Object} Standardised API response
 */
function saveVendor(formData) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) return buildResponse(false, null, "System busy. Please try again.");

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.VENDORS);
    if (!sheet) throw new Error("Vendors sheet not found.");

    const newName = sanitizeString(formData.vendorName, "vendorName");
    if (!newName) throw new Error("Vendor name cannot be empty.");

    const isEdit = !!formData.originalVendorName;
    const originalName = isEdit ? sanitizeString(formData.originalVendorName, "originalName") : newName;

    const data = sheet.getDataRange().getValues();
    let targetRow = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === originalName.toLowerCase()) {
        targetRow = i + 1;
        break;
      }
    }

    if (isEdit && targetRow === -1) throw new Error("Original vendor not found in sheet.");
    if (!isEdit && targetRow !== -1) throw new Error("A vendor with this name already exists.");
    
    // Check for duplicates if renaming
    if (isEdit && newName.toLowerCase() !== originalName.toLowerCase()) {
       const dupeCheck = data.findIndex((r, idx) => idx > 0 && String(r[0]).trim().toLowerCase() === newName.toLowerCase());
       if (dupeCheck !== -1) throw new Error("Another vendor with this new name already exists.");
    }

    const rowData = [
      newName,
      sanitizeString(formData.contact || "", "contact"),
      sanitizeString(formData.address || "", "address"),
      sanitizeString(formData.gstin || "", "gstin"),
      sanitizeString(formData.remarks || "", "remarks")
    ];

    if (isEdit) {
      sheet.getRange(targetRow, 1, 1, 5).setValues([rowData]);
      // Propagate the rename to every sheet holding a de-normalized copy of
      // the vendor name before it's out of scope (see _renameVendorEverywhere).
      // Uses a plain string compare (not case-insensitive) so a casing-only
      // edit ("ABC Traders" -> "abc traders") still cascades — otherwise the
      // Vendors master's own row above is the only place that picks up the
      // new casing, leaving every PO/Bill/Return/Wastage/BOM reference in
      // the old casing indefinitely.
      if (newName !== originalName) {
        _renameVendorEverywhere(originalName, newName);
      }
    } else {
      sheet.appendRow(rowData);
    }

    SpreadsheetApp.flush();
    invalidateListCache(MASTER_DATA_CACHE_KEYS.VENDORS);

    const action = isEdit ? 'UPDATE_VENDOR' : 'CREATE_VENDOR';
    const msg = isEdit ? "Vendor profile updated successfully." : "New Vendor registered.";
    logAction(action, APP_CONFIG.SHEETS.VENDORS, newName, JSON.stringify(formData), 'SUCCESS');

    return buildResponse(true, { name: newName }, msg);

  } catch (error) {
    logAction('SAVE_VENDOR_ERROR', APP_CONFIG.SHEETS.VENDORS, formData ? formData.vendorName : 'unknown', error.message, 'ERROR');
    return buildResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Scans the entire PO Tracker history and auto-populates the Vendors
 * master and Items Master sheets from it.
 *
 * For every PO line item:
 * - Vendor name + contact location are extracted. New vendors are
 *   appended to the Vendors sheet; existing vendors only have their
 *   Contact Number filled in if it is currently blank.
 * - Item name + size + narration are extracted. New (name, size)
 *   combinations are appended to Items Master with the vendor/rate pair
 *   from the PO line. For items that already exist, any vendor not yet
 *   listed against that item is appended with the rate from PO history
 *   (existing vendor/rate pairs are left untouched).
 *
 * Later rows in the PO sheet win when the same vendor/item appears more
 * than once, so the imported rate reflects the most recent PO price.
 *
 * @returns {Object} Standardised API response with a summary of changes
 */
function syncVendorsFromPOHistory() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return buildResponse(false, null, "System busy. Please try again.");

  try {
    const poSheet = getSheet(APP_CONFIG.SHEETS.PO);
    const startRow = APP_CONFIG.PO_SETTINGS.DATA_START_ROW;
    const lastRow = poSheet.getLastRow();

    const summary = { vendorsAdded: 0, vendorsUpdated: 0, itemsAdded: 0, itemsUpdated: 0 };

    if (lastRow < startRow) {
      return buildResponse(true, summary, "No PO history found to sync.");
    }

    const base = PO_COL.PO_NUMBER;
    const numCols = PO_COL.BASE_RATE - base + 1;
    const rows = poSheet.getRange(startRow, base, lastRow - startRow + 1, numCols).getValues();

    const OFF = {
      vendor: PO_COL.VENDOR - base,
      contact: PO_COL.CONTACT - base,
      itemName: PO_COL.ITEM_NAME - base,
      narration: PO_COL.NARRATION - base,
      size: PO_COL.SIZE - base,
      price: PO_COL.PRICE - base,
      baseRate: PO_COL.BASE_RATE - base
    };

    // Aggregate vendors and item/vendor rates from PO history.
    // Later rows overwrite earlier ones, so the latest negotiated rate wins.
    const vendorMap = {};
    const itemMap = {};

    rows.forEach(r => {
      const vendor = String(r[OFF.vendor] || '').trim();
      const contact = String(r[OFF.contact] || '').trim();
      const itemName = String(r[OFF.itemName] || '').trim();
      const narration = String(r[OFF.narration] || '').trim();
      const size = String(r[OFF.size] || '').trim();
      const rawPrice = Number(r[OFF.price]) || 0;
      const rawBaseRate = Number(r[OFF.baseRate]) || 0;
      // Legacy rows predating BASE_RATE have it blank — fall back to the
      // as-entered price (today's implicit Pcs-only assumption).
      const price = rawBaseRate > 0 ? rawBaseRate : rawPrice;

      if (vendor) {
        const vKey = vendor.toLowerCase();
        if (!vendorMap[vKey]) {
          vendorMap[vKey] = { name: vendor, contact: contact };
        } else if (!vendorMap[vKey].contact && contact) {
          vendorMap[vKey].contact = contact;
        }
      }

      if (itemName) {
        const iKey = `${itemName.toLowerCase()}|${size.toLowerCase()}`;
        if (!itemMap[iKey]) {
          itemMap[iKey] = { name: itemName, size: size, narration: narration, vendors: {} };
        } else if (!itemMap[iKey].narration && narration) {
          itemMap[iKey].narration = narration;
        }

        if (vendor && price >= MIN_VENDOR_RATE) {
          itemMap[iKey].vendors[vendor.toLowerCase()] = { vendor: vendor, rate: price };
        }
      }
    });

    // ─── Sync Vendors sheet ───
    let vSheet;
    try {
      vSheet = getSheet(APP_CONFIG.SHEETS.VENDORS);
    } catch (e) {
      vSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(APP_CONFIG.SHEETS.VENDORS);
    }

    if (vSheet.getLastRow() === 0) {
      vSheet.getRange(1, 1, 1, 5)
        .setValues([["Vendor Name", "Contact Number", "Address", "GSTIN", "Remarks"]])
        .setFontWeight("bold");
    }

    const vLastRow = vSheet.getLastRow();
    const existingVendors = {};
    const contactOffset = VENDORS_COL.CONTACT - VENDORS_COL.VENDOR_NAME;
    // Kept around (not just destructured into existingVendors) so a blank-
    // contact fill can be applied in-memory and flushed back in one
    // setValues() call below, instead of one setValue() per vendor.
    let vendorNameContactValues = [];

    if (vLastRow >= 2) {
      vendorNameContactValues = vSheet.getRange(2, VENDORS_COL.VENDOR_NAME, vLastRow - 1, 2).getValues();
      vendorNameContactValues.forEach((row, idx) => {
        const name = String(row[0] || '').trim();
        if (name) {
          existingVendors[name.toLowerCase()] = {
            row: idx + 2,
            contact: String(row[contactOffset] || '').trim()
          };
        }
      });
    }

    const newVendorRows = [];
    let vendorContactsChanged = false;
    Object.values(vendorMap).forEach(v => {
      const existing = existingVendors[v.name.toLowerCase()];
      if (!existing) {
        newVendorRows.push([
          sanitizeString(v.name, 'vendorName'),
          sanitizeString(v.contact, 'contact'),
          '',
          '',
          'Auto-imported from PO history'
        ]);
        summary.vendorsAdded++;
      } else if (!existing.contact && v.contact) {
        vendorNameContactValues[existing.row - 2][contactOffset] = sanitizeString(v.contact, 'contact');
        vendorContactsChanged = true;
        summary.vendorsUpdated++;
      }
    });

    if (vendorContactsChanged) {
      vSheet.getRange(2, VENDORS_COL.VENDOR_NAME, vendorNameContactValues.length, 2).setValues(vendorNameContactValues);
    }

    if (newVendorRows.length) {
      vSheet.getRange(vSheet.getLastRow() + 1, 1, newVendorRows.length, 5).setValues(newVendorRows);
    }

    // ─── Sync Items Master sheet ───
    const iSheet = getSheet(APP_CONFIG.SHEETS.ITEMS);

    if (iSheet.getLastRow() === 0) {
      _writeItemHeaders(iSheet, DEFAULT_VENDOR_SLOTS);
    }

    Object.values(itemMap).forEach(item => {
      const vendorList = Object.values(item.vendors);
      const { matchRow } = _findItemRow(iSheet, item.name, item.size);

      if (matchRow === -1) {
        const targetRow = iSheet.getLastRow() + 1;
        iSheet.getRange(targetRow, ITEMS_COL.ITEM_NAME, 1, 5)
          .setValues([[
            sanitizeString(item.name, 'itemName'),
            sanitizeString(item.size, 'itemSize'),
            '',
            sanitizeString(item.narration, 'narration'),
            ''
          ]]);

        if (vendorList.length) {
          const vendorRow = [];
          vendorList.forEach(v => vendorRow.push(sanitizeString(v.vendor, 'vendor.name'), v.rate));
          iSheet.getRange(targetRow, ITEMS_COL.VENDOR_DATA, 1, vendorRow.length).setValues([vendorRow]);
        }

        syncStockForItem('ensure', { name: item.name, size: item.size });

        summary.itemsAdded++;
        return;
      }

      // Existing item: fill blank narration, append any new vendor rates.
      // Narration and the vendor-pair grid are read together in one call
      // (instead of a single-cell getValue() for narration plus a separate
      // getValues() for the vendor range) since they're adjacent columns.
      let changed = false;

      const lastCol = iSheet.getLastColumn();
      const readWidth = Math.max(lastCol, ITEMS_COL.VENDOR_DATA) - ITEMS_COL.NARRATION + 1;
      const narrationToVendorRow = iSheet.getRange(matchRow, ITEMS_COL.NARRATION, 1, readWidth).getValues()[0];
      const vendorOffset = ITEMS_COL.VENDOR_DATA - ITEMS_COL.NARRATION;

      let narrationUpdate = null;
      if (item.narration && !String(narrationToVendorRow[0] || '').trim()) {
        narrationUpdate = sanitizeString(item.narration, 'narration');
      }

      const existingPairs = [];
      for (let v = vendorOffset; v + 1 < narrationToVendorRow.length; v += 2) {
        const vName = String(narrationToVendorRow[v] || '').trim();
        const vRateRaw = narrationToVendorRow[v + 1];
        if (!vName && (vRateRaw === '' || vRateRaw === null)) break;
        if (vName) existingPairs.push({ vendor: vName, rate: Number(vRateRaw) || 0 });
      }

      const existingKeys = new Set(existingPairs.map(v => v.vendor.toLowerCase()));
      let pairsChanged = false;

      vendorList.forEach(v => {
        const vKey = v.vendor.toLowerCase();
        if (!existingKeys.has(vKey)) {
          existingPairs.push({ vendor: v.vendor, rate: v.rate });
          existingKeys.add(vKey);
          pairsChanged = true;
        }
      });

      if (narrationUpdate !== null) {
        iSheet.getRange(matchRow, ITEMS_COL.NARRATION).setValue(narrationUpdate);
        changed = true;
      }

      if (pairsChanged) {
        const vendorRow = [];
        existingPairs.forEach(v => vendorRow.push(sanitizeString(v.vendor, 'vendor.name'), v.rate));
        iSheet.getRange(matchRow, ITEMS_COL.VENDOR_DATA, 1, vendorRow.length).setValues([vendorRow]);
        changed = true;
      }

      if (changed) summary.itemsUpdated++;
    });

    SpreadsheetApp.flush();

    const msg = `Sync complete: ${summary.vendorsAdded} vendor(s) added, ` +
      `${summary.vendorsUpdated} vendor(s) updated, ${summary.itemsAdded} item(s) added, ` +
      `${summary.itemsUpdated} item(s) updated from PO history.`;

    logAction('SYNC', APP_CONFIG.SHEETS.VENDORS, 'PO_HISTORY', msg, 'SUCCESS');
    invalidateListCache(MASTER_DATA_CACHE_KEYS.VENDORS, MASTER_DATA_CACHE_KEYS.ITEMS);

    return buildResponse(true, summary, msg);
  } catch (error) {
    logAction('SYNC_ERROR', APP_CONFIG.SHEETS.VENDORS, 'PO_HISTORY', error.message, 'ERROR');
    return buildResponse(false, null, "Sync failed: " + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes a vendor record from the master sheet.
 * @param {string} vendorName The name of the vendor to delete
 * @returns {Object} Standardised API response
 */
function deleteVendor(vendorName) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) return buildResponse(false, null, "System busy.");

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.VENDORS);
    if (!sheet) throw new Error("Vendors sheet not found.");

    const targetName = String(vendorName || "").trim().toLowerCase();
    const data = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 1).getValues();
    let targetRow = -1;

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === targetName) {
        targetRow = i + 2;
        break;
      }
    }

    if (targetRow === -1) throw new Error("Vendor not found.");

    sheet.deleteRow(targetRow);
    SpreadsheetApp.flush();
    invalidateListCache(MASTER_DATA_CACHE_KEYS.VENDORS);

    logAction('DELETE_VENDOR', APP_CONFIG.SHEETS.VENDORS, vendorName, `Vendor deleted`, 'SUCCESS');
    
    return buildResponse(true, null, `Vendor "${vendorName}" deleted from master.`);

  } catch (error) {
    logAction('DELETE_VENDOR_ERROR', APP_CONFIG.SHEETS.VENDORS, vendorName, error.message, 'ERROR');
    return buildResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes multiple vendor records from the master sheet in a single batch.
 * @param {Array<string>} vendorNames - Names of the vendors to delete
 * @returns {Object} Standardised API response
 */
function deleteVendorsBulk(vendorNames) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) return buildResponse(false, null, "System busy.");

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.VENDORS);
    if (!sheet) throw new Error("Vendors sheet not found.");

    const targetSet = new Set(
      (vendorNames || []).map(n => String(n || '').trim().toLowerCase()).filter(Boolean)
    );
    if (targetSet.size === 0) {
      return buildResponse(true, null, 'No vendors selected.');
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return buildResponse(true, null, 'No vendors to delete.');
    }

    const { rowsDeleted } = rewriteSheetExcludingRows(sheet, 2, row => {
      const name = String(row[VENDORS_COL.VENDOR_NAME - 1] || '').trim().toLowerCase();
      return targetSet.has(name);
    });

    SpreadsheetApp.flush();
    invalidateListCache(MASTER_DATA_CACHE_KEYS.VENDORS);

    const msg = `Deleted ${rowsDeleted} vendor(s) from master.`;
    logAction('BULK_DELETE', APP_CONFIG.SHEETS.VENDORS, 'multiple', msg, 'SUCCESS');

    return buildResponse(true, null, msg);
  } catch (error) {
    logAction('ERROR', 'deleteVendorsBulk', 'multiple', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete vendors: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Automatically extracts and updates Vendor and Item Master details from a saved PO or Bill.
 * @param {string} vendorName - Name of the vendor
 * @param {string} contact - Contact location or number
 * @param {Array<Object>} items - List of items from PO or Bill
 *    Each item has: name, size, narration (optional), price (rate as entered),
 *    baseRate (optional — rate converted to the item's Base Unit, e.g. Pcs),
 *    poNumber (optional — per-line PO reference, used by Bill saves).
 *    Items Master vendor rates are always per-Base-Unit, so baseRate (when
 *    present) is used instead of the as-entered price — otherwise a line
 *    quoted per Gross/Kg/etc. would poison the item's per-Pc rate.
 * @param {Object} [docInfo] - Source document info, logged to Rate History
 *    @param {Date} [docInfo.date] - PO/Bill date (defaults to now)
 *    @param {string} [docInfo.poNumber] - PO number (fallback when item.poNumber absent)
 *    @param {string} [docInfo.billNumber] - Bill number, if the source is a Bill
 */
function autoExtractFromPoOrBill(vendorName, contact, items, docInfo) {
  try {
    const vNameClean = String(vendorName || '').trim();
    if (!vNameClean) return;

    const contactClean = String(contact || '').trim();

    // 1. Check/Update Vendor (single read, already efficient)
    const vSheet = getSheet(APP_CONFIG.SHEETS.VENDORS);
    if (vSheet) {
      const vLastRow = vSheet.getLastRow();
      const existingVendors = {};
      const contactOffset = VENDORS_COL.CONTACT - VENDORS_COL.VENDOR_NAME;

      if (vLastRow >= 2) {
        vSheet.getRange(2, VENDORS_COL.VENDOR_NAME, vLastRow - 1, 2)
          .getValues()
          .forEach((row, idx) => {
            const name = String(row[0] || '').trim();
            if (name) {
              existingVendors[name.toLowerCase()] = {
                row: idx + 2,
                contact: String(row[contactOffset] || '').trim()
              };
            }
          });
      }

      const existing = existingVendors[vNameClean.toLowerCase()];
      if (!existing) {
        vSheet.appendRow([
          sanitizeString(vNameClean, 'vendorName'),
          sanitizeString(contactClean, 'contact'),
          '',
          '',
          'Auto-extracted on save'
        ]);
      } else if (!existing.contact && contactClean) {
        vSheet.getRange(existing.row, VENDORS_COL.CONTACT).setValue(sanitizeString(contactClean, 'contact'));
      }
    }

    // 2. Check/Update Items and Rates
    // Read the entire Items sheet ONCE into memory, then process all items against
    // that in-memory map. Writes are deferred and executed after the loop.
    // This replaces the previous N-reads pattern (one full sheet read per item).
    const iSheet = getSheet(APP_CONFIG.SHEETS.ITEMS);
    if (!iSheet) return;

    if (iSheet.getLastRow() === 0) {
      _writeItemHeaders(iSheet, DEFAULT_VENDOR_SLOTS);
    }

    const iLastRow = iSheet.getLastRow();
    const iLastCol = Math.max(iSheet.getLastColumn(), ITEMS_COL.VENDOR_DATA + 1);

    // Build in-memory item map from single batch read
    // key: "name|size" (all lowercase) → { sheetRow, currentNarration, currentPurchaseUnit, vendorPairs }
    const itemMemMap = {};

    if (iLastRow >= 2) {
      const allItemData = iSheet.getRange(2, 1, iLastRow - 1, iLastCol).getValues();

      allItemData.forEach((row, idx) => {
        const rowName = String(row[ITEMS_COL.ITEM_NAME - 1] || '').trim();
        if (!rowName) return;

        const rowSize     = String(row[ITEMS_COL.SIZE - 1]      || '').trim();
        const rowNarration = String(row[ITEMS_COL.NARRATION - 1] || '').trim();
        const rowPurchaseUnit = String(row[ITEMS_COL.PURCHASE_UNIT - 1] || '').trim();

        const vendorPairs = [];
        for (let v = ITEMS_COL.VENDOR_DATA - 1; v + 1 < row.length; v += 2) {
          const vn = String(row[v] || '').trim();
          const vr = row[v + 1];
          if (!vn && (vr === '' || vr === null || vr === undefined)) break;
          if (vn) vendorPairs.push({ vendor: vn, rate: Number(vr) || 0 });
        }

        const key = `${rowName.toLowerCase()}|${rowSize.toLowerCase()}`;
        if (!itemMemMap[key]) {
          itemMemMap[key] = {
            sheetRow: idx + 2,
            currentNarration: rowNarration,
            currentPurchaseUnit: rowPurchaseUnit,
            vendorPairs
          };
        }
      });
    }

    // Collect pending writes — execute after the loop to avoid interleaving reads/writes
    const narrationUpdates    = []; // { sheetRow, value }
    const purchaseUnitUpdates = []; // { sheetRow, value }
    const vendorUpdates       = []; // { sheetRow, vendorPairs }
    const newItemRows         = []; // full rows to batch-append
    const newStockItems       = []; // { name, size } to sync to Stock sheet
    const rateHistoryRows  = []; // rows to append to the Rate History sheet

    const docDate = (docInfo && docInfo.date instanceof Date) ? docInfo.date : new Date();
    const docPoNumber = String((docInfo && docInfo.poNumber) || '').trim();
    const docBillNumber = String((docInfo && docInfo.billNumber) || '').trim();

    let nextNewRow = iLastRow + 1; // tracks append position as new items are queued

    items.forEach(item => {
      const itemName     = String(item.name      || '').trim();
      if (!itemName) return;

      const itemSize     = String(item.size      || '').trim();
      const itemNarration = String(item.narration || '').trim();
      const rate = (item.baseRate !== undefined && item.baseRate !== null)
        ? Number(item.baseRate) || 0
        : Number(item.price) || 0;

      if (rate < MIN_VENDOR_RATE) return;

      rateHistoryRows.push([
        docDate,
        sanitizeString(itemName, 'itemName'),
        sanitizeString(itemSize, 'itemSize'),
        sanitizeString(itemNarration, 'narration'),
        sanitizeString(vNameClean, 'vendor.name'),
        rate,
        sanitizeString(String(item.poNumber || '').trim() || docPoNumber, 'poNumber'),
        docBillNumber
      ]);

      // The unit this line was actually purchased in — saved as the item's
      // new Purchase Unit below, regardless of whether it converts cleanly
      // into the item's Base Unit (that's checked separately, and must
      // never block recording what unit was actually used).
      const itemUnit = String(item.unit || '').trim() || 'Pcs';

      const key = `${itemName.toLowerCase()}|${itemSize.toLowerCase()}`;
      const mem = itemMemMap[key];

      if (!mem) {
        // New item — queue for batch append. Fixed columns must match
        // ITEMS_COL exactly (name, size, remarks, narration, spec, baseUnit,
        // purchaseUnit, weightPerBaseUnit) before vendor pairs start —
        // a never-before-seen item has no prior Base Unit, so the unit it
        // was first purchased in is used for both, making conversion an
        // identity no-op until the user edits it in Items Master.
        newItemRows.push([
          sanitizeString(itemName,     'itemName'),
          sanitizeString(itemSize,     'itemSize'),
          '',
          sanitizeString(itemNarration, 'narration'),
          '',
          sanitizeString(itemUnit,     'baseUnit'),
          sanitizeString(itemUnit,     'purchaseUnit'),
          0,
          sanitizeString(vNameClean,   'vendor.name'),
          rate
        ]);
        newStockItems.push({ name: itemName, size: itemSize });

        // Register in map so duplicate items in the same bill don't append twice
        itemMemMap[key] = {
          sheetRow: nextNewRow,
          currentNarration: itemNarration,
          currentPurchaseUnit: itemUnit,
          vendorPairs: [{ vendor: vNameClean, rate }]
        };
        nextNewRow++;
      } else {
        // Existing item — check narration, purchase unit, and vendor rate

        if (itemNarration && !mem.currentNarration) {
          narrationUpdates.push({ sheetRow: mem.sheetRow, value: itemNarration });
          mem.currentNarration = itemNarration; // update map so second pass doesn't re-queue
        }

        // Always reflect the unit actually used on this line as the item's
        // Purchase Unit, so the next PO/Bill form defaults to it.
        if (itemUnit && itemUnit !== mem.currentPurchaseUnit) {
          purchaseUnitUpdates.push({ sheetRow: mem.sheetRow, value: itemUnit });
          mem.currentPurchaseUnit = itemUnit;
        }

        const pairs = mem.vendorPairs.slice();
        const idx   = pairs.findIndex(p => p.vendor.toLowerCase() === vNameClean.toLowerCase());
        let changed = false;

        if (idx === -1) {
          pairs.push({ vendor: vNameClean, rate });
          changed = true;
        } else if (pairs[idx].rate !== rate) {
          pairs[idx].rate = rate;
          changed = true;
        }

        if (changed) {
          vendorUpdates.push({ sheetRow: mem.sheetRow, vendorPairs: pairs });
          mem.vendorPairs = pairs; // update map so second pass doesn't re-queue
        }
      }
    });

    // ── Execute all writes ──────────────────────────────────────────

    // Batch-append new item rows in one call
    if (newItemRows.length > 0) {
      iSheet.getRange(iLastRow + 1, 1, newItemRows.length, newItemRows[0].length)
        .setValues(newItemRows);
    }

    // Narration fills (rare — only when a blank narration gets populated)
    narrationUpdates.forEach(({ sheetRow, value }) => {
      iSheet.getRange(sheetRow, ITEMS_COL.NARRATION)
        .setValue(sanitizeString(value, 'narration'));
    });

    // Purchase Unit — kept in sync with whatever unit was last actually used
    purchaseUnitUpdates.forEach(({ sheetRow, value }) => {
      iSheet.getRange(sheetRow, ITEMS_COL.PURCHASE_UNIT)
        .setValue(sanitizeString(value, 'purchaseUnit'));
    });

    // Vendor rate updates (clear old range, write new pairs)
    vendorUpdates.forEach(({ sheetRow, vendorPairs }) => {
      const vendorRow = [];
      vendorPairs.forEach(p => vendorRow.push(sanitizeString(p.vendor, 'vendor.name'), p.rate));
      iSheet.getRange(sheetRow, ITEMS_COL.VENDOR_DATA, 1, iLastCol - ITEMS_COL.VENDOR_DATA + 1).clearContent();
      iSheet.getRange(sheetRow, ITEMS_COL.VENDOR_DATA, 1, vendorRow.length).setValues([vendorRow]);
    });

    // Sync new items to Stock sheet
    newStockItems.forEach(({ name, size }) => {
      syncStockForItem('ensure', { name, size });
    });

    // Append rate snapshots to Rate History (best-effort — must not block the save)
    if (rateHistoryRows.length > 0) {
      try {
        _appendRateHistoryRows(rateHistoryRows);
      } catch (rateHistoryError) {
        Log.error('[autoExtractFromPoOrBill] Rate History append failed:', rateHistoryError.message);
        logAction('AUTO_EXTRACT_ERROR', 'RateHistory', vendorName, rateHistoryError.message, 'ERROR');
      }
    }

    invalidateListCache(MASTER_DATA_CACHE_KEYS.VENDORS, MASTER_DATA_CACHE_KEYS.ITEMS);

  } catch (error) {
    Log.error('[autoExtractFromPoOrBill] Error:', error.message);
    logAction('AUTO_EXTRACT_ERROR', 'autoExtractFromPoOrBill', vendorName, error.message, 'ERROR');
  }
}

/**
 * Self-healing migration: ensures the Rate History sheet has the expected
 * headers (including PO Number / Bill Number), writing them if the sheet
 * is blank or missing the trailing reference columns.
 * Result is cached per-execution so repeated saves don't re-check headers.
 *
 * @param {Sheet} sheet - The Google Sheet object for Rate History
 * @private
 */
function _ensureRateHistoryColumns(sheet) {
  const cache = CacheService.getScriptCache();
  if (cache.get('rate_history_columns_ok') === '1') return;

  const expectedHeaders = ['Date', 'Item Name', 'Size', 'Narration', 'Vendor Name', 'Rate', 'PO Number', 'Bill Number'];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    SpreadsheetApp.flush();
    cache.put('rate_history_columns_ok', '1', 3600);
    return;
  }

  const lastCol = sheet.getLastColumn();
  const headers = lastCol >= 1 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const hasPoNumber = headers.some(h => String(h).trim().toLowerCase() === 'po number');
  const hasBillNumber = headers.some(h => String(h).trim().toLowerCase() === 'bill number');

  if (!hasPoNumber) {
    sheet.getRange(1, RATE_HISTORY_COL.PO_NUMBER).setValue('PO Number');
  }
  if (!hasBillNumber) {
    sheet.getRange(1, RATE_HISTORY_COL.BILL_NUMBER).setValue('Bill Number');
  }
  if (!hasPoNumber || !hasBillNumber) {
    SpreadsheetApp.flush();
  }

  cache.put('rate_history_columns_ok', '1', 3600);
}

/**
 * Batch-appends rows to the Rate History sheet, healing its headers first.
 * @param {Array<Array>} rows - Rows matching RATE_HISTORY_COL order
 * @private
 */
function _appendRateHistoryRows(rows) {
  if (!rows || rows.length === 0) return;

  const sheet = getSheet(APP_CONFIG.SHEETS.RATE_HISTORY);
  _ensureRateHistoryColumns(sheet);

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}