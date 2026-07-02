/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_bom.gs — BILL OF MATERIALS (BOM) MODULE
 * 
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Complete CRUD operations for Bill of Materials (BOM) including:
 * - Retrieval and grouping of BOMs by Product ID (getBOMData)
 * - Creation and updates (saveBOM)
 * - Deletion of BOM configurations (deleteBOM)
 * - Auto-generation of unique Product IDs (getNextProductId)
 * 
 * Sheet Layout (BOM):
 * ───────────────────────────────────────────────────────────────────────────
 * Col A (1):   Product ID (primary key for grouping, e.g. PRD-1001)
 * Col B (2):   Product Name
 * Col C (3):   Item (component name)
 * Col D (4):   Size (component size)
 * Col E (5):   Narration (component narration)
 * Col F (6):   Rate (chosen vendor rate)
 * Col G (7):   Vendor (chosen vendor name)
 * Col H (8):   Quantity per Product
 * Col J (10):  Process Group — reference to Process Master.Process ID
 *
 * Sheet Layout (BOM Additional Costs):
 * ───────────────────────────────────────────────────────────────────────────
 * Col A (1):   Product ID (links back to BOM sheet)
 * Col B (2):   Cost Description (e.g. "Labor - Fitting", "Paint Cost")
 * Col C (3):   Rate (per unit of product)
 * Col D (4):   Process (free-string Process Name — same convention as
 *              Contractor Rates; matches a Process Master name or the
 *              "Dispatch / Logistics" virtual category)
 *
 * A Product is a collection of real Processes (each contributing its own
 * components + cost) plus product-level additional costs — component
 * "Process Group" (BOM_COL.PROCESS_GROUP) is no longer a free-text label,
 * it is a real foreign key into Process Master, populated via a dropdown
 * in the Products UI.
 * ═══════════════════════════════════════════════════════════════════════════
 */

const BOM_LOCK_TIMEOUT_MS = 15000;

// ─────────────────────────────────────────────────────────────────────────
// BOM ACCESS CONTROL
//
// The BOM sheet contains commercially sensitive data (vendor rates, costs,
// margins). Access to the full BOM dataset (getBOMData/saveBOM/deleteBOM/
// getNextProductId) requires a short-lived access token obtained via
// verifyBOMAccess(password). Staff still get a cost-free product list via
// getBOMProductionData() so the Production tab keeps working.
// ─────────────────────────────────────────────────────────────────────────
const BOM_AUTH_CACHE_PREFIX = 'BOM_AUTH_TOKEN_';
const BOM_AUTH_TTL_SECONDS = 21600; // 6 hours (CacheService maximum)
const BOM_PASSWORD_PROPERTY = 'BOM_PASSWORD_HASH';

/**
 * Hashes a password with SHA-256 for storage/comparison.
 * @param {string} password
 * @returns {string} Hex-encoded digest
 */
function _hashBOMPassword(password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password || ''));
  return bytes.map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}

/**
 * Sets (or changes) the BOM access password. Not exposed to the client UI —
 * run this once from the Apps Script editor via ONE_TIME_setBOMPassword().
 * @param {string} newPassword
 */
function setBOMPassword(newPassword) {
  if (!newPassword || String(newPassword).length < 4) {
    return buildResponse(false, null, 'Password must be at least 4 characters.');
  }
  PropertiesService.getScriptProperties().setProperty(BOM_PASSWORD_PROPERTY, _hashBOMPassword(newPassword));
  return buildResponse(true, null, 'BOM password updated successfully.');
}

/**
 * ONE-TIME SETUP — run manually from the Apps Script editor:
 * 1. Replace 'CHANGE_ME' below with your desired BOM password.
 * 2. Select this function in the editor's function dropdown and click "Run".
 * 3. Afterwards, change the string back to 'CHANGE_ME' (or anything) so the
 *    real password isn't left sitting in the source code.
 */
function ONE_TIME_setBOMPassword() {
  const result = setBOMPassword('Sajan@1995');
  console.log(result.message);
}

/**
 * Verifies a BOM access password and, if correct, issues a temporary token.
 * @param {string} password
 * @returns {Object} { success, data: { token }, message }
 */
function verifyBOMAccess(password) {
  try {
    const storedHash = PropertiesService.getScriptProperties().getProperty(BOM_PASSWORD_PROPERTY);
    if (!storedHash) {
      return buildResponse(false, null, 'BOM password has not been configured yet. Run ONE_TIME_setBOMPassword in the Apps Script editor.');
    }

    if (_hashBOMPassword(password) !== storedHash) {
      return buildResponse(false, null, 'Incorrect password.');
    }

    const token = Utilities.getUuid();
    CacheService.getScriptCache().put(BOM_AUTH_CACHE_PREFIX + token, '1', BOM_AUTH_TTL_SECONDS);

    return buildResponse(true, { token: token }, 'Access granted.');
  } catch (error) {
    console.error('[verifyBOMAccess] Error:', error.message);
    return buildResponse(false, null, 'Failed to verify password: ' + error.message);
  }
}

/**
 * Throws if the given token is missing/invalid/expired.
 * @param {string} token
 */
function _requireBOMAccess(token) {
  const valid = !!token && CacheService.getScriptCache().get(BOM_AUTH_CACHE_PREFIX + token) === '1';
  if (!valid) {
    throw new Error('UNAUTHORIZED_BOM');
  }
}

/**
 * Initializes the BOM sheet with correct headers.
 * Run once manually or triggered automatically when sheet is missing.
 */
function initBOMSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.BOM);
    if (!sheet) {
      sheet = ss.insertSheet(APP_CONFIG.SHEETS.BOM);
    }

    const headers = [
      'Product ID',
      'Product Name',
      'Item',
      'Size',
      'Narration',
      'Rate',
      'Vendor',
      'Quantity per Product',
      'Remarks',
      'Process Group',
      'Color',
      'Sequence'
    ];

    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    SpreadsheetApp.flush();
    return buildResponse(true, null, 'BOM sheet initialized successfully.');
  } catch (error) {
    console.error('[initBOMSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize BOM sheet: ' + error.message);
  }
}

/**
 * Ensures the BOM sheet has a Remarks column (Col I).
 * Older sheets created before the Remarks feature only have 8 columns,
 * so this expands them in place without disturbing existing data.
 */
function ensureBOMRemarksColumn(sheet) {
  try {
    if (sheet.getLastColumn() < BOM_COL.REMARKS) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), BOM_COL.REMARKS - sheet.getLastColumn());
      sheet.getRange(1, BOM_COL.REMARKS)
        .setValue('Remarks')
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    console.error('[ensureBOMRemarksColumn] Error:', error.message);
  }
}

/**
 * Ensures the BOM sheet has a Process Group column (Col J).
 * Older sheets created before the sub-group feature only have 9 columns,
 * so this expands them in place without disturbing existing data.
 */
function ensureBOMProcessGroupColumn(sheet) {
  try {
    if (sheet.getLastColumn() < BOM_COL.PROCESS_GROUP) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), BOM_COL.PROCESS_GROUP - sheet.getLastColumn());
      sheet.getRange(1, BOM_COL.PROCESS_GROUP)
        .setValue('Process Group')
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    console.error('[ensureBOMProcessGroupColumn] Error:', error.message);
  }
}

/**
 * Ensures the BOM sheet has a Color column (Col K). Older sheets created
 * before color variants existed only have 10 columns, so this expands them
 * in place without disturbing existing data. A blank Color means the row
 * applies to every color of that Model — see module_bom.gs header comment.
 */
function ensureBOMColorColumn(sheet) {
  try {
    if (sheet.getLastColumn() < BOM_COL.COLOR) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), BOM_COL.COLOR - sheet.getLastColumn());
      sheet.getRange(1, BOM_COL.COLOR)
        .setValue('Color')
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    console.error('[ensureBOMColorColumn] Error:', error.message);
  }
}

/**
 * Ensures the BOM sheet has a Sequence column (Col L), used for manual
 * drag-and-drop ordering on the Products tab. Older sheets predating this
 * feature only have 11 columns, so this expands them in place. On first
 * run (column just added, or already present but entirely blank/zero —
 * e.g. a sheet created before this ran once), backfills every existing
 * product with a sequence based on its current first-appearance order in
 * the sheet, so legacy data gets a stable starting order instead of all
 * sharing sequence 0.
 */
function ensureBOMSequenceColumn(sheet) {
  try {
    if (sheet.getLastColumn() < BOM_COL.SEQUENCE) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), BOM_COL.SEQUENCE - sheet.getLastColumn());
      sheet.getRange(1, BOM_COL.SEQUENCE)
        .setValue('Sequence')
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const productIds = sheet.getRange(2, BOM_COL.PRODUCT_ID, lastRow - 1, 1).getValues();
    const sequences = sheet.getRange(2, BOM_COL.SEQUENCE, lastRow - 1, 1).getValues();

    const hasAnySequence = sequences.some(r => Number(r[0]) > 0);
    if (hasAnySequence) return;

    const seqByProductId = {};
    let nextSeq = 1;
    for (let i = 0; i < productIds.length; i++) {
      const pid = String(productIds[i][0] || '').trim();
      if (!pid) continue;
      if (!seqByProductId.hasOwnProperty(pid)) seqByProductId[pid] = nextSeq++;
      sequences[i][0] = seqByProductId[pid];
    }

    sheet.getRange(2, BOM_COL.SEQUENCE, sequences.length, 1).setValues(sequences);
  } catch (error) {
    console.error('[ensureBOMSequenceColumn] Error:', error.message);
  }
}

/**
 * Initializes the BOM Additional Costs sheet with correct headers.
 * Stores per-product dynamic/overhead costs (labor, paint, fitting, etc.)
 * Run once manually or triggered automatically when sheet is missing.
 */
function initBOMCostsSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(APP_CONFIG.SHEETS.BOM_COSTS);
    if (!sheet) {
      sheet = ss.insertSheet(APP_CONFIG.SHEETS.BOM_COSTS);
    }

    const headers = [
      'Product ID',
      'Cost Description',
      'Rate',
      'Process',
      'Contractor'
    ];

    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');

    SpreadsheetApp.flush();
    return buildResponse(true, null, 'BOM Additional Costs sheet initialized successfully.');
  } catch (error) {
    console.error('[initBOMCostsSheet] Error:', error.message);
    return buildResponse(false, null, 'Failed to initialize BOM Additional Costs sheet: ' + error.message);
  }
}

/**
 * Backfills the "Process" and "Contractor" columns on BOM Additional Costs
 * sheets created before the contractor rate card feature existed, so
 * legacy rows don't throw when read/written.
 */
function ensureBOMCostsExtraColumns(sheet) {
  try {
    if (sheet.getLastColumn() < BOM_COSTS_COL.CONTRACTOR_NAME) {
      const startCol = sheet.getLastColumn() + 1;
      sheet.insertColumnsAfter(sheet.getLastColumn(), BOM_COSTS_COL.CONTRACTOR_NAME - sheet.getLastColumn());
      sheet.getRange(1, startCol, 1, BOM_COSTS_COL.CONTRACTOR_NAME - startCol + 1)
        .setValues([['Process', 'Contractor']])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
    }
  } catch (error) {
    console.error('[ensureBOMCostsExtraColumns] Error:', error.message);
  }
}

/**
 * Retrieves all Product BOMs grouped by Product ID, including cost/rate
 * details. Requires a valid BOM access token (see verifyBOMAccess).
 * @param {string} token - Access token from verifyBOMAccess
 */
function getBOMData(token) {
  try {
    _requireBOMAccess(token);

    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.BOM);
    } catch (e) {
      initBOMSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.BOM);
    }

    ensureBOMRemarksColumn(sheet);
    ensureBOMProcessGroupColumn(sheet);
    ensureBOMColorColumn(sheet);
    ensureBOMSequenceColumn(sheet);

    const processNameById = {};
    (_getAllProcessRecords() || []).forEach(p => {
      processNameById[p.processId.toLowerCase()] = p.processName;
    });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return buildResponse(true, []);
    }

    const numCols = Math.max(sheet.getLastColumn(), BOM_COL.SEQUENCE);
    const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    const productMap = {};

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const productId = String(row[BOM_COL.PRODUCT_ID - 1] || '').trim();
      const productName = String(row[BOM_COL.PRODUCT_NAME - 1] || '').trim();
      const itemName = String(row[BOM_COL.ITEM_NAME - 1] || '').trim();

      if (!productId || !productName || !itemName) continue;

      if (!productMap[productId]) {
        productMap[productId] = {
          productId: productId,
          productName: productName,
          components: [],
          colors: [],
          totalCost: 0,
          totalQty: 0,
          additionalCosts: [],
          totalAdditionalCost: 0,
          remarks: String(row[BOM_COL.REMARKS - 1] || '').trim(),
          sequence: Number(row[BOM_COL.SEQUENCE - 1]) || 0
        };
      }

      const rate = Number(row[BOM_COL.RATE - 1]) || 0;
      const qty = Number(row[BOM_COL.QTY_PER_PRODUCT - 1]) || 0;
      const lineCost = rate * qty;
      const processId = String(row[BOM_COL.PROCESS_GROUP - 1] || '').trim();
      const color = String(row[BOM_COL.COLOR - 1] || '').trim();

      productMap[productId].components.push({
        itemName: itemName,
        size: String(row[BOM_COL.SIZE - 1] || '').trim(),
        narration: String(row[BOM_COL.NARRATION - 1] || '').trim(),
        rate: rate,
        vendor: String(row[BOM_COL.VENDOR - 1] || '').trim(),
        qtyPerProduct: qty,
        lineCost: lineCost,
        processId: processId,
        processGroup: processNameById[processId.toLowerCase()] || processId || 'General',
        color: color
      });

      if (color && !productMap[productId].colors.includes(color)) {
        productMap[productId].colors.push(color);
      }

      productMap[productId].totalCost += lineCost;
      productMap[productId].totalQty += qty;
    }

    // Load additional/dynamic costs (labor, paint, fitting, overheads, etc.)
    let costsSheet;
    try {
      costsSheet = getSheet(APP_CONFIG.SHEETS.BOM_COSTS);
    } catch (e) {
      initBOMCostsSheet();
      costsSheet = getSheet(APP_CONFIG.SHEETS.BOM_COSTS);
    }

    ensureBOMCostsExtraColumns(costsSheet);

    const costsLastRow = costsSheet.getLastRow();
    if (costsLastRow >= 2) {
      const costsData = costsSheet.getRange(2, 1, costsLastRow - 1, BOM_COSTS_COL.CONTRACTOR_NAME).getValues();
      for (let i = 0; i < costsData.length; i++) {
        const row = costsData[i];
        const productId = String(row[BOM_COSTS_COL.PRODUCT_ID - 1] || '').trim();
        const description = String(row[BOM_COSTS_COL.DESCRIPTION - 1] || '').trim();
        if (!productId || !description || !productMap[productId]) continue;

        const rate = Number(row[BOM_COSTS_COL.RATE - 1]) || 0;
        const processName = String(row[BOM_COSTS_COL.PROCESS_NAME - 1] || '').trim();
        const contractorName = String(row[BOM_COSTS_COL.CONTRACTOR_NAME - 1] || '').trim();
        productMap[productId].additionalCosts.push({ description: description, rate: rate, processName: processName, contractorName: contractorName });
        productMap[productId].totalAdditionalCost += rate;
      }
    }

    const products = Object.values(productMap);
    products.forEach(p => {
      p.grandTotal = p.totalCost + p.totalAdditionalCost;
    });
    // Sort by manual display order (drag-and-drop reorderable on the Products tab)
    products.sort((a, b) => a.sequence - b.sequence);

    return buildResponse(true, products);
  } catch (error) {
    if (error.message === 'UNAUTHORIZED_BOM') {
      return buildResponse(false, null, 'This section is password-protected.');
    }
    console.error('[getBOMData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load BOM data: ' + error.message);
  }
}

/**
 * Returns a Product ID (lowercased) -> final-stage Process Master record
 * map, derived from each product's existing BOM component rows (no new
 * field to maintain — a product's BOM components are already tagged with
 * which Process Master row consumes them via BOM_COL.PROCESS_GROUP; a
 * product spans several real Processes, see module header).
 *
 * Only resolves a product when exactly one of the distinct Processes
 * referenced by its BOM rows is flagged isFinalStage — 0 or 2+ matches
 * are ambiguous and intentionally omitted so callers (PI -> Production
 * auto-queueing) fall back to a manual flow instead of guessing which
 * Process actually produces the finished, dispatchable good.
 * @returns {Object} { [productIdLower]: { processId, processName, sequence, lotPrefix, isFinalStage, outputItemName } }
 */
function getBOMFinalStageProcessMap() {
  const map = {};
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.BOM);
    } catch (e) {
      return map;
    }
    ensureBOMProcessGroupColumn(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return map;

    const data = sheet.getRange(2, 1, lastRow - 1, BOM_COL.PROCESS_GROUP).getValues();
    const processIdsByProduct = {};
    data.forEach(row => {
      const productId = String(row[BOM_COL.PRODUCT_ID - 1] || '').trim();
      const processId = String(row[BOM_COL.PROCESS_GROUP - 1] || '').trim();
      if (!productId || !processId) return;
      const key = productId.toLowerCase();
      if (!processIdsByProduct[key]) processIdsByProduct[key] = new Set();
      processIdsByProduct[key].add(processId.toLowerCase());
    });

    const finalStageById = {};
    (typeof _getAllProcessRecords === 'function' ? _getAllProcessRecords() : []).forEach(p => {
      if (p.isFinalStage) finalStageById[p.processId.toLowerCase()] = p;
    });

    Object.keys(processIdsByProduct).forEach(productKey => {
      const matches = Array.from(processIdsByProduct[productKey])
        .map(pid => finalStageById[pid])
        .filter(Boolean);
      if (matches.length === 1) {
        map[productKey] = matches[0];
      }
    });

    return map;
  } catch (e) {
    console.error('[getBOMFinalStageProcessMap] Error:', e.message);
    return map;
  }
}

/**
 * Retrieves a cost-free view of all Product BOMs (Product ID, Product Name,
 * and component item/qty/vendor details, but no rates or costs). Available
 * to all users so the Production tab can populate its Product dropdown and
 * generate production sheets without exposing pricing data.
 */
function getBOMProductionData() {
  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.BOM);
    } catch (e) {
      initBOMSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.BOM);
    }

    ensureBOMRemarksColumn(sheet);
    ensureBOMProcessGroupColumn(sheet);
    ensureBOMColorColumn(sheet);
    ensureBOMSequenceColumn(sheet);

    const processNameById = {};
    (_getAllProcessRecords() || []).forEach(p => {
      processNameById[p.processId.toLowerCase()] = p.processName;
    });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return buildResponse(true, []);
    }

    const numCols = Math.max(sheet.getLastColumn(), BOM_COL.COLOR);
    const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    const productMap = {};

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const productId = String(row[BOM_COL.PRODUCT_ID - 1] || '').trim();
      const productName = String(row[BOM_COL.PRODUCT_NAME - 1] || '').trim();
      const itemName = String(row[BOM_COL.ITEM_NAME - 1] || '').trim();

      if (!productId || !productName || !itemName) continue;

      if (!productMap[productId]) {
        productMap[productId] = {
          productId: productId,
          productName: productName,
          components: [],
          colors: []
        };
      }

      const processId = String(row[BOM_COL.PROCESS_GROUP - 1] || '').trim();
      const color = String(row[BOM_COL.COLOR - 1] || '').trim();
      productMap[productId].components.push({
        itemName: itemName,
        size: String(row[BOM_COL.SIZE - 1] || '').trim(),
        narration: String(row[BOM_COL.NARRATION - 1] || '').trim(),
        vendor: String(row[BOM_COL.VENDOR - 1] || '').trim(),
        qtyPerProduct: Number(row[BOM_COL.QTY_PER_PRODUCT - 1]) || 0,
        processId: processId,
        processGroup: processNameById[processId.toLowerCase()] || processId || 'General',
        color: color
      });

      if (color && !productMap[productId].colors.includes(color)) {
        productMap[productId].colors.push(color);
      }
    }

    const products = Object.values(productMap);
    products.sort((a, b) => b.productId.localeCompare(a.productId));

    return buildResponse(true, products);
  } catch (error) {
    console.error('[getBOMProductionData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load product list: ' + error.message);
  }
}

/**
 * Auto-generates the next sequential Product ID.
 * Format: PRD-1001, PRD-1002, ...
 * Requires a valid BOM access token (see verifyBOMAccess).
 * @param {string} token - Access token from verifyBOMAccess
 */
function getNextProductId(token) {
  try {
    _requireBOMAccess(token);

    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.BOM);
    } catch (e) {
      initBOMSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.BOM);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return 'PRD-1001';

    const ids = sheet.getRange(2, BOM_COL.PRODUCT_ID, lastRow - 1, 1).getValues();
    let maxNum = 1000;

    ids.forEach(row => {
      const idStr = String(row[0] || '').trim();
      const match = idStr.match(/^PRD-(\d+)$/i);
      if (match) {
        const num = parseInt(match[1], 10);
        if (num > maxNum) maxNum = num;
      }
    });

    return 'PRD-' + (maxNum + 1);
  } catch (error) {
    console.error('[getNextProductId] Error:', error.message);
    return 'PRD-1001'; // Fallback
  }
}

/**
 * Creates a new Product BOM or updates an existing one.
 * Uses document-level lock and transaction-like block replacement.
 * Requires a valid BOM access token (see verifyBOMAccess).
 * @param {Object} formData
 * @param {string} token - Access token from verifyBOMAccess
 */
function saveBOM(formData, token) {
  try {
    _requireBOMAccess(token);
  } catch (error) {
    return buildResponse(false, null, 'This section is password-protected.');
  }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(BOM_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    let sheet;
    try {
      sheet = getSheet(APP_CONFIG.SHEETS.BOM);
    } catch (e) {
      initBOMSheet();
      sheet = getSheet(APP_CONFIG.SHEETS.BOM);
    }

    ensureBOMRemarksColumn(sheet);
    ensureBOMProcessGroupColumn(sheet);
    ensureBOMColorColumn(sheet);

    let costsSheet;
    try {
      costsSheet = getSheet(APP_CONFIG.SHEETS.BOM_COSTS);
    } catch (e) {
      initBOMCostsSheet();
      costsSheet = getSheet(APP_CONFIG.SHEETS.BOM_COSTS);
    }

    ensureBOMCostsExtraColumns(costsSheet);

    const productName = sanitizeString(formData.productName, 'productName');
    if (!productName) {
      return buildResponse(false, null, 'Product Name is required.');
    }

    const remarks = sanitizeString(formData.remarks || '', 'remarks').slice(0, APP_CONFIG.VALIDATION.MAX_REMARKS_LENGTH);

    let components;
    try {
      components = typeof formData.components === 'string'
        ? JSON.parse(formData.components)
        : (formData.components || []);
    } catch (e) {
      return buildResponse(false, null, 'Invalid components data format.');
    }

    if (!Array.isArray(components) || components.length === 0) {
      return buildResponse(false, null, 'BOM must contain at least one component.');
    }

    let additionalCosts;
    try {
      additionalCosts = typeof formData.additionalCosts === 'string'
        ? JSON.parse(formData.additionalCosts)
        : (formData.additionalCosts || []);
    } catch (e) {
      return buildResponse(false, null, 'Invalid additional costs data format.');
    }

    if (!Array.isArray(additionalCosts)) {
      additionalCosts = [];
    }

    // Check if this is an Edit
    const isEdit = !!formData.productId;
    let productId = isEdit ? sanitizeString(formData.productId, 'productId') : '';
    let sequence;

    if (isEdit) {
      // Confirm product exists in sheet, and capture its current Sequence
      // (before its rows get deleted below) so editing never moves it.
      const lastDataRow = Math.max(1, sheet.getLastRow() - 1);
      const currentIds = sheet.getRange(2, BOM_COL.PRODUCT_ID, lastDataRow, 1).getValues();
      const currentSeqs = sheet.getRange(2, BOM_COL.SEQUENCE, lastDataRow, 1).getValues();
      const matchIdx = currentIds.findIndex(r => String(r[0]).trim().toLowerCase() === productId.toLowerCase());
      if (matchIdx === -1) {
        return buildResponse(false, null, `Product BOM with ID "${productId}" not found.`);
      }
      sequence = Number(currentSeqs[matchIdx][0]) || 0;

      // Delete existing component rows for this Product ID
      deleteRowsById(productId, sheet, 2, BOM_COL.PRODUCT_ID);

      // Delete existing additional cost rows for this Product ID
      deleteRowsById(productId, costsSheet, 2, BOM_COSTS_COL.PRODUCT_ID);
    } else {
      // Generate new Product ID and append it to the end of the manual order
      productId = getNextProductId(token);
      const lastRow = sheet.getLastRow();
      let maxSeq = 0;
      if (lastRow >= 2) {
        sheet.getRange(2, BOM_COL.SEQUENCE, lastRow - 1, 1).getValues().forEach(r => {
          const v = Number(r[0]) || 0;
          if (v > maxSeq) maxSeq = v;
        });
      }
      sequence = maxSeq + 1;
    }

    // Prepare rows to insert
    const rowsToWrite = [];
    components.forEach(comp => {
      const itemName = sanitizeString(comp.itemName, 'itemName');
      if (!itemName) return;

      const size = sanitizeString(comp.size || '', 'size');
      const narration = sanitizeString(comp.narration || '', 'narration');
      const rate = validateNumber(comp.rate, 0, 10000000);
      const vendor = sanitizeString(comp.vendor || '', 'vendor');
      const qtyPerProduct = validateNumber(comp.qtyPerProduct, 0.0001, 1000000);
      const processId = sanitizeString(comp.processId || '', 'processId');
      const color = sanitizeString(comp.color || '', 'color');

      rowsToWrite.push([
        productId,
        productName,
        itemName,
        size,
        narration,
        rate,
        vendor,
        qtyPerProduct,
        remarks,
        processId,
        color,
        sequence
      ]);
    });

    if (rowsToWrite.length === 0) {
      return buildResponse(false, null, 'Please specify valid component items.');
    }

    // Append component rows
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToWrite.length, 12).setValues(rowsToWrite);

    // Prepare and append additional/dynamic cost rows (labor, paint, fitting, overheads, etc.)
    const costRowsToWrite = [];
    additionalCosts.forEach(cost => {
      const description = sanitizeString(cost.description || '', 'description');
      if (!description) return;

      const rate = validateNumber(cost.rate, 0, 10000000);
      const processName = sanitizeString(cost.processName || '', 'processName');
      const contractorName = sanitizeString(cost.contractorName || '', 'contractorName');
      costRowsToWrite.push([productId, description, rate, processName, contractorName]);
    });

    if (costRowsToWrite.length > 0) {
      const costsStartRow = costsSheet.getLastRow() + 1;
      costsSheet.getRange(costsStartRow, 1, costRowsToWrite.length, 5).setValues(costRowsToWrite);
    }

    SpreadsheetApp.flush();

    const logMsg = isEdit
      ? `BOM updated: ${productName} (${productId}) with ${rowsToWrite.length} components and ${costRowsToWrite.length} additional cost lines.`
      : `BOM created: ${productName} (${productId}) with ${rowsToWrite.length} components and ${costRowsToWrite.length} additional cost lines.`;

    logAction(isEdit ? 'UPDATE' : 'CREATE', APP_CONFIG.SHEETS.BOM, productId, logMsg, 'SUCCESS');

    return buildResponse(true, { productId: productId, productName: productName }, isEdit ? 'BOM updated successfully.' : 'Product BOM saved successfully.');
  } catch (error) {
    console.error('[saveBOM] Error:', error.message);
    logAction('ERROR', 'saveBOM', formData.productId || 'NEW', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to save BOM: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes all component rows of a Product BOM.
 * Before deleting, checks if the Product ID is in use by the Production ledger.
 * Requires a valid BOM access token (see verifyBOMAccess).
 * @param {string} productId
 * @param {string} token - Access token from verifyBOMAccess
 */
function deleteBOM(productId, token) {
  try {
    _requireBOMAccess(token);
  } catch (error) {
    return buildResponse(false, null, 'This section is password-protected.');
  }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(BOM_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const bomSheet = getSheet(APP_CONFIG.SHEETS.BOM);
    if (!bomSheet) throw new Error('BOM sheet not found.');

    const prodIdClean = String(productId || '').trim();
    if (!prodIdClean) return buildResponse(false, null, 'Product ID is required.');

    // Check if used in Production
    let prodSheet;
    try {
      prodSheet = getSheet(APP_CONFIG.SHEETS.PRODUCTION);
    } catch (e) {
      // Production sheet does not exist yet, so it cannot be in use
    }

    if (prodSheet) {
      const pLastRow = prodSheet.getLastRow();
      if (pLastRow >= 2) {
        const prodRefs = prodSheet.getRange(2, PRODUCTION_COL.PRODUCT_ID, pLastRow - 1, 1).getValues();
        const inUse = prodRefs.some(row => String(row[0]).trim().toLowerCase() === prodIdClean.toLowerCase());
        if (inUse) {
          return buildResponse(
            false,
            null,
            `Cannot delete BOM: Product ID "${prodIdClean}" has already been used in Production lots. Clear production logs first.`
          );
        }
      }
    }

    const rowsDeleted = deleteRowsById(prodIdClean, bomSheet, 2, BOM_COL.PRODUCT_ID);

    // Also remove any associated additional/dynamic cost rows
    try {
      const costsSheet = getSheet(APP_CONFIG.SHEETS.BOM_COSTS);
      deleteRowsById(prodIdClean, costsSheet, 2, BOM_COSTS_COL.PRODUCT_ID);
    } catch (e) {
      // BOM Additional Costs sheet does not exist yet, nothing to clean up
    }

    SpreadsheetApp.flush();

    const msg = `BOM for Product ID "${prodIdClean}" deleted (${rowsDeleted} rows removed).`;
    logAction('DELETE', APP_CONFIG.SHEETS.BOM, prodIdClean, msg, 'SUCCESS');

    return buildResponse(true, null, msg);
  } catch (error) {
    console.error('[deleteBOM] Error:', error.message);
    logAction('ERROR', 'deleteBOM', productId, error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete BOM: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes multiple Product BOMs in a single batch.
 * Any product currently referenced in the Production ledger is skipped
 * (and reported back) rather than failing the entire batch.
 * Requires a valid BOM access token (see verifyBOMAccess).
 * @param {Array<string>} productIds
 * @param {string} token - Access token from verifyBOMAccess
 */
function deleteBOMsBulk(productIds, token) {
  try {
    _requireBOMAccess(token);
  } catch (error) {
    return buildResponse(false, null, 'This section is password-protected.');
  }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(BOM_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const bomSheet = getSheet(APP_CONFIG.SHEETS.BOM);
    if (!bomSheet) throw new Error('BOM sheet not found.');

    const requested = (productIds || []).map(p => String(p || '').trim()).filter(Boolean);
    if (requested.length === 0) {
      return buildResponse(true, null, 'No BOMs selected.');
    }

    // Determine which Product IDs are currently in use by Production
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
        const prodRefs = prodSheet.getRange(2, PRODUCTION_COL.PRODUCT_ID, pLastRow - 1, 1).getValues();
        const usedIds = new Set(prodRefs.map(row => String(row[0]).trim().toLowerCase()));
        requested.forEach(id => {
          if (usedIds.has(id.toLowerCase())) inUseSet.add(id);
        });
      }
    }

    const toDelete = requested.filter(id => !inUseSet.has(id));
    const targetSet = new Set(toDelete.map(id => id));

    let rowsDeleted = 0;
    if (targetSet.size > 0) {
      rowsDeleted = _rewriteWithoutMatchingRowsBulk(bomSheet, 2, BOM_COL.PRODUCT_ID, targetSet).rowsDeleted;

      // Also remove any associated additional/dynamic cost rows
      try {
        const costsSheet = getSheet(APP_CONFIG.SHEETS.BOM_COSTS);
        _rewriteWithoutMatchingRowsBulk(costsSheet, 2, BOM_COSTS_COL.PRODUCT_ID, targetSet);
      } catch (e) {
        // BOM Additional Costs sheet does not exist yet, nothing to clean up
      }

      SpreadsheetApp.flush();
    }

    let msg = `Deleted ${toDelete.length} BOM(s) (${rowsDeleted} rows removed).`;
    if (inUseSet.size > 0) {
      msg += ` Skipped ${inUseSet.size} BOM(s) still in use by Production: ${Array.from(inUseSet).join(', ')}.`;
    }
    logAction('BULK_DELETE', APP_CONFIG.SHEETS.BOM, 'multiple', msg, 'SUCCESS');

    return buildResponse(true, { skipped: Array.from(inUseSet) }, msg);
  } catch (error) {
    console.error('[deleteBOMsBulk] Error:', error.message);
    logAction('ERROR', 'deleteBOMsBulk', 'multiple', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to delete BOMs: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Persists a manual drag-and-drop reorder from the Products tab. Rewrites
 * Sequence 1..N for every row matching a Product ID in orderedProductIds,
 * in that order (every component row of a product shares the same value).
 * Requires a valid BOM access token (see verifyBOMAccess).
 * @param {string[]} orderedProductIds - Product IDs in their new desired order
 * @param {string} token - Access token from verifyBOMAccess
 */
function reorderBOM(orderedProductIds, token) {
  try {
    _requireBOMAccess(token);
  } catch (error) {
    return buildResponse(false, null, 'This section is password-protected.');
  }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(BOM_LOCK_TIMEOUT_MS)) {
    return buildResponse(false, null, 'System is busy. Please try again.');
  }

  try {
    const sheet = getSheet(APP_CONFIG.SHEETS.BOM);
    if (!sheet) throw new Error('BOM sheet not found.');

    const order = (orderedProductIds || []).map(id => String(id || '').trim()).filter(Boolean);
    if (order.length === 0) {
      return buildResponse(false, null, 'No product order provided.');
    }

    const newSeqById = {};
    order.forEach((id, i) => { newSeqById[id.toLowerCase()] = i + 1; });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return buildResponse(true, null, 'No products to reorder.');
    }

    const productIds = sheet.getRange(2, BOM_COL.PRODUCT_ID, lastRow - 1, 1).getValues();
    const sequences = sheet.getRange(2, BOM_COL.SEQUENCE, lastRow - 1, 1).getValues();
    for (let i = 0; i < productIds.length; i++) {
      const id = String(productIds[i][0] || '').trim().toLowerCase();
      if (newSeqById.hasOwnProperty(id)) {
        sequences[i][0] = newSeqById[id];
      }
    }
    sheet.getRange(2, BOM_COL.SEQUENCE, sequences.length, 1).setValues(sequences);

    SpreadsheetApp.flush();

    const msg = `Reordered ${order.length} product(s).`;
    logAction('UPDATE', APP_CONFIG.SHEETS.BOM, 'multiple', msg, 'SUCCESS');

    return buildResponse(true, null, msg);
  } catch (error) {
    console.error('[reorderBOM] Error:', error.message);
    logAction('ERROR', 'reorderBOM', 'multiple', error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to reorder products: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Rewrites BOM component rows referencing a renamed/merged item, so BOM
 * recipes stay attributed to the item's current name + size.
 *
 * Runs under the caller's existing document lock — does not acquire its own.
 *
 * @param {string} oldName
 * @param {string} oldSize
 * @param {string} newName
 * @param {string} newSize
 */
function backfillBOMItemRefs(oldName, oldSize, newName, newSize) {
  const sheet = getSheet(APP_CONFIG.SHEETS.BOM);
  if (!sheet) return;

  const startRow = APP_CONFIG.BOM_SETTINGS.DATA_START_ROW;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  const numRows = lastRow - startRow + 1;
  const range = sheet.getRange(startRow, BOM_COL.ITEM_NAME, numRows, 2);
  const values = range.getValues();

  const tOldName = String(oldName || '').trim().toLowerCase();
  const tOldSize = String(oldSize || '').trim().toLowerCase();

  let changed = false;
  for (let i = 0; i < values.length; i++) {
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
