/**
 * ═══════════════════════════════════════════════════════════════════════════
 * config.gs — CENTRAL CONFIGURATION
 * 
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Single source of truth for all application settings, sheet names, column
 * mappings, and magic constants. Using Object.freeze() to ensure immutability
 * and prevent accidental runtime modifications.
 * 
 * Maintenance:
 * ───────────────────────────────────────────────────────────────────────────
 * - When adding new sheets: update SHEETS and corresponding COL objects
 * - When reordering columns: update the COL constants and all associated
 *   read/write functions
 * - When changing row positions: update *_SETTINGS objects
 * - Never modify frozen objects at runtime — extend config instead
 * 
 * Usage:
 * ───────────────────────────────────────────────────────────────────────────
 * const sheetName = APP_CONFIG.SHEETS.PO;
 * const poNumCol = PO_COL.PO_NUMBER;
 * 
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ─────────────────────────────────────────────────────────────────────────
// APPLICATION METADATA
// ─────────────────────────────────────────────────────────────────────────
const APP_CONFIG = Object.freeze({
  /**
   * Application display name
   */
  APP_TITLE: 'Maharaja Bikes ERP',

  /**
   * Application version for tracking and debugging
   * Update when making significant changes
   */
  APP_VERSION: '1.0.0',

  /**
   * Environment flag (development/staging/production)
   */
  ENVIRONMENT: 'production',

  /**
   * Sheet names — Master reference for all sheet operations
   * Must exactly match the sheet names in the spreadsheet
   */
  SHEETS: Object.freeze({
    PO: 'PO Tracker',
    BILL: 'Bill Ledger',
    RETURN: 'Return Ledger',
    ITEMS: 'Items Master',
    RATE_HISTORY: 'Rate History',
    VENDORS: 'Vendors',
    BOM: 'BOM',
    BOM_COSTS: 'BOM Additional Costs',
    PRODUCTION: 'Production',
    STOCK: 'Stock',
    CLIENTS: 'Clients',
    CLIENT_ORDERS: 'Client Orders',
    DISPATCH: 'Dispatch',
    PROCESS_MASTER: 'Process Master',
    PROCESS_COMPONENTS: 'Process Components',
    PROCESS_COLOR_LINKS: 'Process Color Links',
    PROCESS_COLOR_OVERRIDES: 'Process Color Overrides',
    WAREHOUSE_POOL: 'Warehouse Pool',
    WAREHOUSE_POOL_OPENING: 'Warehouse Pool Opening',
    UNITS: 'Unit Master',
    COLOR_MASTER: 'Color Master',
    MODEL_MASTER: 'Model Master',
    PROCESS_TYPE_MASTER: 'Process Type Master',
    CONTRACTORS: 'Contractors',
    CONTRACTOR_RATES: 'Contractor Rates',
    CONTRACTOR_PAYMENTS: 'Contractor Payments',
    LOGS: 'Logs',  // For error tracking and audit
    WASTAGE: 'Wastage Log',  // Component-wise wastage / loss tracking
    ISSUE: 'Issued Stock Log'  // Ad-hoc item issuance outside any Process BOM / Production lot
  }),

  /**
   * Settings for PO sheet structure
   */
  PO_SETTINGS: Object.freeze({
    // Number of header rows (titles, labels, etc.)
    HEADER_ROWS: 2,

    // Row number where data begins (1-based)
    DATA_START_ROW: 3,

    // Column index for PO Number (used as unique key)
    ID_COL_INDEX: 1,

    // Maximum rows before needing pagination
    MAX_ROWS_PER_QUERY: 10000
  }),

  /**
   * Settings for Bill sheet structure
   */
  BILL_SETTINGS: Object.freeze({
    HEADER_ROWS: 1,
    DATA_START_ROW: 2,
    PO_NUM_COL_INDEX: 1,
    MAX_ROWS_PER_QUERY: 10000
  }),

  /**
   * Settings for Return (vendor goods return / debit note) sheet structure
   */
  RETURN_SETTINGS: Object.freeze({
    HEADER_ROWS: 1,
    DATA_START_ROW: 2,
    MAX_ROWS_PER_QUERY: 10000
  }),

  /**
   * Settings for Wastage Log sheet structure
   */
  WASTAGE_SETTINGS: Object.freeze({
    HEADER_ROWS: 1,
    DATA_START_ROW: 2,
    MAX_ROWS_PER_QUERY: 10000
  }),

  /**
   * Settings for Issued Stock Log sheet structure
   */
  ISSUE_SETTINGS: Object.freeze({
    HEADER_ROWS: 1,
    DATA_START_ROW: 2,
    MAX_ROWS_PER_QUERY: 10000
  }),

  /**
   * Settings for Items Master sheet
   */
  ITEMS_SETTINGS: Object.freeze({
    HEADER_ROWS: 1,
    DATA_START_ROW: 2,
    NAME_COL_INDEX: 1,  // Unique identifier for items
    MAX_ROWS_PER_QUERY: 5000
  }),

  /**
   * Settings for BOM sheet structure
   */
  BOM_SETTINGS: Object.freeze({
    HEADER_ROWS: 1,
    DATA_START_ROW: 2,
    ID_COL_INDEX: 1,
    MAX_ROWS_PER_QUERY: 10000
  }),

  /**
   * Settings for Production sheet structure
   */
  PRODUCTION_SETTINGS: Object.freeze({
    HEADER_ROWS: 1,
    DATA_START_ROW: 2,
    DATE_COL_INDEX: 1,
    MAX_ROWS_PER_QUERY: 10000
  }),

  /**
   * Settings for Unit Master sheet structure
   */
  UNITS_SETTINGS: Object.freeze({
    HEADER_ROWS: 1,
    DATA_START_ROW: 2,
    UNIT_NAME_COL_INDEX: 1,
    MAX_ROWS_PER_QUERY: 500
  }),

  /**
   * Settings for Stock sheet structure
   */
  STOCK_SETTINGS: Object.freeze({
    HEADER_ROWS: 1,
    DATA_START_ROW: 2,
    ITEM_NAME_COL_INDEX: 1,
    MAX_ROWS_PER_QUERY: 10000
  }),

  /**
   * Settings for Clients sheet structure
   */
  CLIENTS_SETTINGS: Object.freeze({
    HEADER_ROWS: 1,
    DATA_START_ROW: 2,
    ID_COL_INDEX: 1,
    MAX_ROWS_PER_QUERY: 5000
  }),

  /**
   * Settings for Client Orders (PI / Estimates) sheet structure
   */
  CLIENT_ORDERS_SETTINGS: Object.freeze({
    HEADER_ROWS: 1,
    DATA_START_ROW: 2,
    ID_COL_INDEX: 1,
    MAX_ROWS_PER_QUERY: 10000
  }),

  /**
   * Settings for Dispatch sheet structure
   */
  DISPATCH_SETTINGS: Object.freeze({
    HEADER_ROWS: 1,
    DATA_START_ROW: 2,
    ID_COL_INDEX: 1,
    MAX_ROWS_PER_QUERY: 10000
  }),

  /**
   * Settings for Process Master sheet structure
   */
  PROCESS_SETTINGS: Object.freeze({
    HEADER_ROWS: 1,
    DATA_START_ROW: 2,
    ID_COL_INDEX: 1,
    MAX_ROWS_PER_QUERY: 1000
  }),

  /**
   * Settings for Process Components sheet structure
   */
  PROCESS_COMPONENTS_SETTINGS: Object.freeze({
    HEADER_ROWS: 1,
    DATA_START_ROW: 2,
    ID_COL_INDEX: 1,
    MAX_ROWS_PER_QUERY: 5000
  }),

  /**
   * Settings for Contractors sheet structure
   */
  CONTRACTORS_SETTINGS: Object.freeze({
    HEADER_ROWS: 1,
    DATA_START_ROW: 2,
    ID_COL_INDEX: 1,
    MAX_ROWS_PER_QUERY: 5000
  }),

  /**
   * Settings for Contractor Rates sheet structure
   */
  CONTRACTOR_RATES_SETTINGS: Object.freeze({
    HEADER_ROWS: 1,
    DATA_START_ROW: 2,
    ID_COL_INDEX: 1,
    MAX_ROWS_PER_QUERY: 5000
  }),

  /**
   * Settings for Contractor Payments sheet structure
   */
  CONTRACTOR_PAYMENTS_SETTINGS: Object.freeze({
    HEADER_ROWS: 1,
    DATA_START_ROW: 2,
    ID_COL_INDEX: 2,
    MAX_ROWS_PER_QUERY: 10000
  }),

  /**
   * API rate limiting and timeouts
   */
  API: Object.freeze({
    REQUEST_TIMEOUT_MS: 30000,  // 30 seconds
    MAX_RETRIES: 3,
    RETRY_DELAY_MS: 1000
  }),

  /**
   * Validation constraints
   */
  VALIDATION: Object.freeze({
    PO_NUMBER_MIN: 1000,
    PO_NUMBER_MAX: 999999,
    MIN_ITEM_QTY: 0.01,
    MAX_ITEM_QTY: 999999,
    MIN_PRICE: 0,
    MAX_PRICE: 999999999,
    MAX_VENDOR_NAME_LENGTH: 100,
    MAX_REMARKS_LENGTH: 500
  })
});

// ─────────────────────────────────────────────────────────────────────────
// PO SHEET COLUMN MAPPINGS (1-based indexing)
// 
// Column positions are frozen to maintain database integrity.
// Any changes require migration of existing data.
//
// Current Structure:
// A | B | C | D | E | F | G | H | I | J | K | L | M | N | O | P
// 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10| 11| 12| 13| 14| 15| 16
// ─────────────────────────────────────────────────────────────────────────

/**
 * PO status vocabulary — the one source of truth for the 3 derived status
 * strings module_po.js#_attachPoStatus computes from the Bill ledger.
 * Mirrored client-side (unchanged) in Script_ApiCore.html, shared by both
 * the desktop (Script_PO.html) and mobile (Mobile_Script.html) shells, so
 * badge/chip classes and filter logic key off this instead of repeating the
 * literal strings in three places.
 */
const PO_STATUS = Object.freeze({
  ISSUED: 'PO Issued',
  PARTIAL: 'Partially Received',
  COMPLETED: 'Completed'
});

const PO_COL = Object.freeze({
  PO_NUMBER:        1,      // Unique identifier
  PO_DATE:          2,      // DD/MM/YYYY
  VENDOR:           3,      // Vendor name
  CONTACT:          4,      // Contact location
  ITEM_NAME:        5,      // Item name from master
  NARRATION:        6,      // Item description
  SIZE:             7,      // Item size/variant
  QTY:              8,      // Quantity ordered
  UNIT:             9,      // Unit (Pcs, Kg, L, etc.)
  PO_DESCRIPTION:   10,     // PO-level description
  PO_REMARKS:       11,     // PO-level remarks
  PRICE:            12,     // Unit price
  TOTAL:            13,     // Line total (QTY × PRICE)
  TOTAL_QTY:        14,     // Total items in PO
  SUPPLIER_REMARKS: 15,     // Supplier reference/notes
  BASE_QTY:         16,     // QTY converted to the item's Base Unit (e.g. Pcs) — drives stock/BOM math
  BASE_RATE:        17      // PRICE converted to a per-Base-Unit rate
});

/**
 * Reverse mapping for PO columns (name -> index)
 * Useful for dynamic lookups
 */
const PO_COL_NAMES = Object.freeze({
  'poNumber': PO_COL.PO_NUMBER,
  'poDate': PO_COL.PO_DATE,
  'vendor': PO_COL.VENDOR,
  'contact': PO_COL.CONTACT,
  'itemName': PO_COL.ITEM_NAME,
  'narration': PO_COL.NARRATION,
  'size': PO_COL.SIZE,
  'qty': PO_COL.QTY,
  'unit': PO_COL.UNIT,
  'poDescription': PO_COL.PO_DESCRIPTION,
  'poRemarks': PO_COL.PO_REMARKS,
  'price': PO_COL.PRICE,
  'total': PO_COL.TOTAL,
  'totalQty': PO_COL.TOTAL_QTY,
  'supplierRemarks': PO_COL.SUPPLIER_REMARKS,
  'baseQty': PO_COL.BASE_QTY,
  'baseRate': PO_COL.BASE_RATE
});

// ─────────────────────────────────────────────────────────────────────────
// BILL SHEET COLUMN MAPPINGS (1-based indexing)
//
// Current Structure:
// A | B | C | D | E | F | G | H | I | J | K | L | M
// 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10| 11| 12| 13
// ─────────────────────────────────────────────────────────────────────────
const BILL_COL = Object.freeze({
  PO_NUMBER:      1,      // Reference to PO or 'DIRECT'
  BILL_NUMBER:    2,      // Unique bill identifier
  BILL_DATE:      3,      // DD/MM/YYYY
  VENDOR:         4,      // Vendor name
  CONTACT:        5,      // Contact location
  ITEM_NAME:      6,      // Item name
  SIZE:           7,      // Item size/variant
  NARRATION:      8,      // Item narration/description
  QTY:            9,      // Quantity billed
  UNIT:           10,     // Unit (Pcs, Kg, L, etc.)
  PRICE:          11,     // Unit price
  GST:            12,     // GST rate (%)
  TOTAL:          13,     // Line total with GST
  REMARKS:        14,     // Bill remarks
  BASE_QTY:       15,     // QTY converted to the item's Base Unit (e.g. Pcs) — drives stock/BOM math
  BASE_RATE:      16,     // PRICE converted to a per-Base-Unit rate
  AFFECTS_STOCK:  17,     // 'Y' (default/legacy) or 'N' — 'N' means this line is recorded for the
                           // ledger only and is excluded from Stock's Billed Qty sum (see
                           // checkStockAdjustmentConflicts in module_stock.js)
  BILL_TYPE:      18,     // 'GOODS' (default/legacy) or 'LABOR' — a Labor Job (job-work) bill from
                           // a contractor. VENDOR holds the contractor's name for these rows.
  PROCESS_NAME:   19,     // LABOR rows only: the Process (Process Master) this job-work line is
                           // for. ITEM_NAME is forced server-side to that process's Output Item
                           // Name — see saveBill.
  COLOR:          20,     // LABOR rows only: the color produced/processed on this line, picked
                           // from that Process's color sub-groups (getProcessColorGroups).
  ISSUING_PARTY:  21,     // LABOR rows only: optional free-text — who issued the raw material/job
                           // to the contractor for this bill (bill-header field, same value on
                           // every row of one bill, like VENDOR/CONTACT).
  MANUFACTURING_VENDOR: 22 // LABOR rows only: optional free-text — the vendor who manufactured the
                           // raw item(s) processed on this job-work bill.
});

/**
 * Reverse mapping for Bill columns
 */
const BILL_COL_NAMES = Object.freeze({
  'poNumber': BILL_COL.PO_NUMBER,
  'billNumber': BILL_COL.BILL_NUMBER,
  'billDate': BILL_COL.BILL_DATE,
  'vendor': BILL_COL.VENDOR,
  'contact': BILL_COL.CONTACT,
  'itemName': BILL_COL.ITEM_NAME,
  'size': BILL_COL.SIZE,
  'narration': BILL_COL.NARRATION,
  'qty': BILL_COL.QTY,
  'unit': BILL_COL.UNIT,
  'price': BILL_COL.PRICE,
  'gst': BILL_COL.GST,
  'total': BILL_COL.TOTAL,
  'remarks': BILL_COL.REMARKS,
  'baseQty': BILL_COL.BASE_QTY,
  'baseRate': BILL_COL.BASE_RATE,
  'affectsStock': BILL_COL.AFFECTS_STOCK,
  'billType': BILL_COL.BILL_TYPE,
  'processName': BILL_COL.PROCESS_NAME,
  'color': BILL_COL.COLOR,
  'issuingParty': BILL_COL.ISSUING_PARTY,
  'manufacturingVendor': BILL_COL.MANUFACTURING_VENDOR
});

// ─────────────────────────────────────────────────────────────────────────
// RETURN (VENDOR GOODS RETURN / DEBIT NOTE) SHEET COLUMN MAPPINGS
// (1-based indexing)
//
// Records stock sent back to a vendor (defective, excess, wrong item, etc.)
// — the mirror image of the Bill Ledger. QTY/BASE_QTY here are debited from
// Stock instead of credited (see module_stock.js), and TOTAL reduces the
// amount owed to the vendor.
// ─────────────────────────────────────────────────────────────────────────
const RETURN_COL = Object.freeze({
  RETURN_NUMBER:  1,      // Unique return/debit-note identifier
  RETURN_DATE:    2,      // DD/MM/YYYY
  VENDOR:         3,      // Vendor name
  CONTACT:        4,      // Contact location
  BILL_NUMBER:    5,      // Optional reference to the original Bill Ledger entry
  ITEM_NAME:      6,      // Item name
  SIZE:           7,      // Item size/variant
  NARRATION:      8,      // Item narration/description
  QTY:            9,      // Quantity returned
  UNIT:           10,     // Unit (Pcs, Kg, L, etc.)
  PRICE:          11,     // Unit credit/refund rate
  TOTAL:          12,     // Line total (QTY x PRICE)
  REASON:         13,     // Why the goods were returned (defective, excess, wrong item, ...)
  REMARKS:        14,     // Return-level remarks
  BASE_QTY:       15,     // QTY converted to the item's Base Unit — drives stock math
  BASE_RATE:      16      // PRICE converted to a per-Base-Unit rate
});

/**
 * Reverse mapping for Return columns
 */
const RETURN_COL_NAMES = Object.freeze({
  'returnNumber': RETURN_COL.RETURN_NUMBER,
  'returnDate': RETURN_COL.RETURN_DATE,
  'vendor': RETURN_COL.VENDOR,
  'contact': RETURN_COL.CONTACT,
  'billNumber': RETURN_COL.BILL_NUMBER,
  'itemName': RETURN_COL.ITEM_NAME,
  'size': RETURN_COL.SIZE,
  'narration': RETURN_COL.NARRATION,
  'qty': RETURN_COL.QTY,
  'unit': RETURN_COL.UNIT,
  'price': RETURN_COL.PRICE,
  'total': RETURN_COL.TOTAL,
  'reason': RETURN_COL.REASON,
  'remarks': RETURN_COL.REMARKS,
  'baseQty': RETURN_COL.BASE_QTY,
  'baseRate': RETURN_COL.BASE_RATE
});

// ─────────────────────────────────────────────────────────────────────────
// WASTAGE LOG SHEET COLUMN MAPPINGS (1-based indexing)
//
// Tracks component-wise material wastage / losses.
// BASE_QTY drives stock deduction the same way Returns do.
// ─────────────────────────────────────────────────────────────────────────
const WASTAGE_COL = Object.freeze({
  WASTAGE_ID: 1,    // Auto-generated unique ID (e.g. WST-20260701-001)
  DATE:       2,    // DD/MM/YYYY
  VENDOR:     3,    // Optional vendor name
  ITEM_NAME:  4,    // Component / raw material name
  SIZE:       5,    // Item size / variant
  QTY:        6,    // Quantity wasted
  UNIT:       7,    // Unit (Pcs, Kg, L, etc.)
  REASON:     8,    // Why the item was wasted (required)
  REMARKS:    9,    // Additional notes (optional)
  BASE_QTY:   10    // QTY converted to item's Base Unit — drives stock math
});

// ─────────────────────────────────────────────────────────────────────────
// ISSUED STOCK LOG SHEET COLUMN MAPPINGS (1-based indexing)
//
// Ad-hoc issuance of Stock items — e.g. extra/replacement components a
// contractor needs that aren't part of the Process's own recipe (BOM) and
// are deliberately kept OUT of any Production lot's Components Consumed
// list. BASE_QTY drives stock deduction the same way Wastage does.
// ─────────────────────────────────────────────────────────────────────────
const ISSUE_COL = Object.freeze({
  ISSUE_ID:   1,    // Auto-generated unique ID (e.g. ISS-20260701-001)
  DATE:       2,    // DD/MM/YYYY
  ISSUED_TO:  3,    // Who/what it was issued for (required — person, department, or purpose)
  REFERENCE:  4,    // Optional free-text reference (e.g. Production Lot #)
  ITEM_NAME:  5,    // Component / raw material name
  SIZE:       6,    // Item size / variant
  QTY:        7,    // Quantity issued
  UNIT:       8,    // Unit (Pcs, Kg, L, etc.)
  REMARKS:    9,    // Additional notes (optional)
  BASE_QTY:   10,   // QTY converted to item's Base Unit — drives stock math
  RATE:       11,   // Optional per-item rate — blank/0 means "not priced"
  VENDOR:     12    // Optional — set when stock is issued to a Vendor Master entry (feeds Vendor Ledger)
});

// ─────────────────────────────────────────────────────────────────────────
// ITEMS MASTER SHEET COLUMN MAPPINGS (1-based indexing)
//
// Current Structure:
// A | B | C | D | E | F
// 1 | 2 | 3 | 4 | 5 | 6
// ─────────────────────────────────────────────────────────────────────────
const ITEMS_COL = Object.freeze({
  ITEM_NAME:            1,      // Unique item identifier
  SIZE:                 2,      // Item size/variant
  REMARKS:              3,      // General remarks
  NARRATION:            4,      // Item description
  SPECIFICATION:        5,      // Technical specs
  BASE_UNIT:            6,      // Unit stock/BOM/production are tracked in (default 'Pcs')
  PURCHASE_UNIT:        7,      // Unit vendors quote rates in (e.g. 'Gross') — default = Base Unit
  WEIGHT_PER_BASE_UNIT: 8,      // Grams per Base Unit — only needed when Purchase Unit's family differs from Base Unit's (e.g. bought by Kg, used by Pcs)
  VENDOR_DATA:          9       // JSON: [{vendor, rate}, ...]
});

/**
 * Reverse mapping for Items columns
 */
const ITEMS_COL_NAMES = Object.freeze({
  'name': ITEMS_COL.ITEM_NAME,
  'size': ITEMS_COL.SIZE,
  'remarks': ITEMS_COL.REMARKS,
  'narration': ITEMS_COL.NARRATION,
  'specification': ITEMS_COL.SPECIFICATION,
  'baseUnit': ITEMS_COL.BASE_UNIT,
  'purchaseUnit': ITEMS_COL.PURCHASE_UNIT,
  'weightPerBaseUnit': ITEMS_COL.WEIGHT_PER_BASE_UNIT,
  'vendors': ITEMS_COL.VENDOR_DATA
});

// ─────────────────────────────────────────────────────────────────────────
// VENDORS SHEET COLUMN MAPPINGS (1-based indexing)
// ─────────────────────────────────────────────────────────────────────────
const VENDORS_COL = Object.freeze({
  VENDOR_NAME:    1,      // Unique vendor identifier
  CONTACT:        2,      // Contact number
  ADDRESS:        3,      // Vendor Address
  GST_NUMBER:     4,      // GST registration number
  REMARKS:        5       // Global Remarks
});

/**
 * Reverse mapping for Vendors columns
 */
const VENDORS_COL_NAMES = Object.freeze({
  'vendor': VENDORS_COL.VENDOR_NAME,
  'contact': VENDORS_COL.CONTACT,
  'address': VENDORS_COL.ADDRESS,
  'gstNumber': VENDORS_COL.GST_NUMBER,
  'remarks': VENDORS_COL.REMARKS
});

// ─────────────────────────────────────────────────────────────────────────
// RATE HISTORY SHEET COLUMN MAPPINGS (1-based indexing)
// ─────────────────────────────────────────────────────────────────────────
const RATE_HISTORY_COL = Object.freeze({
  DATE:         1,      // Date the rate was recorded (from PO or Bill)
  ITEM_NAME:    2,      // Item name
  SIZE:         3,      // Item size/variant
  NARRATION:    4,      // Item narration/description
  VENDOR_NAME:  5,      // Vendor name
  RATE:         6,      // Rate recorded (per Base Unit)
  PO_NUMBER:    7,      // Source PO number, if any
  BILL_NUMBER:  8       // Source Bill number, if any
});

/**
 * Settings for Rate History sheet structure
 */
const RATE_HISTORY_SETTINGS = Object.freeze({
  HEADER_ROWS: 1,
  DATA_START_ROW: 2
});

// ─────────────────────────────────────────────────────────────────────────
// BOM SHEET COLUMN MAPPINGS (1-based indexing)
// ─────────────────────────────────────────────────────────────────────────
const BOM_COL = Object.freeze({
  PRODUCT_ID:        1,      // Unique identifier
  PRODUCT_NAME:      2,      // Product Name (the "Model")
  ITEM_NAME:         3,      // Component Item Name
  SIZE:              4,      // Component Size
  NARRATION:         5,      // Component Narration
  RATE:              6,      // Selected Vendor Negotiated Rate
  VENDOR:            7,      // Selected Vendor Name
  QTY_PER_PRODUCT:   8,      // Qty needed per unit of product
  REMARKS:           9,      // Product-level remarks / additional details
  PROCESS_GROUP:     10,     // Reference to Process Master.Process ID — which process this component group belongs to
  COLOR:             11,     // Color variant this row applies to; blank = common to every color of this Model
  SEQUENCE:          12      // Manual display order for the Products tab (shared by every row of the same Product ID)
});

const BOM_COL_NAMES = Object.freeze({
  'productId': BOM_COL.PRODUCT_ID,
  'productName': BOM_COL.PRODUCT_NAME,
  'itemName': BOM_COL.ITEM_NAME,
  'size': BOM_COL.SIZE,
  'narration': BOM_COL.NARRATION,
  'rate': BOM_COL.RATE,
  'vendor': BOM_COL.VENDOR,
  'qtyPerProduct': BOM_COL.QTY_PER_PRODUCT,
  'remarks': BOM_COL.REMARKS,
  'processGroup': BOM_COL.PROCESS_GROUP,
  'color': BOM_COL.COLOR,
  'sequence': BOM_COL.SEQUENCE
});

// ─────────────────────────────────────────────────────────────────────────
// BOM ADDITIONAL COSTS SHEET COLUMN MAPPINGS (1-based indexing)
// Stores per-product dynamic/overhead costs (labor, paint, fitting, etc.)
// that are added on top of component material costs.
// ─────────────────────────────────────────────────────────────────────────
const BOM_COSTS_COL = Object.freeze({
  PRODUCT_ID:        1,      // Reference to BOM Product ID
  DESCRIPTION:       2,      // Cost line description (e.g. "Labor - Fitting")
  RATE:              3,      // Cost amount per unit of product
  PROCESS_NAME:      4,      // Optional: which process this cost belongs to (free string — matches Process Master name or the "Dispatch / Logistics" virtual category, same convention as Contractor Rates)
  CONTRACTOR_NAME:   5       // Optional: which contractor rate this was sourced from
});

const BOM_COSTS_COL_NAMES = Object.freeze({
  'productId': BOM_COSTS_COL.PRODUCT_ID,
  'description': BOM_COSTS_COL.DESCRIPTION,
  'rate': BOM_COSTS_COL.RATE,
  'processName': BOM_COSTS_COL.PROCESS_NAME,
  'contractorName': BOM_COSTS_COL.CONTRACTOR_NAME
});

// ─────────────────────────────────────────────────────────────────────────
// PRODUCTION SHEET COLUMN MAPPINGS (1-based indexing)
// ─────────────────────────────────────────────────────────────────────────
const PRODUCTION_COL = Object.freeze({
  DATE:              1,      // DD/MM/YYYY date
  PRODUCT_ID:        2,      // (Legacy/optional) Product ID — only set as a tag for final-stage lots
  PRODUCT_NAME:      3,      // (Legacy/optional) Product Name (de-normalized copy of the tag)
  QTY:               4,      // Lot Quantity produced
  ASSIGNED_BY:       5,      // Assigned By staff
  ASSIGNED_TO:       6,      // Assigned To worker/contractor (required)
  STATUS:            7,      // Production status
  REMARKS:           8,      // Production remarks/notes
  CUSTOM_COMPONENTS: 9,      // JSON: components consumed for this lot — [{itemName, size, sourceType, qty}, ...]
  SHEET_REMARKS:    10,      // Notes on what was changed on the production sheet, why, or for whom
  PROCESS_ID:       11,      // Reference to Process Master.Process ID (required)
  LOT_NUMBER:       12,      // Auto-generated, scoped per-process (e.g. LOT-FP-0012)
  CONSUMED_FROM_PROCESS_QTY: 13, // Qty of the predecessor process WIP output consumed by this lot
  CONTRACTOR_RATE:  14,      // Rate/unit snapshot from Contractor Rates at save time (0 if Assigned To has no rate card entry)
  CONTRACTOR_PAYABLE: 15,    // = QTY x CONTRACTOR_RATE, snapshot at save time
  OUTPUT_ITEM_NAME: 16,      // De-normalized copy of Process Master.Output Item Name at save time
  COMPONENTS_CONSUMED: 17,   // JSON: [{itemName, size, color, sourceType, qty, colorGroup}, ...] — required at log time, drives Stock/Warehouse Pool debits. colorGroup is 'COMMON' (applies to every color in the batch) or a specific Color Master name.
  COLOR: 18,                 // Display string of every color produced in this lot (comma-joined when multiple, blank if the process has no color sub-groups).
  COLOR_BREAKDOWN: 19,       // JSON: [{color, qty}, ...] — per-color quantity split for this lot. One Production row now covers an entire multi-color batch instead of one row per color.
  ORDER_NUMBER: 20           // Structured reference to Client Orders.Order Number, set only when this lot was auto-queued from a PI/Estimate line (_pushOrderLinesToProduction). Blank for manually-logged lots. The REMARKS column also carries a free-text "Auto-queued from PI ..." note for humans; this column is what deleteClientOrder/deleteClientOrdersBulk actually check before allowing a delete.
});

const PRODUCTION_COL_NAMES = Object.freeze({
  'date': PRODUCTION_COL.DATE,
  'productId': PRODUCTION_COL.PRODUCT_ID,
  'productName': PRODUCTION_COL.PRODUCT_NAME,
  'qty': PRODUCTION_COL.QTY,
  'assignedBy': PRODUCTION_COL.ASSIGNED_BY,
  'assignedTo': PRODUCTION_COL.ASSIGNED_TO,
  'status': PRODUCTION_COL.STATUS,
  'remarks': PRODUCTION_COL.REMARKS,
  'customComponents': PRODUCTION_COL.CUSTOM_COMPONENTS,
  'sheetRemarks': PRODUCTION_COL.SHEET_REMARKS,
  'processId': PRODUCTION_COL.PROCESS_ID,
  'lotNumber': PRODUCTION_COL.LOT_NUMBER,
  'consumedFromProcessQty': PRODUCTION_COL.CONSUMED_FROM_PROCESS_QTY,
  'contractorRate': PRODUCTION_COL.CONTRACTOR_RATE,
  'contractorPayable': PRODUCTION_COL.CONTRACTOR_PAYABLE,
  'outputItemName': PRODUCTION_COL.OUTPUT_ITEM_NAME,
  'componentsConsumed': PRODUCTION_COL.COMPONENTS_CONSUMED,
  'color': PRODUCTION_COL.COLOR,
  'colorBreakdown': PRODUCTION_COL.COLOR_BREAKDOWN,
  'orderNumber': PRODUCTION_COL.ORDER_NUMBER
});

// ─────────────────────────────────────────────────────────────────────────
// PROCESS MASTER SHEET COLUMN MAPPINGS (1-based indexing)
// Defines the pluggable production process types (Frame Painting, Rim
// Assembly, Frame Fitting, Packing, ...) and the Sequence that determines
// the work-in-progress (WIP) chain between them.
// ─────────────────────────────────────────────────────────────────────────
const PROCESS_COL = Object.freeze({
  PROCESS_ID:     1,      // Unique identifier (PRC-1001, ...)
  PROCESS_NAME:   2,      // e.g. "Frame Painting"
  SEQUENCE:       3,      // Integer order in the WIP chain (lower = earlier stage)
  LOT_PREFIX:     4,      // Short code used in Production lot numbers (e.g. "FP")
  IS_FINAL_STAGE: 5,      // TRUE for the stage whose output becomes finished-goods stock
  ACTIVE:         6,      // TRUE/FALSE
  REMARKS:        7,      // Free text
  OUTPUT_ITEM_NAME: 8,    // Warehouse Pool item this process produces per unit (e.g. "Painted Frame")
  PROCESS_TYPE:   9,      // Reference to Process Type Master.Name — categorizes the process (e.g. "Painting", "Welding") for grouping; optional, blank reads as "General"
  PRIMARY_COLOR_AXIS: 10, // Optional Color Axis label (see PROCESS_COMPONENTS_COL.COLOR_AXIS) whose checked "Colors to Produce" rows determine a lot's total quantity, when this process has 2+ independent color axes. Blank = legacy single-list/cross-product behavior (see computeColorGroupsForProcess).
  DISPATCH_DIFFERENTIATOR: 11 // FINAL-STAGE ONLY. Optional axis key (see computeColorAxesForProcess's `key`) naming which one of this process's color axes identifies the product on the Ready to Dispatch list. A final-stage output's pool color is a composite of every axis it combined (e.g. "Black / Blue-White / Brown / Kraft"); dispatch usually cares about ONE of those (the frame color, say), so the list is split into one row per distinct value of this axis and labelled "<Output Item Name> / <value>". Blank = one aggregate row per output, exactly as before this column existed.
});

const PROCESS_COL_NAMES = Object.freeze({
  'processId': PROCESS_COL.PROCESS_ID,
  'processName': PROCESS_COL.PROCESS_NAME,
  'sequence': PROCESS_COL.SEQUENCE,
  'lotPrefix': PROCESS_COL.LOT_PREFIX,
  'isFinalStage': PROCESS_COL.IS_FINAL_STAGE,
  'active': PROCESS_COL.ACTIVE,
  'remarks': PROCESS_COL.REMARKS,
  'outputItemName': PROCESS_COL.OUTPUT_ITEM_NAME,
  'processType': PROCESS_COL.PROCESS_TYPE,
  'primaryColorAxis': PROCESS_COL.PRIMARY_COLOR_AXIS,
  'dispatchDifferentiator': PROCESS_COL.DISPATCH_DIFFERENTIATOR
});

// ─────────────────────────────────────────────────────────────────────────
// PROCESS COMPONENTS SHEET COLUMN MAPPINGS (1-based indexing)
// A per-process checklist of which Items Master components are typically
// involved in that process (e.g. Paint + Thinner for Frame Painting),
// independent of any specific product's BOM. Quantities are product-
// specific and stay in BOM — this is just "what materials does this
// process use," not "how much."
// ─────────────────────────────────────────────────────────────────────────
const PROCESS_COMPONENTS_COL = Object.freeze({
  PROCESS_ID:    1,   // Reference to Process Master.Process ID
  ITEM_NAME:     2,   // Component item name (from Items Master, or when SOURCE_TYPE is POOL, an upstream process Output Item Name)
  SIZE:          3,   // Component size
  NARRATION:     4,   // Component narration
  QTY_PER_UNIT:  5,   // Qty needed per unit of process output (the process recipe)
  REMARKS:       6,   // Remarks
  SOURCE_TYPE:   7,   // 'ITEM' (raw material, debited from Stock) or 'POOL' (debited from Warehouse Pool). Defaults to 'ITEM'.
  COLOR_GROUP:   8,   // 'COMMON' (applies to every color variant) or a Color Master name (only offered/consumed when a Production lot is logged under that color). Defaults to 'COMMON'.
  COLOR_AXIS:    9,   // Optional label grouping this row's COLOR_GROUP value with its siblings into one independently-selectable "Colors to Produce" checkbox group (e.g. "Mudguard Color") — see computeColorGroupsForProcess. Blank = not manually assigned; falls back to auto-detected Warehouse Pool axis grouping, or the legacy flat bucket if there's no pool axis either.
  UNIT: 10            // Optional unit this row's QTY_PER_UNIT is expressed in (e.g. "Kg", "Gross"). Blank (the default for every pre-existing row) means "already in the item's Base Unit," preserving old behavior exactly — module_stock.js#_getBilledAndConsumedQtyMaps only converts via convertQtyToBaseUnit when this is non-blank and differs from the item's own Base Unit.
});

const PROCESS_COMPONENTS_COL_NAMES = Object.freeze({
  'processId': PROCESS_COMPONENTS_COL.PROCESS_ID,
  'itemName': PROCESS_COMPONENTS_COL.ITEM_NAME,
  'size': PROCESS_COMPONENTS_COL.SIZE,
  'narration': PROCESS_COMPONENTS_COL.NARRATION,
  'qtyPerUnit': PROCESS_COMPONENTS_COL.QTY_PER_UNIT,
  'remarks': PROCESS_COMPONENTS_COL.REMARKS,
  'sourceType': PROCESS_COMPONENTS_COL.SOURCE_TYPE,
  'colorGroup': PROCESS_COMPONENTS_COL.COLOR_GROUP,
  'colorAxis': PROCESS_COMPONENTS_COL.COLOR_AXIS,
  'unit': PROCESS_COMPONENTS_COL.UNIT
});

// Allowed values for PROCESS_COMPONENTS_COL.SOURCE_TYPE / a component's source
const COMPONENT_SOURCE_TYPES = Object.freeze({
  ITEM: 'ITEM',
  POOL: 'POOL'
});

// Sentinel value for PROCESS_COMPONENTS_COL.COLOR_GROUP marking a component as
// shared across every color variant of a process (as opposed to a specific
// Color Master name, which scopes the component to lots logged under that color).
const COMPONENT_COLOR_GROUP_COMMON = 'COMMON';

// Joins per-axis literal colors into one composite output color label when a
// process draws from 2+ independently-colored Warehouse Pool inputs (e.g. a
// Fitted Rim colored BCP/Black assembled with a Frame colored Blue-White/
// Orange-White/...) — see computeColorGroupsForProcess. A real producible
// unit needs one color from EVERY axis at once, so the composite ("BCP /
// Blue-White") is the actual variant identity; the individual axis tokens
// remain the literal values used to debit each input's own Warehouse Pool
// color bucket (see _componentColorMatchesBreakdown in module_production.js).
const COLOR_COMBO_DELIMITER = ' / ';

// ─────────────────────────────────────────────────────────────────────────
// PROCESS COLOR LINKS SHEET COLUMN MAPPINGS (1-based indexing)
// Manually-declared color correspondences between two AXES (e.g. Process B's
// "Orange/White" <-> Process D's "Orange"; or, since AXIS_A_KEY/AXIS_B_KEY
// were added, two axes of the very SAME process, e.g. its own tag-based
// "Rim Color" <-> "Mudguard Color"), so any recipe drawing on both treats
// them as one paired color choice instead of cross-multiplying every color
// of one against every color of the other (see computeColorGroupsForProcess/
// computeColorAxesForProcess in module_process.js). One row per paired color
// — a 6-color link is 6 rows. Storage is directional (A/B columns) but
// always read symmetrically: a process's links are every row where it
// appears in EITHER column.
// AXIS_A_KEY/AXIS_B_KEY are new columns, blank on every row saved before
// they existed — a blank axis key means "that side's own (single) Warehouse
// Pool-detected axis", the original/only meaning this sheet ever had, so
// every pre-existing row keeps resolving exactly as before
// (ensureProcessColorLinksAxisColumns lazily backfills the columns
// themselves; no value migration is needed since blank IS the compatible
// default). A non-blank axis key (e.g. "tag:mudguard color") names a
// specific manually-tagged axis, and is what makes a same-process link
// (PROCESS_A_ID === PROCESS_B_ID, two DIFFERENT axis keys) meaningful
// instead of ambiguous/self-referential.
// ─────────────────────────────────────────────────────────────────────────
const PROCESS_COLOR_LINKS_COL = Object.freeze({
  PROCESS_A_ID: 1,  // Reference to Process Master.Process ID
  COLOR_A:      2,  // One of Process A's own colors (see getProcessColorGroups)
  PROCESS_B_ID: 3,  // Reference to Process Master.Process ID (the linked process — may equal PROCESS_A_ID)
  COLOR_B:      4,  // The corresponding color on Process B
  AXIS_A_KEY:   5,  // Optional — which of Process A's axes COLOR_A belongs to; blank = its pool axis
  AXIS_B_KEY:   6   // Optional — which of Process B's axes COLOR_B belongs to; blank = its pool axis
});

const PROCESS_COLOR_LINKS_COL_NAMES = Object.freeze({
  'processAId': PROCESS_COLOR_LINKS_COL.PROCESS_A_ID,
  'colorA': PROCESS_COLOR_LINKS_COL.COLOR_A,
  'processBId': PROCESS_COLOR_LINKS_COL.PROCESS_B_ID,
  'colorB': PROCESS_COLOR_LINKS_COL.COLOR_B,
  'axisAKey': PROCESS_COLOR_LINKS_COL.AXIS_A_KEY,
  'axisBKey': PROCESS_COLOR_LINKS_COL.AXIS_B_KEY
});

const PROCESS_COLOR_LINKS_SETTINGS = Object.freeze({
  HEADER_ROWS: 1,
  DATA_START_ROW: 2,
  MAX_ROWS_PER_QUERY: 5000
});

// ─────────────────────────────────────────────────────────────────────────
// PROCESS COLOR OVERRIDES SHEET COLUMN MAPPINGS (1-based indexing)
// One row per (Process, Color) the user has manually curated away from the
// auto-computed default (see computeColorGroupsForProcess in
// module_process.js, and the "Colors to Produce" / Warehouse Pool breakdown
// dialog's Delete/Add Combination actions in module_process.js's
// excludeWarehousePoolColors/includeWarehousePoolColor). ACTION is always
// the CURRENT state for that pair — a later save for the same (Process,
// Color) overwrites the row rather than appending a second one (see
// _writeProcessColorOverrides), so this stays a small set, not a growing log.
//   EXCLUDE: this color must never appear as a known/placeholder combination
//     for this process, even though it would otherwise show up via Color
//     Master widening or prior INCLUDE. Rejected server-side (see
//     excludeWarehousePoolColors) if the color is actually configured on the
//     process's own recipe, or has real Warehouse Pool production/
//     consumption history — this table can only hide zero-data noise, never
//     real configuration or real data.
//   INCLUDE: force-adds this color as known for this process even though
//     nothing else (recipe, pool history, Color Master) would produce it —
//     e.g. a genuinely one-off combination the operator wants to pre-seed
//     opening stock for. Also how a prior EXCLUDE gets undone (INCLUDE
//     simply overwrites it for that color).
// ─────────────────────────────────────────────────────────────────────────
const PROCESS_COLOR_OVERRIDES_COL = Object.freeze({
  PROCESS_ID: 1,
  COLOR: 2,
  ACTION: 3  // 'INCLUDE' | 'EXCLUDE'
});

// ─────────────────────────────────────────────────────────────────────────
// WAREHOUSE POOL SHEET COLUMN MAPPINGS (1-based indexing)
// Ready stock of intermediate/finished process outputs (painted frames,
// fitted rims, fitted frames, packed products, ...). Credited when a
// Production lot for the originating process is marked Completed; debited
// when a downstream process consumes it (POOL-sourced component) or when
// a Product-tagged final-stage credit is Dispatched.
// ─────────────────────────────────────────────────────────────────────────
const WAREHOUSE_POOL_COL = Object.freeze({
  OUTPUT_ITEM_NAME: 1,    // Pool item name (Process Master.Output Item Name)
  PROCESS_ID:        2,   // Originating Process Master.Process ID
  PRODUCT_TAG:       3,   // Blank for intermediate WIP; Product ID when a final-stage lot was tagged
  PRODUCED_QTY:      4,   // Cumulative qty credited from Completed Production lots
  CONSUMED_QTY:      5,   // Cumulative qty debited by downstream Production/Dispatch
  AVAILABLE_QTY:     6,   // = PRODUCED_QTY - CONSUMED_QTY
  COLOR:             7    // Color Master name this bucket was credited under, blank for color-agnostic/uncolored output
});

const WAREHOUSE_POOL_COL_NAMES = Object.freeze({
  'outputItemName': WAREHOUSE_POOL_COL.OUTPUT_ITEM_NAME,
  'processId': WAREHOUSE_POOL_COL.PROCESS_ID,
  'productTag': WAREHOUSE_POOL_COL.PRODUCT_TAG,
  'producedQty': WAREHOUSE_POOL_COL.PRODUCED_QTY,
  'consumedQty': WAREHOUSE_POOL_COL.CONSUMED_QTY,
  'availableQty': WAREHOUSE_POOL_COL.AVAILABLE_QTY,
  'color': WAREHOUSE_POOL_COL.COLOR
});

const WAREHOUSE_POOL_SETTINGS = Object.freeze({
  HEADER_ROWS: 1,
  DATA_START_ROW: 2,
  MAX_ROWS_PER_QUERY: 10000
});

// ─────────────────────────────────────────────────────────────────────────
// WAREHOUSE POOL OPENING SHEET COLUMN MAPPINGS (1-based indexing)
// Manual opening-balance entries for the Warehouse Pool — lets stock that
// already existed before this app went live (or a manual correction) be
// seeded into a pool bucket. Unlike the Warehouse Pool sheet itself, this
// sheet is never wiped by recalculateWarehousePool(); it's read once per
// rebuild and folded into the matching bucket's Produced Qty alongside
// Completed Production lots, so it survives indefinitely.
// ─────────────────────────────────────────────────────────────────────────
const WAREHOUSE_POOL_OPENING_COL = Object.freeze({
  OUTPUT_ITEM_NAME: 1,    // Pool item name (Process Master.Output Item Name) — auto-filled from Process
  PROCESS_ID:        2,   // Originating Process Master.Process ID
  PRODUCT_TAG:       3,   // Optional — Product ID, only meaningful for a final-stage process
  COLOR:             4,   // Optional — Color Master name, only meaningful for a color-tracked process
  QTY:               5,   // Opening quantity
  DATE:              6,   // Date the opening balance was recorded
  REMARKS:           7
});

const WAREHOUSE_POOL_OPENING_COL_NAMES = Object.freeze({
  'outputItemName': WAREHOUSE_POOL_OPENING_COL.OUTPUT_ITEM_NAME,
  'processId': WAREHOUSE_POOL_OPENING_COL.PROCESS_ID,
  'productTag': WAREHOUSE_POOL_OPENING_COL.PRODUCT_TAG,
  'color': WAREHOUSE_POOL_OPENING_COL.COLOR,
  'qty': WAREHOUSE_POOL_OPENING_COL.QTY,
  'date': WAREHOUSE_POOL_OPENING_COL.DATE,
  'remarks': WAREHOUSE_POOL_OPENING_COL.REMARKS
});

const WAREHOUSE_POOL_OPENING_SETTINGS = Object.freeze({
  HEADER_ROWS: 1,
  DATA_START_ROW: 2,
  MAX_ROWS_PER_QUERY: 5000
});

// ─────────────────────────────────────────────────────────────────────────
// CONTRACTORS SHEET COLUMN MAPPINGS (1-based indexing)
// Master data for independent contractors who run production processes
// (and, via the virtual "Dispatch / Logistics" category, logistics too).
// ─────────────────────────────────────────────────────────────────────────
const CONTRACTORS_COL = Object.freeze({
  CONTRACTOR_NAME: 1,     // Unique identifier
  CONTACT:         2,     // Contact number
  ADDRESS:         3,     // Contractor Address
  GST_PAN:         4,     // GST/PAN number
  REMARKS:         5      // Global Remarks
});

const CONTRACTORS_COL_NAMES = Object.freeze({
  'contractorName': CONTRACTORS_COL.CONTRACTOR_NAME,
  'contact': CONTRACTORS_COL.CONTACT,
  'address': CONTRACTORS_COL.ADDRESS,
  'gstPan': CONTRACTORS_COL.GST_PAN,
  'remarks': CONTRACTORS_COL.REMARKS
});

// ─────────────────────────────────────────────────────────────────────────
// CONTRACTOR RATES SHEET COLUMN MAPPINGS (1-based indexing)
// The rate card: negotiated rate per unit for a (Contractor, Process) pair.
// Process Name is a free string matching either an active Process Master
// name or the hardcoded "Dispatch / Logistics" virtual category.
// ─────────────────────────────────────────────────────────────────────────
const CONTRACTOR_RATES_COL = Object.freeze({
  CONTRACTOR_NAME: 1,
  PROCESS_NAME:    2,
  RATE_PER_UNIT:   3,
  REMARKS:         4
});

const CONTRACTOR_RATES_COL_NAMES = Object.freeze({
  'contractorName': CONTRACTOR_RATES_COL.CONTRACTOR_NAME,
  'processName': CONTRACTOR_RATES_COL.PROCESS_NAME,
  'ratePerUnit': CONTRACTOR_RATES_COL.RATE_PER_UNIT,
  'remarks': CONTRACTOR_RATES_COL.REMARKS
});

// ─────────────────────────────────────────────────────────────────────────
// CONTRACTOR PAYMENTS SHEET COLUMN MAPPINGS (1-based indexing)
// Datewise record of actual cash payments made to a contractor, used
// together with Production's accrued Payable amounts to compute each
// contractor's running account balance.
// ─────────────────────────────────────────────────────────────────────────
const CONTRACTOR_PAYMENTS_COL = Object.freeze({
  DATE:            1,
  CONTRACTOR_NAME: 2,
  AMOUNT:          3,
  MODE_REFERENCE:  4,   // e.g. "Cash", "UPI - txn123", cheque no.
  REMARKS:         5
});

const CONTRACTOR_PAYMENTS_COL_NAMES = Object.freeze({
  'date': CONTRACTOR_PAYMENTS_COL.DATE,
  'contractorName': CONTRACTOR_PAYMENTS_COL.CONTRACTOR_NAME,
  'amount': CONTRACTOR_PAYMENTS_COL.AMOUNT,
  'modeReference': CONTRACTOR_PAYMENTS_COL.MODE_REFERENCE,
  'remarks': CONTRACTOR_PAYMENTS_COL.REMARKS
});

// Hardcoded virtual "process" category for logistics rate-card entries.
// Not a real Process Master row — never appears in the Production process
// dropdown or the WIP chain, only as an option inside the Contractor Rate
// Card and Dispatch logistics contractor lookup.
const LOGISTICS_PROCESS_NAME = 'Dispatch / Logistics';

// ─────────────────────────────────────────────────────────────────────────
// UNIT MASTER SHEET COLUMN MAPPINGS (1-based indexing)
// Fixed-ratio unit conversions, grouped by FAMILY (e.g. units that all
// convert to each other by a constant factor: Pcs/Dozen/Gross are all
// 'Count'; Gram/Kg are all 'Weight'). Crossing families (e.g. Kg -> Pcs)
// is NOT expressed here — that needs an item-specific weight, see
// ITEMS_COL.WEIGHT_PER_BASE_UNIT.
// ─────────────────────────────────────────────────────────────────────────
const UNITS_COL = Object.freeze({
  UNIT_NAME:        1,      // Unique identifier (e.g. 'Gross')
  FAMILY:           2,      // Conversion family (e.g. 'Count', 'Weight', 'Volume')
  FACTOR_TO_BASE:   3,      // How many of the family's base unit (factor 1) this unit equals
  REMARKS:          4       // Free text
});

// ─────────────────────────────────────────────────────────────────────────
// MASTER DATA CACHE KEYS
// Used with utils.js's getCachedListResponse()/invalidateListCache() to
// wrap the rarely-changing master-data "get list" endpoints in a short
// CacheService TTL. getProcessData(activeOnly) has two distinct result
// sets depending on its argument, so it gets two keys.
// ─────────────────────────────────────────────────────────────────────────
const MASTER_DATA_CACHE_KEYS = Object.freeze({
  VENDORS:             'mdcache_vendors',
  UNITS:               'mdcache_units',
  CONTRACTORS:         'mdcache_contractors',
  PROCESS_ALL:         'mdcache_process_all',
  PROCESS_ACTIVE:      'mdcache_process_active',
  COLOR_MASTER:        'mdcache_tags_color',
  MODEL_MASTER:        'mdcache_tags_model',
  PROCESS_TYPE_MASTER: 'mdcache_tags_processtype',
  ITEMS:               'mdcache_items'
});

const UNITS_COL_NAMES = Object.freeze({
  'unitName': UNITS_COL.UNIT_NAME,
  'family': UNITS_COL.FAMILY,
  'factorToBase': UNITS_COL.FACTOR_TO_BASE,
  'remarks': UNITS_COL.REMARKS
});

// ─────────────────────────────────────────────────────────────────────────
// TAG MASTER COLUMN MAPPING (1-based indexing)
// Shared layout for the simple "name list" masters — Color Master and Model
// Master — both just a unique Name + optional Remarks. See module_tags.js.
// ─────────────────────────────────────────────────────────────────────────
const TAG_COL = Object.freeze({
  NAME:     1,
  REMARKS:  2
});

// ─────────────────────────────────────────────────────────────────────────
// STOCK SHEET COLUMN MAPPINGS (1-based indexing)
// ─────────────────────────────────────────────────────────────────────────
const STOCK_COL = Object.freeze({
  ITEM_NAME:     1,
  SIZE:          2,
  INITIAL_STOCK: 3,
  CURRENT_STOCK: 4,
  THRESHOLD:     5,
  DEAD_STOCK:    6
});

const STOCK_COL_NAMES = Object.freeze({
  'itemName':     STOCK_COL.ITEM_NAME,
  'size':          STOCK_COL.SIZE,
  'initialStock': STOCK_COL.INITIAL_STOCK,
  'currentStock': STOCK_COL.CURRENT_STOCK,
  'threshold':     STOCK_COL.THRESHOLD,
  'deadStock':    STOCK_COL.DEAD_STOCK
});

// ─────────────────────────────────────────────────────────────────────────
// CLIENTS SHEET COLUMN MAPPINGS (1-based indexing)
// ─────────────────────────────────────────────────────────────────────────
const CLIENTS_COL = Object.freeze({
  CLIENT_NAME:    1,      // Unique client identifier
  CONTACT:        2,      // Contact number
  ADDRESS:        3,      // Client Address
  GST_NUMBER:     4,      // GST registration number
  REMARKS:        5       // Global Remarks
});

const CLIENTS_COL_NAMES = Object.freeze({
  'clientName': CLIENTS_COL.CLIENT_NAME,
  'contact': CLIENTS_COL.CONTACT,
  'address': CLIENTS_COL.ADDRESS,
  'gstNumber': CLIENTS_COL.GST_NUMBER,
  'remarks': CLIENTS_COL.REMARKS
});

// ─────────────────────────────────────────────────────────────────────────
// CLIENT ORDERS (PI / ESTIMATES) SHEET COLUMN MAPPINGS (1-based indexing)
// Line-item ledger grouped by Order Number, similar to BOM components
// grouped by Product ID.
// ─────────────────────────────────────────────────────────────────────────
const CLIENT_ORDERS_COL = Object.freeze({
  ORDER_NUMBER:      1,      // Unique PI/Estimate identifier (ORD-1001, ...)
  ORDER_DATE:        2,      // DD/MM/YYYY
  CLIENT_NAME:       3,      // Reference to Clients sheet
  PRODUCT_ID:        4,      // Reference to BOM.Product ID
  PRODUCT_NAME:      5,      // Product Name (de-normalized copy)
  QTY_ORDERED:       6,      // Quantity ordered for this line
  STATUS:            7,      // Estimate | Order Confirmed | Cancelled
  ORDER_REMARKS:     8,      // Order-level remarks (repeated per line)
  LINE_REMARKS:      9,      // Per-product line remarks
  PRODUCTION_PUSHED: 10      // 'Yes' once auto-queued into Production with a real Process; 'Manual' if it couldn't be auto-resolved (retried on every save) and needs to be logged by hand in the Production tab; blank if not yet attempted
});

const CLIENT_ORDERS_COL_NAMES = Object.freeze({
  'orderNumber': CLIENT_ORDERS_COL.ORDER_NUMBER,
  'orderDate': CLIENT_ORDERS_COL.ORDER_DATE,
  'clientName': CLIENT_ORDERS_COL.CLIENT_NAME,
  'productId': CLIENT_ORDERS_COL.PRODUCT_ID,
  'productName': CLIENT_ORDERS_COL.PRODUCT_NAME,
  'qty': CLIENT_ORDERS_COL.QTY_ORDERED,
  'status': CLIENT_ORDERS_COL.STATUS,
  'orderRemarks': CLIENT_ORDERS_COL.ORDER_REMARKS,
  'lineRemarks': CLIENT_ORDERS_COL.LINE_REMARKS,
  'productionPushed': CLIENT_ORDERS_COL.PRODUCTION_PUSHED
});

// ─────────────────────────────────────────────────────────────────────────
// DISPATCH SHEET COLUMN MAPPINGS (1-based indexing)
// ─────────────────────────────────────────────────────────────────────────
const DISPATCH_COL = Object.freeze({
  DISPATCH_NUMBER: 1,      // Unique identifier (DSP-1001, ...)
  DISPATCH_DATE:   2,      // DD/MM/YYYY
  ORDER_NUMBER:    3,      // Optional reference to Client Orders sheet
  CLIENT_NAME:     4,      // Reference to Clients sheet
  PRODUCT_ID:      5,      // Reference to BOM.Product ID
  PRODUCT_NAME:    6,      // Product Name (de-normalized copy)
  QTY:             7,      // Quantity dispatched
  TRANSPORT:       8,      // Transport/vehicle details
  REMARKS:         9,      // Dispatch remarks
  INVOICE_NUMBER:  10,     // Optional bill/invoice number (can be added later)
  PRIVATE_MARK:    11,     // Optional private mark
  GR_NUMBER:       12,     // Optional GR (Goods Receipt) number
  LOGISTICS_CONTRACTOR: 13, // Optional logistics contractor name (Select2 tag, same pattern as Production Assigned To)
  LOGISTICS_RATE:  14,     // Rate/unit snapshot from Contractor Rates (Dispatch / Logistics category) at save time
  LOGISTICS_COST:  15,     // = QTY x LOGISTICS_RATE, snapshot at save time
  RATE:            16,     // Optional per-unit billing/sale rate for this line item (record-keeping only)
  AMOUNT:          17      // = QTY x RATE, 0 when RATE is blank
});

const DISPATCH_COL_NAMES = Object.freeze({
  'dispatchNumber': DISPATCH_COL.DISPATCH_NUMBER,
  'dispatchDate': DISPATCH_COL.DISPATCH_DATE,
  'orderNumber': DISPATCH_COL.ORDER_NUMBER,
  'clientName': DISPATCH_COL.CLIENT_NAME,
  'productId': DISPATCH_COL.PRODUCT_ID,
  'productName': DISPATCH_COL.PRODUCT_NAME,
  'qty': DISPATCH_COL.QTY,
  'transport': DISPATCH_COL.TRANSPORT,
  'remarks': DISPATCH_COL.REMARKS,
  'invoiceNumber': DISPATCH_COL.INVOICE_NUMBER,
  'privateMark': DISPATCH_COL.PRIVATE_MARK,
  'grNumber': DISPATCH_COL.GR_NUMBER,
  'logisticsContractor': DISPATCH_COL.LOGISTICS_CONTRACTOR,
  'logisticsRate': DISPATCH_COL.LOGISTICS_RATE,
  'logisticsCost': DISPATCH_COL.LOGISTICS_COST,
  'rate': DISPATCH_COL.RATE,
  'amount': DISPATCH_COL.AMOUNT
});

// ─────────────────────────────────────────────────────────────────────────
// LOGS SHEET COLUMN MAPPINGS (1-based indexing)
// For audit trail and error tracking
// ─────────────────────────────────────────────────────────────────────────
const LOGS_COL = Object.freeze({
  TIMESTAMP:      1,      // ISO 8601 timestamp
  USER:           2,      // User email
  ACTION:         3,      // Action type (CREATE, UPDATE, DELETE)
  SHEET:          4,      // Sheet affected
  RECORD_ID:      5,      // Primary key of affected record
  DETAILS:        6,      // JSON details of the action
  STATUS:         7       // SUCCESS, ERROR, WARNING
});

// ─────────────────────────────────────────────────────────────────────────
// DATE & TIME FORMATS
// ─────────────────────────────────────────────────────────────────────────
const DATE_FORMATS = Object.freeze({
  DISPLAY: 'DD/MM/YYYY',      // For UI display
  ISO: 'YYYY-MM-DD',          // ISO 8601
  TIMESTAMP: 'YYYY-MM-DD HH:mm:ss'  // For logs
});

// ─────────────────────────────────────────────────────────────────────────
// ERROR CODES & MESSAGES
// ─────────────────────────────────────────────────────────────────────────
const ERROR_CODES = Object.freeze({
  INVALID_PO_NUMBER: 'INVALID_PO_NUMBER',
  SHEET_NOT_FOUND: 'SHEET_NOT_FOUND',
  INVALID_DATA: 'INVALID_DATA',
  DUPLICATE_ENTRY: 'DUPLICATE_ENTRY',
  MISSING_REQUIRED_FIELD: 'MISSING_REQUIRED_FIELD',
  UNAUTHORIZED: 'UNAUTHORIZED',
  RATE_LIMIT_EXCEEDED: 'RATE_LIMIT_EXCEEDED',
  UNKNOWN_ERROR: 'UNKNOWN_ERROR'
});

const ERROR_MESSAGES = Object.freeze({
  INVALID_PO_NUMBER: 'PO Number must be between 1000 and 999999',
  SHEET_NOT_FOUND: 'Required sheet not found in spreadsheet',
  INVALID_DATA: 'Data format or values are invalid',
  DUPLICATE_ENTRY: 'This record already exists',
  MISSING_REQUIRED_FIELD: 'Required field is missing',
  UNAUTHORIZED: 'You do not have permission to perform this action',
  RATE_LIMIT_EXCEEDED: 'Too many requests. Please try again later',
  UNKNOWN_ERROR: 'An unknown error occurred. Please contact support'
});

// ─────────────────────────────────────────────────────────────────────────
// HELPER FUNCTIONS
// ─────────────────────────────────────────────────────────────────────────

/**
 * Get the correct sheet name from config
 * @param {string} sheetKey - Key from APP_CONFIG.SHEETS
 * @returns {string} Sheet name
 * @throws {Error} If sheet key not found
 */
function getSheetName(sheetKey) {
  if (!APP_CONFIG.SHEETS.hasOwnProperty(sheetKey)) {
    throw new Error(`Sheet key "${sheetKey}" not found in configuration`);
  }
  return APP_CONFIG.SHEETS[sheetKey];
}

/**
 * Get column index by field name
 * @param {string} sheetKey - 'PO', 'BILL', 'ITEMS', etc.
 * @param {string} fieldName - Field name (e.g., 'vendor', 'poNumber')
 * @returns {number} Column index (1-based)
 */
function getColumnIndex(sheetKey, fieldName) {
  let colNames;
  
  switch (sheetKey.toUpperCase()) {
    case 'PO':
      colNames = PO_COL_NAMES;
      break;
    case 'BILL':
      colNames = BILL_COL_NAMES;
      break;
    case 'RETURN':
      colNames = RETURN_COL_NAMES;
      break;
    case 'ITEMS':
      colNames = ITEMS_COL_NAMES;
      break;
    case 'VENDORS':
      colNames = VENDORS_COL_NAMES;
      break;
    case 'BOM':
      colNames = BOM_COL_NAMES;
      break;
    case 'BOM_COSTS':
      colNames = BOM_COSTS_COL_NAMES;
      break;
    case 'PRODUCTION':
      colNames = PRODUCTION_COL_NAMES;
      break;
    case 'STOCK':
      colNames = STOCK_COL_NAMES;
      break;
    case 'CLIENTS':
      colNames = CLIENTS_COL_NAMES;
      break;
    case 'CLIENT_ORDERS':
      colNames = CLIENT_ORDERS_COL_NAMES;
      break;
    case 'DISPATCH':
      colNames = DISPATCH_COL_NAMES;
      break;
    case 'PROCESS_MASTER':
      colNames = PROCESS_COL_NAMES;
      break;
    case 'PROCESS_COMPONENTS':
      colNames = PROCESS_COMPONENTS_COL_NAMES;
      break;
    case 'PROCESS_COLOR_LINKS':
      colNames = PROCESS_COLOR_LINKS_COL_NAMES;
      break;
    case 'WAREHOUSE_POOL_OPENING':
      colNames = WAREHOUSE_POOL_OPENING_COL_NAMES;
      break;
    case 'WAREHOUSE_POOL':
      colNames = WAREHOUSE_POOL_COL_NAMES;
      break;
    case 'UNITS':
      colNames = UNITS_COL_NAMES;
      break;
    case 'CONTRACTORS':
      colNames = CONTRACTORS_COL_NAMES;
      break;
    case 'CONTRACTOR_RATES':
      colNames = CONTRACTOR_RATES_COL_NAMES;
      break;
    case 'CONTRACTOR_PAYMENTS':
      colNames = CONTRACTOR_PAYMENTS_COL_NAMES;
      break;
    case 'LOGS':
      colNames = LOGS_COL;
      break;
    default:
      throw new Error(`Unknown sheet key: ${sheetKey}`);
  }
  
  if (!colNames.hasOwnProperty(fieldName)) {
    throw new Error(`Field "${fieldName}" not found for sheet "${sheetKey}"`);
  }
  
  return colNames[fieldName];
}

/**
 * Get all sheet information for validation
 * @returns {Object} Configuration object
 */
function getConfig() {
  return {
    sheetNames: APP_CONFIG.SHEETS,
    poColumns: PO_COL,
    billColumns: BILL_COL,
    itemsColumns: ITEMS_COL,
    vendorsColumns: VENDORS_COL,
    validation: APP_CONFIG.VALIDATION,
    dateFormats: DATE_FORMATS
  };
}

/**
 * Log an action to the Logs sheet
 * @param {string} action - Action type (CREATE, UPDATE, DELETE, READ, ERROR)
 * @param {string} sheet - Sheet name
 * @param {string} recordId - Primary key or identifier
 * @param {string} details - JSON details
 * @param {string} status - SUCCESS, ERROR, WARNING
 */
// Cached within a single GAS execution — getEmail() is slow and never changes mid-call
let _cachedUserEmail = null;

function _getActiveUserEmail() {
  if (_cachedUserEmail) return _cachedUserEmail;
  try {
    _cachedUserEmail = Session.getActiveUser().getEmail() || 'unknown';
  } catch (e) {
    _cachedUserEmail = 'unknown';
  }
  return _cachedUserEmail;
}

function logAction(action, sheet, recordId, details = '', status = 'SUCCESS') {
  try {
    const logsSheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(getSheetName('LOGS'));

    if (!logsSheet) {
      Logger.log('Logs sheet not found. Skipping log entry.');
      return;
    }

    logsSheet.appendRow([
      new Date().toISOString(),
      _getActiveUserEmail(),
      action,
      sheet,
      recordId,
      details,
      status
    ]);
  } catch (error) {
    Logger.log(`Error logging action: ${error.message}`);
  }
}

/**
 * getRecentNotificationLogs()
 *
 * Read-only counterpart to logAction() — feeds the web app's notification
 * bell (App.Notify.checkBackendAlerts, Script_Core.html) so script errors
 * and Internal Ledger Audit findings (module_audit.js) that never produced
 * a toast in front of the user (background triggers, another tab/user, a
 * resumed session) still surface somewhere. Scans only the most recent rows
 * of the Logs sheet (cheap, bounded) for ERROR/WARNING entries from the
 * last 7 days, newest first, capped to 30.
 *
 * @returns {Object} API response; data = Array<{ key, timestamp, action,
 *   source, recordId, details, status }> — `key` is a stable per-row
 *   dedupe key the client persists so the same entry isn't re-notified on
 *   every reload.
 */
function getRecentNotificationLogs() {
  try {
    const logsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getSheetName('LOGS'));
    if (!logsSheet) return buildResponse(true, [], 'No logs sheet found.');

    const lastRow = logsSheet.getLastRow();
    if (lastRow < 2) return buildResponse(true, [], 'No log entries yet.');

    const MAX_ROWS_SCANNED = 500;
    const MAX_RESULTS = 30;
    const startRow = Math.max(2, lastRow - MAX_ROWS_SCANNED + 1);
    const numRows = lastRow - startRow + 1;
    const data = logsSheet.getRange(startRow, 1, numRows, 7).getValues();
    const cutoff = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);

    const results = [];
    for (let i = data.length - 1; i >= 0 && results.length < MAX_RESULTS; i--) {
      const row = data[i];
      const status = String(row[LOGS_COL.STATUS - 1] || '');
      if (status !== 'ERROR' && status !== 'WARNING') continue;

      const tsRaw = row[LOGS_COL.TIMESTAMP - 1];
      const ts = tsRaw instanceof Date ? tsRaw : new Date(tsRaw);
      if (isNaN(ts.getTime()) || ts < cutoff) continue;

      const action = String(row[LOGS_COL.ACTION - 1] || '');
      const recordId = String(row[LOGS_COL.RECORD_ID - 1] || '');
      results.push({
        key: `${ts.toISOString()}|${action}|${recordId}`,
        timestamp: ts.toISOString(),
        action,
        source: String(row[LOGS_COL.SHEET - 1] || ''),
        recordId,
        details: String(row[LOGS_COL.DETAILS - 1] || ''),
        status
      });
    }

    return buildResponse(true, results, `${results.length} recent issue(s).`);
  } catch (error) {
    return handleError('getRecentNotificationLogs', error);
  }
}
