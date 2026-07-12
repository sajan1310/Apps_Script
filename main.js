/**
 * ═══════════════════════════════════════════════════════════════════════════
 * main.gs — MAIN ENTRY POINT & ROUTING
 * 
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * - Primary entry point for Google Apps Script web app (doGet)
 * - Component/module loading with error boundary isolation
 * - Security headers and viewport configuration
 * - Graceful error handling with user-friendly fallback UI
 * - Deep linking and route parameter support
 * 
 * Architecture:
 * ───────────────────────────────────────────────────────────────────────────
 * doGet(e)
 *   ├─ Template creation from Index.html
 *   ├─ Route parameter passing
 *   ├─ Security headers (XFrame, viewport)
 *   ├─ Error handling with fallback UI
 *   └─ Production-safe error reporting
 * 
 * include(filename)
 *   ├─ Modular component loading
 *   ├─ Error boundary isolation
 *   ├─ Graceful degradation
 *   └─ Visible error messages in DOM
 * 
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ─────────────────────────────────────────────────────────────────────────
// CONFIGURATION CONSTANTS
// ─────────────────────────────────────────────────────────────────────────

/**
 * Default application title (fallback if APP_CONFIG unavailable)
 */
const DEFAULT_APP_TITLE = 'Maharaja Bikes ERP';

/**
 * Production environment flag
 * Set to false for development/debugging
 */
const PRODUCTION_MODE = true;

// ─────────────────────────────────────────────────────────────────────────
// MAIN ENTRY POINT
// ─────────────────────────────────────────────────────────────────────────

/**
 * doGet(e)
 * 
 * Main entry point for Google Apps Script web app.
 * Handles:
 * - Template loading from Index.html
 * - Route parameter parsing (for future deep linking)
 * - Security headers configuration
 * - Graceful error handling with fallback UI
 * 
 * @param {Object} e - Request parameter object from Google Apps Script
 *   - e.parameter: URL parameters (querystring)
 *   - e.pathInfo: URL path information
 *   - e.contentLength: Request body size
 * 
 * @returns {HtmlOutput} Rendered HTML output to serve to client
 * 
 * @example
 * // Normal access
 * // https://script.google.com/macros/d/{SCRIPT_ID}/usercodeappserver
 * 
 * // With route parameters (for future deep linking)
 * // https://script.google.com/macros/d/{SCRIPT_ID}/usercodeappserver?tab=billLedger&poId=1001
 */
function doGet(e) {
  try {
    // ─────────────────────────────────────────────────────────────────
    // STEP 1: Determine Application Title
    // ─────────────────────────────────────────────────────────────────
    // Graceful fallback in case APP_CONFIG fails to initialize
    let appTitle = DEFAULT_APP_TITLE;

    try {
      if (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.APP_TITLE) {
        appTitle = APP_CONFIG.APP_TITLE;
      }
    } catch (configError) {
      Log.warn('[doGet] APP_CONFIG not available, using default title:', configError);
    }

    // ─────────────────────────────────────────────────────────────────
    // STEP 2: Route by ?ui= param — the sandboxed iframe can't read the
    // User-Agent, so mobile vs. desktop is an explicit URL switch rather
    // than a UA sniff. ?ui=mobile serves the separate app-like shell
    // (Mobile_Index.html); anything else serves the existing desktop Index.
    // ─────────────────────────────────────────────────────────────────
    const params = e ? (e.parameter || {}) : {};
    const isMobileUi = String(params.ui || '').toLowerCase() === 'mobile';
    const templateName = isMobileUi ? 'Mobile_Index' : 'Index';

    // ─────────────────────────────────────────────────────────────────
    // STEP 3: Load and Configure HTML Template
    // ─────────────────────────────────────────────────────────────────
    const template = HtmlService.createTemplateFromFile(templateName);

    // Pass URL parameters into template for future routing/deep linking support
    // Examples: ?tab=billLedger, ?poId=1001, etc.
    template.routeData = params;
    template.isProduction = PRODUCTION_MODE;

    // ─────────────────────────────────────────────────────────────────
    // STEP 4: Render Template and Configure Output
    // ─────────────────────────────────────────────────────────────────
    const output = template.evaluate();

    // Set application title
    output.setTitle(isMobileUi ? appTitle + ' — Mobile' : appTitle);

    // Configure viewport for mobile/responsive design. Locking maximum-scale
    // / user-scalable=no blocks pinch-zoom entirely (WCAG 1.4.4) on a
    // data-dense ERP — removed from both shells. iOS auto-zoom-on-focus is
    // instead prevented the correct way: keeping form field font-size >=16px
    // at mobile widths (see Styles.html's <=768px block, Mobile_Styles.html).
    output.addMetaTag('viewport', 'width=device-width, initial-scale=1');

    // Security: Prevent clickjacking by controlling iframe embedding
    // XFrameOptionsMode.DEFAULT allows same-origin framing
    output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);

    // Security headers (X-Content-Type-Options and referrer) are handled automatically by Google Apps Script

    Logger.log(`[doGet] Successfully rendered ${appTitle}${isMobileUi ? ' (mobile)' : ''}`);

    return output;
  } catch (error) {
    Log.error('[doGet] Critical routing error:', error);
    return _renderErrorFallback(error);
  }
}

// ─────────────────────────────────────────────────────────────────────────
// COMPONENT LOADER
// ─────────────────────────────────────────────────────────────────────────

/**
 * include(filename)
 * 
 * Securely loads and injects modular HTML/CSS/JS components.
 * 
 * Purpose:
 * - Modular architecture (separate files for views, styles, scripts)
 * - DRY principle (reuse components across templates)
 * - Error boundary isolation (component failure doesn't crash entire app)
 * - Graceful degradation (shows error message, not stack trace)
 * 
 * Usage in templates:
 * <?!= include('Styles'); ?>
 * <?!= include('View_Print'); ?>
 * <?!= include('Script'); ?>
 * 
 * Why not just inline everything?
 * - Easier to maintain and test individual components
 * - Google Apps Script has file size limits
 * - Separation of concerns (views, styles, logic)
 * - Reusable across multiple templates
 * 
 * @param {string} filename - Name of file to include (without .html extension)
 *   - All files must be in the Apps Script project
 *   - Examples: 'Styles', 'View_Print', 'Script', 'Modals'
 * 
 * @returns {string} HTML content of the file, or error message if load fails
 * 
 * @example
 * // In Index.html template:
 * <?!= include('View_Print'); ?>
 * 
 * // Returns content of View_Print.html or error div on failure
 */
function include(filename) {
  if (!filename || typeof filename !== 'string') {
    Log.error('[include] Invalid filename provided');
    return _renderComponentError('Invalid filename');
  }

  try {
    // Attempt to load the raw file content first using createHtmlOutputFromFile.
    // This allows us to bypass the server-side template engine for static components
    // and script files that do not contain Google Apps Script scriptlets (<? ... ?>).
    // This prevents parsing errors caused by literal less-than signs ("<") used as
    // comparison operators in JavaScript modules.
    const rawContent = HtmlService.createHtmlOutputFromFile(filename).getContent();

    // Check if the component contains any dynamic Apps Script scriptlets.
    // If it does (like Index.html or Script.html), evaluate it as a template.
    // If it doesn't (like Styles.html or Modals.html), serve the raw content.
    if (rawContent.indexOf('<?') !== -1) {
      const fileContent = HtmlService.createTemplate(rawContent).evaluate().getContent();
      Logger.log(`[include] Successfully evaluated component with scriptlets: ${filename}`);
      return fileContent;
    } else {
      Logger.log(`[include] Successfully loaded static/script component: ${filename}`);
      return rawContent;
    }
  } catch (error) {
    Log.error(`[include] Component load failure [${filename}]:`, error);

    // If it's a script component, return a JS-safe console error to avoid crashing client-side JS parser
    if (filename.toLowerCase().indexOf('script') !== -1) {
      return `\nLog.error("[include] Failed to load script component '${filename}': ${error.message.replace(/"/g, '\\"')}");\n`;
    }

    // Return visible error message in DOM (error boundary)
    // This allows the rest of the application to render even if one module is missing
    // Better than crashing the entire template
    return _renderComponentError(filename, error.message);
  }
}

// ─────────────────────────────────────────────────────────────────────────
// ERROR HANDLING & FALLBACK RENDERING
// ─────────────────────────────────────────────────────────────────────────

/**
 * _renderErrorFallback(error)
 * 
 * Renders a user-friendly error page for critical failures.
 * 
 * Called when:
 * - Template loading fails
 * - doGet() throws an exception
 * - Configuration initialization fails
 * 
 * Design:
 * - Professional appearance
 * - Helpful error message (but not sensitive details)
 * - Icon and clear visual hierarchy
 * - Responsive on all screen sizes
 * 
 * @param {Error} error - Error object
 * @returns {HtmlOutput} Error fallback HTML output
 * 
 * @private
 */
function _renderErrorFallback(error) {
  const errorMessage = error instanceof Error ? error.message : String(error);
  
  const errorHtml = `
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <!-- Meta tags omitted due to HtmlService restrictions -->
      <title>Application Error</title>
      <style>
        * {
          margin: 0;
          padding: 0;
          box-sizing: border-box;
        }
        
        body {
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
          background: linear-gradient(135deg, #f9fafb 0%, #f3f4f6 100%);
          display: flex;
          align-items: center;
          justify-content: center;
          min-height: 100vh;
          padding: 20px;
        }
        
        .error-container {
          background: white;
          border-radius: 12px;
          box-shadow: 0 20px 60px rgba(0, 0, 0, 0.1);
          padding: 60px 40px;
          max-width: 500px;
          text-align: center;
          animation: slideUp 400ms ease-out;
        }
        
        @keyframes slideUp {
          from {
            opacity: 0;
            transform: translateY(20px);
          }
          to {
            opacity: 1;
            transform: translateY(0);
          }
        }
        
        .error-icon {
          width: 80px;
          height: 80px;
          background: #fee2e2;
          border-radius: 50%;
          display: flex;
          align-items: center;
          justify-content: center;
          margin: 0 auto 24px;
          flex-shrink: 0;
        }
        
        .error-icon svg {
          width: 48px;
          height: 48px;
          color: #dc2626;
        }
        
        h1 {
          font-size: 28px;
          font-weight: 700;
          color: #111827;
          margin-bottom: 12px;
        }
        
        .error-message {
          font-size: 16px;
          color: #6b7280;
          margin-bottom: 24px;
          line-height: 1.6;
        }
        
        .error-code {
          background: #fef2f2;
          border-left: 4px solid #dc2626;
          padding: 16px;
          border-radius: 6px;
          margin-bottom: 24px;
          text-align: left;
        }
        
        .error-code code {
          font-family: 'Courier New', Courier, monospace;
          font-size: 13px;
          color: #991b1b;
          word-break: break-all;
          display: block;
          line-height: 1.5;
        }
        
        .error-help {
          font-size: 14px;
          color: #9ca3af;
          margin-bottom: 24px;
        }
        
        .action-buttons {
          display: flex;
          gap: 12px;
          justify-content: center;
          flex-wrap: wrap;
        }
        
        button {
          padding: 12px 24px;
          border: none;
          border-radius: 6px;
          font-size: 14px;
          font-weight: 600;
          cursor: pointer;
          transition: all 200ms ease;
        }
        
        .btn-reload {
          background: #2563eb;
          color: white;
        }
        
        .btn-reload:hover {
          background: #1d4ed8;
          box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3);
        }
        
        .btn-support {
          background: transparent;
          color: #2563eb;
          border: 2px solid #2563eb;
        }
        
        .btn-support:hover {
          background: #eff6ff;
        }
        
        @media (max-width: 480px) {
          .error-container {
            padding: 40px 24px;
          }
          
          h1 {
            font-size: 24px;
          }
          
          .error-message {
            font-size: 14px;
          }
        }
      </style>
    </head>
    <body>
      <div class="error-container">
        <div class="error-icon">
          <svg fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"></path>
          </svg>
        </div>
        
        <h1>System Unavailable</h1>
        
        <p class="error-message">
          We encountered a problem loading the application interface. Our team has been notified.
        </p>
        
        <div class="error-code">
          <code>${errorMessage}</code>
        </div>
        
        <p class="error-help">
          Try reloading the page. If the problem persists, please contact support.
        </p>
        
        <div class="action-buttons">
          <button class="btn-reload" onclick="window.location.reload();">
            Reload Page
          </button>
          <button class="btn-support" onclick="window.open('mailto:support@maharajabikes.com');">
            Contact Support
          </button>
        </div>
      </div>
    </body>
    </html>
  `;

  return HtmlService.createHtmlOutput(errorHtml)
    .setTitle('Application Error')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * _renderComponentError(filename, errorMessage)
 * 
 * Renders a visible error message in the DOM when a component fails to load.
 * Allows the rest of the application to continue rendering.
 * 
 * @param {string} filename - Name of component that failed to load
 * @param {string} errorMessage - Error message from the exception
 * @returns {string} HTML error boundary div
 * 
 * @private
 */
function _renderComponentError(filename, errorMessage = '') {
  return `
    <!-- ERROR BOUNDARY: FAILED TO LOAD COMPONENT '${filename}' -->
    <div style="background-color: #fef2f2;
                border-left: 4px solid #dc2626;
                color: #991b1b;
                padding: 16px;
                margin: 16px 0;
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                font-size: 13px;
                border-radius: 4px;
                box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);">
      <strong style="display: block; margin-bottom: 8px; font-size: 14px;">
        ⚠️ Component Load Error
      </strong>
      <div style="margin-bottom: 8px;">
        Failed to inject module: <code style="background: rgba(156, 27, 27, 0.1); padding: 2px 6px; border-radius: 2px;">${filename}.html</code>
      </div>
      <div style="font-size: 12px; opacity: 0.8;">
        ${errorMessage || 'Check Apps Script files to ensure this module exists.'}
      </div>
    </div>
  `;
}

// ─────────────────────────────────────────────────────────────────────────
// HEALTH CHECK & DIAGNOSTICS
// ─────────────────────────────────────────────────────────────────────────

/**
 * getSystemStatus()
 * 
 * Returns system status and diagnostic information.
 * Useful for debugging and monitoring.
 * 
 * @returns {Object} System status object
 */
function getSystemStatus() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    const sheets = [];
    spreadsheet.getSheets().forEach(sheet => {
      sheets.push({
        name: sheet.getName(),
        rows: sheet.getLastRow(),
        columns: sheet.getLastColumn()
      });
    });

    return {
      success: true,
      data: {
        appTitle: DEFAULT_APP_TITLE,
        environment: PRODUCTION_MODE ? 'production' : 'development',
        spreadsheetId: spreadsheet.getId(),
        spreadsheetName: spreadsheet.getName(),
        sheets: sheets,
        timestamp: new Date().toISOString()
      },
      message: 'System operational'
    };
  } catch (error) {
    return buildErrorResponse(ERROR_CODES.UNKNOWN_ERROR, error.message);
  }
}

/**
 * testConnection()
 * 
 * Simple test to verify backend is responding correctly.
 * Call from frontend to verify GAS connection is working.
 * 
 * @returns {Object} Test result
 */
function testConnection() {
  try {
    return {
      success: true,
      data: {
        timestamp: new Date().toISOString(),
        appTitle: DEFAULT_APP_TITLE,
        configAvailable: typeof APP_CONFIG !== 'undefined'
      },
      message: 'Connection successful'
    };
  } catch (error) {
    return buildErrorResponse(ERROR_CODES.UNKNOWN_ERROR, 'Connection test failed');
  }
}

/**
 * Simple trigger that runs when a user edits the spreadsheet directly.
 * Auto-triggers stock recalculation if relevant sheets are edited.
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    
    // Defer dynamic calls to prevent loading crashes if config is not resolved
    if (typeof APP_CONFIG === 'undefined') return;

    const billsName = APP_CONFIG.SHEETS.BILL;
    const prodName = APP_CONFIG.SHEETS.PRODUCTION;
    const stockName = APP_CONFIG.SHEETS.STOCK;
    const bomName = APP_CONFIG.SHEETS.BOM;
    const itemsName = APP_CONFIG.SHEETS.ITEMS;

    if (sheetName === billsName || sheetName === prodName || sheetName === stockName || sheetName === bomName) {
      if (typeof recalculateStock === 'function') {
        recalculateStock();
      }
    }

    const dispatchName = APP_CONFIG.SHEETS.DISPATCH;
    if ((sheetName === prodName || sheetName === dispatchName) && typeof recalculateWarehousePool === 'function') {
      recalculateWarehousePool();
    }

    // Keep Stock and Items Master in sync when either is edited directly on
    // the sheet (not through the web app, which already syncs via saveItem/
    // syncStockForItem). Header row edits (row 1) are skipped.
    if (e.range.getRow() > 1) {
      if (sheetName === stockName) {
        // A single-cell edit to Item Name/Size carries e.oldValue — use it to
        // detect an in-place rename and repoint Item Master before
        // importItemsFromStock() runs, so the rename doesn't get treated as
        // a brand-new item (leaving the old name behind as an untracked row).
        if (e.oldValue !== undefined && e.range.getNumRows() === 1 && e.range.getNumColumns() === 1) {
          const col = e.range.getColumn();
          if ((col === STOCK_COL.ITEM_NAME || col === STOCK_COL.SIZE) && typeof renameItemMasterForStockEdit === 'function') {
            const row = e.range.getRow();
            const rowVals = sheet.getRange(row, STOCK_COL.ITEM_NAME, 1, 2).getValues()[0];
            const currentName = String(rowVals[0] || '').trim();
            const currentSize = String(rowVals[1] || '').trim();
            const oldName = col === STOCK_COL.ITEM_NAME ? String(e.oldValue || '').trim() : currentName;
            const oldSize = col === STOCK_COL.SIZE ? String(e.oldValue || '').trim() : currentSize;
            renameItemMasterForStockEdit(oldName, oldSize, currentName, currentSize);
          }
        }
        if (typeof importItemsFromStock === 'function') {
          importItemsFromStock();
        }
      } else if (sheetName === itemsName && typeof syncAllItemsToStock === 'function') {
        syncAllItemsToStock();
      }
    }
  } catch (err) {
    Log.error('[onEdit] Error:', err.message);
  }
}
