/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_audit.gs — INTERNAL LEDGER AUDIT (backend-only)
 *
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Cross-checks the PO, Bill (received), Return, Wastage, and Issued Stock
 * ledgers against each other so calculation mistakes (a mis-linked PO, a
 * typo'd quantity, a return that doesn't match anything ever billed) surface
 * on their own instead of silently compounding. Same philosophy as
 * module_bom.js's getBomProcessComponentsDrift(): read-only, never blocks a
 * save, never auto-fixes anything — a human reviews and corrects from the
 * relevant ledger.
 *
 * This is deliberately NOT wired into the web app UI. It runs unattended on
 * a time-driven trigger (see installInternalLedgerAuditTrigger()) and writes
 * findings to the Logs sheet via logAction() — an internal/backend safety
 * net, not a user-facing feature.
 *
 * Dependencies (shared globals, same Apps Script project namespace):
 * ───────────────────────────────────────────────────────────────────────────
 * - module_po.js (getPOData)
 * - module_bill.js (getBillData, _aggregateBilledBaseQtyByPo)
 * - module_return.js (getReturnData)
 * - module_wastage.js (getWastageData)
 * - module_issue.js (getIssueData)
 * - utils.js / config.js (buildResponse, logAction, Log)
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * _computeInternalLedgerAuditFindings()
 *
 * Pure (no writes, no logging) reconciliation pass across all 5 ledgers —
 * kept separate from runInternalLedgerAudit() so the check logic itself
 * stays trivially testable without a fake Logs sheet.
 *
 * Checks performed (each independent):
 *  1. over_billed_po_line — a PO line's total billed qty exceeds what was
 *     ordered on that line. getPOData()'s own status math (_attachPoStatus)
 *     clamps this at 0 internally (Math.max(0, …) on pendingQty) so it never
 *     surfaces anywhere else — this recomputes from the same unclamped
 *     item.receivedQty instead of touching that shared status math.
 *  2. orphaned_return_bill_ref — a Return names a Bill Number that doesn't
 *     exist for that vendor (the field is free-text, never validated at
 *     entry).
 *  3. return_item_not_on_bill — a Return names a real Bill, but has an item
 *     (name+size) that bill never actually billed.
 *  4. over_consumed_item — total Returned + Wasted + Issued base qty for an
 *     (item, size) exceeds total Billed base qty for that same item+size,
 *     across every vendor combined. Not vendor-scoped: Wastage/Issue both
 *     carry an OPTIONAL vendor field that's usually blank, so keying this to
 *     vendor would silently miss most real cases.
 *
 * @returns {{ findings: Array, message: string }} findings = one entry per
 *   discrepancy, each shaped { type, ...fields }; message summarizes counts
 *   by type.
 * @throws {Error} if any underlying ledger fails to load — caller decides
 *   how to handle/log that.
 */
function _computeInternalLedgerAuditFindings() {
  const billResponse = getBillData();
  if (!billResponse.success) throw new Error('Failed to load Bill data: ' + billResponse.message);
  const bills = billResponse.data || [];

  // Shared with getPOData() below so it doesn't separately re-aggregate the
  // whole Bill Ledger a second time just to derive PO status — same
  // optimization suggestPoAllocations() already relies on.
  const billedMap = _aggregateBilledBaseQtyByPo(bills);
  const poResponse = getPOData(billedMap);
  if (!poResponse.success) throw new Error('Failed to load PO data: ' + poResponse.message);
  const pos = poResponse.data || [];

  const returnResponse = getReturnData();
  if (!returnResponse.success) throw new Error('Failed to load Return data: ' + returnResponse.message);
  const returns = returnResponse.data || [];

  const wastageResponse = getWastageData();
  if (!wastageResponse.success) throw new Error('Failed to load Wastage data: ' + wastageResponse.message);
  const wastageRecords = wastageResponse.data || [];

  const issueResponse = getIssueData();
  if (!issueResponse.success) throw new Error('Failed to load Issued Stock data: ' + issueResponse.message);
  const issueRecords = issueResponse.data || [];

  const findings = [];

  // ── Check 1: over-billed PO lines ──────────────────────────────────
  pos.forEach(function(po) {
    (po.items || []).forEach(function(item) {
      const orderedQty = Number(item.qty) || 0;
      const receivedQty = Number(item.receivedQty) || 0;
      const overBy = receivedQty - orderedQty;
      if (overBy > 0.0001) {
        findings.push({
          type: 'over_billed_po_line',
          poNumber: po.poNumber,
          vendor: po.vendor,
          itemName: item.name,
          size: item.size,
          narration: item.narration,
          unit: item.unit,
          orderedQty: orderedQty,
          billedQty: parseFloat(receivedQty.toFixed(4)),
          overBy: parseFloat(overBy.toFixed(4))
        });
      }
    });
  });

  // ── Checks 2 & 3: Return -> Bill reference integrity ────────────────
  const billByKey = new Map();
  bills.forEach(function(bill) {
    const key = String(bill.vendor || '').trim().toLowerCase() + '|' + String(bill.billNumber || '').trim().toLowerCase();
    billByKey.set(key, bill);
  });

  returns.forEach(function(ret) {
    const billNumber = String(ret.billNumber || '').trim();
    if (!billNumber) return; // "Original Bill #" is optional — nothing to check

    const key = String(ret.vendor || '').trim().toLowerCase() + '|' + billNumber.toLowerCase();
    const bill = billByKey.get(key);

    if (!bill) {
      findings.push({
        type: 'orphaned_return_bill_ref',
        returnNumber: ret.returnNumber,
        vendor: ret.vendor,
        billNumber: billNumber
      });
      return; // nothing to compare items against
    }

    const billItemKeys = new Set(
      (bill.items || []).map(function(i) {
        return String(i.name || '').trim().toLowerCase() + '|' + String(i.size || '').trim().toLowerCase();
      })
    );

    (ret.items || []).forEach(function(item) {
      const itemKey = String(item.name || '').trim().toLowerCase() + '|' + String(item.size || '').trim().toLowerCase();
      if (!billItemKeys.has(itemKey)) {
        findings.push({
          type: 'return_item_not_on_bill',
          returnNumber: ret.returnNumber,
          vendor: ret.vendor,
          billNumber: billNumber,
          itemName: item.name,
          size: item.size
        });
      }
    });
  });

  // ── Check 4: over-consumed items (Return + Wastage + Issue vs Bill) ──
  const itemKeyOf = function(name, size) {
    return String(name || '').trim().toLowerCase() + '|' + String(size || '').trim().toLowerCase();
  };

  const billedBaseQtyByItem = new Map();
  bills.forEach(function(bill) {
    (bill.items || []).forEach(function(item) {
      const key = itemKeyOf(item.name, item.size);
      billedBaseQtyByItem.set(key, (billedBaseQtyByItem.get(key) || 0) + (Number(item.baseQty) || 0));
    });
  });

  const consumedByItem = new Map();
  const addConsumption = function(records, field) {
    records.forEach(function(rec) {
      (rec.items || []).forEach(function(item) {
        const key = itemKeyOf(item.name, item.size);
        if (!consumedByItem.has(key)) {
          consumedByItem.set(key, {
            name: item.name, size: item.size, unit: item.unit,
            returnedBaseQty: 0, wastedBaseQty: 0, issuedBaseQty: 0
          });
        }
        consumedByItem.get(key)[field] += (Number(item.baseQty) || 0);
      });
    });
  };
  addConsumption(returns, 'returnedBaseQty');
  addConsumption(wastageRecords, 'wastedBaseQty');
  addConsumption(issueRecords, 'issuedBaseQty');

  consumedByItem.forEach(function(entry, key) {
    const totalConsumedBaseQty = entry.returnedBaseQty + entry.wastedBaseQty + entry.issuedBaseQty;
    const billedBaseQty = billedBaseQtyByItem.get(key) || 0;
    const overBy = totalConsumedBaseQty - billedBaseQty;
    if (overBy > 0.0001) {
      findings.push({
        type: 'over_consumed_item',
        itemName: entry.name,
        size: entry.size,
        unit: entry.unit,
        totalBilledBaseQty: parseFloat(billedBaseQty.toFixed(4)),
        totalReturnedBaseQty: parseFloat(entry.returnedBaseQty.toFixed(4)),
        totalWastedBaseQty: parseFloat(entry.wastedBaseQty.toFixed(4)),
        totalIssuedBaseQty: parseFloat(entry.issuedBaseQty.toFixed(4)),
        totalConsumedBaseQty: parseFloat(totalConsumedBaseQty.toFixed(4)),
        overBy: parseFloat(overBy.toFixed(4))
      });
    }
  });

  const counts = findings.reduce(function(acc, f) { acc[f.type] = (acc[f.type] || 0) + 1; return acc; }, {});
  const message = findings.length === 0
    ? 'No discrepancies found — PO, Bill, Return, Wastage, and Issued Stock ledgers reconcile cleanly.'
    : `${findings.length} discrepanc${findings.length === 1 ? 'y' : 'ies'} found: ` +
      Object.keys(counts).map(function(t) { return `${counts[t]} ${t.replace(/_/g, ' ')}`; }).join(', ');

  return { findings: findings, message: message };
}

/**
 * _describeAuditFinding(f)
 *
 * One-line, human-readable rendering of a finding for the Logs sheet's
 * free-text Details column.
 */
function _describeAuditFinding(f) {
  switch (f.type) {
    case 'over_billed_po_line':
      return `PO #${f.poNumber} (${f.vendor}): ${f.itemName}${f.size ? ' (' + f.size + ')' : ''} — ordered ${f.orderedQty}, billed ${f.billedQty} ${f.unit || ''} (${f.overBy} over)`;
    case 'orphaned_return_bill_ref':
      return `Return #${f.returnNumber} (${f.vendor}) names Bill #${f.billNumber}, which doesn't exist for this vendor`;
    case 'return_item_not_on_bill':
      return `Return #${f.returnNumber} (${f.vendor}) — ${f.itemName}${f.size ? ' (' + f.size + ')' : ''} isn't on Bill #${f.billNumber}`;
    case 'over_consumed_item':
      return `${f.itemName}${f.size ? ' (' + f.size + ')' : ''} — returned ${f.totalReturnedBaseQty} + wasted ${f.totalWastedBaseQty} + issued ${f.totalIssuedBaseQty} = ${f.totalConsumedBaseQty} consumed, but only ${f.totalBilledBaseQty} ever billed (${f.overBy} over)`;
    default:
      return JSON.stringify(f);
  }
}

/**
 * runInternalLedgerAudit()
 *
 * Backend-only entry point — bound to the scheduled trigger (see
 * installInternalLedgerAuditTrigger()). Not called from the web app.
 *
 * Computes findings via _computeInternalLedgerAuditFindings() and writes one
 * Logs-sheet row per finding (action LEDGER_AUDIT_FINDING, status WARNING)
 * plus one summary row (action LEDGER_AUDIT_SUMMARY, status SUCCESS) every
 * run. The summary row is written even with zero findings, so a gap in the
 * Logs sheet reliably means the trigger stopped firing, not that everything
 * is fine.
 *
 * @returns {Object} API response; data = { findingsCount, findings } —
 *   returned for convenience when invoked manually from the Apps Script
 *   editor, not read by anything else.
 */
function runInternalLedgerAudit() {
  try {
    const result = _computeInternalLedgerAuditFindings();

    result.findings.forEach(function(f) {
      const recordId = f.poNumber || f.returnNumber || (f.itemName ? f.itemName + (f.size ? ' (' + f.size + ')' : '') : 'N/A');
      logAction('LEDGER_AUDIT_FINDING', 'Internal Audit', recordId, f.type + ': ' + _describeAuditFinding(f), 'WARNING');
    });

    logAction('LEDGER_AUDIT_SUMMARY', 'Internal Audit', 'ALL', result.message, 'SUCCESS');

    return buildResponse(true, { findingsCount: result.findings.length, findings: result.findings }, result.message);
  } catch (error) {
    Log.error('[runInternalLedgerAudit] Error:', error.message);
    logAction('LEDGER_AUDIT_SUMMARY', 'Internal Audit', 'ALL', 'Audit run failed: ' + error.message, 'ERROR');
    return buildResponse(false, null, 'Failed to run internal ledger audit: ' + error.message);
  }
}

/**
 * installInternalLedgerAuditTrigger()
 *
 * One-time setup: installs an hourly time-driven trigger that runs
 * runInternalLedgerAudit() unattended — same pattern as
 * installAutoMergeDuplicatesTrigger() in module_items.js. Run manually once
 * from the Apps Script editor. Safe to call again — clears any existing
 * trigger for this handler first, so re-running setup can't leave two
 * triggers fighting over the same document lock every hour.
 */
function installInternalLedgerAuditTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runInternalLedgerAudit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('runInternalLedgerAudit')
    .timeBased()
    .everyHours(1)
    .create();
}
