/**
 * ═══════════════════════════════════════════════════════════════════════════
 * module_dashboard.gs — DASHBOARD MODULE
 *
 * Purpose:
 * ───────────────────────────────────────────────────────────────────────────
 * Single read-only aggregator (getDashboardData) that rolls up KPIs, the
 * process WIP pipeline, production status breakdown, dispatch trend, low
 * stock items, and contractor payables into one response — so the Dashboard
 * tab can render with a single google.script.run round trip instead of one
 * per widget. Every number here is derived from existing sheets/getters; this
 * module owns no sheet of its own.
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * Open POs: PO lines whose Base Qty hasn't been fully billed yet.
 * Reuses module_po.js/module_bill.js's own billed-qty aggregation so this
 * stays in sync with however "billed" is defined there.
 * @returns {{count: number, value: number}}
 * @private
 */
function _getOpenPoSummary() {
  // Computed once and shared with getPOData below (which would otherwise
  // aggregate the entire Bill Ledger a second time just to derive status)
  // instead of each doing its own full Bill-sheet read.
  const billedMap = _aggregateBilledBaseQtyByPo();
  const poResp = getPOData(billedMap);
  if (!poResp.success) return { count: 0, value: 0 };

  let openCount = 0;
  let openValue = 0;

  (poResp.data || []).forEach(po => {
    let hasOpenLine = false;
    (po.items || []).forEach(item => {
      const key = _buildPoLineKey(po.poNumber, item.name, item.size, item.narration);
      const billedBaseQty = billedMap[key] || 0;
      const remainingBaseQty = (item.baseQty || item.qty) - billedBaseQty;
      if (remainingBaseQty > 0.0001) {
        hasOpenLine = true;
        openValue += remainingBaseQty * (item.baseRate || item.price || 0);
      }
    });
    if (hasOpenLine) openCount++;
  });

  return { count: openCount, value: Math.round(openValue * 100) / 100 };
}

/**
 * Sums a ledger's count + value (or qty) for a given calendar month, where
 * monthsAgo=0 is the current month and monthsAgo=1 is the prior month — lets
 * every "this month" KPI also report its "vs last month" comparison from the
 * same records without a second sheet read.
 * @param {Array} records
 * @param {string} dateField - name of the record's YYYY-MM-DD date field
 * @param {string} [amountField] - name of the record's numeric field to sum;
 *   omitted when only a count is needed
 * @param {number} monthsAgo
 * @returns {{count: number, value: number}}
 * @private
 */
function _summarizeByMonth(records, dateField, amountField, monthsAgo) {
  const ref = new Date();
  ref.setMonth(ref.getMonth() - monthsAgo);
  const monthKey = ref.getFullYear() + '-' + String(ref.getMonth() + 1).padStart(2, '0');

  let count = 0;
  let value = 0;
  (records || []).forEach(rec => {
    const raw = rec[dateField];
    if (!raw || !String(raw).startsWith(monthKey)) return;
    count++;
    if (amountField) value += rec[amountField] || 0;
  });

  return { count: count, value: Math.round(value * 100) / 100 };
}

/**
 * Process WIP pipeline: for every active process (sorted by Sequence), the
 * qty currently in motion — summed from Production lots logged under that
 * process whose Status is "Pending" or "In Progress" (Completed/Cancelled
 * lots are excluded; this is "what's being worked right now", not finished
 * output sitting in the Warehouse Pool) — broken down by Model + Size, with
 * each group's title being the concatenation of the two.
 *
 * A lot's Model comes from its Product Name tag (blank for intermediate,
 * non-final-stage lots, which carry no Product tag). Size comes from the
 * lot's Color Breakdown entries (each color sub-group records its own
 * size); a lot with no breakdown contributes a single sizeless group.
 *
 * @param {Array} productionLots - getProductionData().data
 * @returns {Array<{processId, processName, sequence, totalQty, totalLotCount,
 *   groups: Array<{title, qty, lotCount}>}>}
 * @private
 */
function _getPipelineData(productionLots) {
  const processes = typeof getProcessData === 'function' ? (getProcessData(true).data || []) : [];

  const ACTIVE_STATUSES = new Set(['Pending', 'In Progress']);
  const groupsByProcess = {}; // processIdLower -> { title -> { qty, lotCount } }

  (productionLots || []).forEach(lot => {
    if (!ACTIVE_STATUSES.has(lot.status)) return;
    const processKey = String(lot.processId || '').trim().toLowerCase();
    if (!processKey) return;

    if (!groupsByProcess[processKey]) groupsByProcess[processKey] = {};
    const bucket = groupsByProcess[processKey];

    const modelName = String(lot.productName || '').trim();
    const breakdown = Array.isArray(lot.colorBreakdown) ? lot.colorBreakdown : [];

    const addToGroup = (size, qty) => {
      const title = [modelName, String(size || '').trim()].filter(Boolean).join(' ') || 'Unspecified';
      if (!bucket[title]) bucket[title] = { qty: 0, lotCount: 0 };
      bucket[title].qty += qty;
      bucket[title].lotCount += 1;
    };

    if (breakdown.length > 0) {
      breakdown.forEach(cb => addToGroup(cb.size, Number(cb.qty) || 0));
    } else {
      addToGroup('', lot.qty);
    }
  });

  return processes
    .map(p => {
      const bucket = groupsByProcess[p.processId.toLowerCase()] || {};
      const groups = Object.keys(bucket)
        .map(title => ({
          title: title,
          qty: Math.round(bucket[title].qty * 100) / 100,
          lotCount: bucket[title].lotCount
        }))
        .sort((a, b) => b.qty - a.qty);

      return {
        processId: p.processId,
        processName: p.processName,
        sequence: p.sequence,
        totalQty: Math.round(groups.reduce((sum, g) => sum + g.qty, 0) * 100) / 100,
        totalLotCount: groups.reduce((sum, g) => sum + g.lotCount, 0),
        groups: groups
      };
    })
    // Only processes with at least one active (Pending/In Progress) lot —
    // otherwise every Process Master row shows up as its own empty stage
    // ("No active lots") whenever nothing is actually in motion.
    .filter(p => p.totalLotCount > 0)
    .sort((a, b) => a.sequence - b.sequence);
}

/**
 * Production lot counts grouped by Status.
 * @returns {Array<{status: string, count: number}>}
 * @private
 */
function _getProductionStatusBreakdown(productionLots) {
  const counts = {};
  productionLots.forEach(lot => {
    const status = lot.status || 'Pending';
    counts[status] = (counts[status] || 0) + 1;
  });
  return Object.keys(counts).map(status => ({ status, count: counts[status] }));
}

/**
 * Dispatched quantity per day for the trailing 30 days (oldest first), so
 * a 0-qty day still shows up as a gap in the trend chart.
 * @returns {Array<{date: string, qty: number}>} date is YYYY-MM-DD
 * @private
 */
function _getDispatchTrend(dispatchRecords) {
  const DAYS = 30;
  const qtyByDate = {};
  dispatchRecords.forEach(d => {
    if (!d.dateRaw) return;
    const dateKey = d.dateRaw.slice(0, 10);
    qtyByDate[dateKey] = (qtyByDate[dateKey] || 0) + d.qty;
  });

  const trend = [];
  const today = new Date();
  for (let i = DAYS - 1; i >= 0; i--) {
    const d = new Date(today);
    d.setDate(d.getDate() - i);
    const dateKey = Utilities.formatDate(d, _getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    trend.push({ date: dateKey, qty: Math.round((qtyByDate[dateKey] || 0) * 100) / 100 });
  }
  return trend;
}

/**
 * Single aggregator backing the Dashboard tab. Pure read — composes existing
 * getters/maps instead of touching any sheet directly.
 */
function getDashboardData() {
  try {
    const stockResp = getStockData();
    const stockRecords = (stockResp && stockResp.data) || [];

    const productionResp = getProductionData();
    const productionLots = (productionResp && productionResp.data) || [];

    const dispatchResp = getDispatchData();
    const dispatchRecords = (dispatchResp && dispatchResp.data) || [];

    const readyResp = getReadyToDispatchData();
    const readyRecords = (readyResp && readyResp.data) || [];

    const contractorLedgerResp = getContractorLedgerData();
    const contractorLedger = (contractorLedgerResp && contractorLedgerResp.data) || [];

    const returnResp = getReturnData();
    const returnRecords = (returnResp && returnResp.data) || [];

    const wastageResp = getWastageData();
    const wastageRecords = (wastageResp && wastageResp.data) || [];

    const lowStockFull = stockRecords
      .filter(item => item.isLowStock)
      .map(item => ({
        name: item.name,
        size: item.size,
        currentStock: item.currentStock,
        threshold: item.threshold,
        deficit: Math.round((item.threshold - item.currentStock) * 100) / 100
      }))
      .sort((a, b) => b.deficit - a.deficit);
    const lowStockItems = lowStockFull.slice(0, 10);
    const lowStockTotalDeficit = Math.round(
      lowStockFull.reduce((sum, i) => sum + Math.max(i.deficit, 0), 0) * 100
    ) / 100;

    const activeProductionLots = productionLots.filter(l =>
      l.status === 'Pending' || l.status === 'In Progress');
    const pendingProductionCount = activeProductionLots.length;
    const oldestActiveLotDateRaw = activeProductionLots.reduce((oldest, l) => {
      if (!l.dateRaw) return oldest;
      return (!oldest || l.dateRaw < oldest) ? l.dateRaw : oldest;
    }, null);
    const oldestPendingProductionDays = oldestActiveLotDateRaw
      ? Math.max(0, Math.floor((Date.now() - new Date(oldestActiveLotDateRaw).getTime()) / 86400000))
      : null;

    const readyToDispatchUnits = readyRecords.reduce((sum, r) => sum + Math.max(r.readyQty, 0), 0);

    const readyToDispatchFull = readyRecords
      .filter(r => r.readyQty > 0.0001)
      .map(r => ({ productId: r.productId, productName: r.productName, readyQty: Math.round(r.readyQty * 100) / 100 }))
      .sort((a, b) => b.readyQty - a.readyQty);
    const readyToDispatchItems = readyToDispatchFull.slice(0, 10);

    const contractorPayablesFull = contractorLedger
      .map(c => ({ contractorName: c.contractorName, balanceDue: Math.round(c.balanceDue * 100) / 100 }))
      .filter(c => c.balanceDue > 0.0001)
      .sort((a, b) => b.balanceDue - a.balanceDue);
    const contractorPayablesDue = contractorPayablesFull.slice(0, 10);

    const billResp = getBillData();
    const billRecords = (billResp && billResp.data) || [];

    const openPoSummary = _getOpenPoSummary();
    const billsThisMonth = _summarizeByMonth(billRecords, 'billDateRaw', 'totalAmount', 0);
    const billsLastMonth = _summarizeByMonth(billRecords, 'billDateRaw', 'totalAmount', 1);
    const returnsThisMonth = _summarizeByMonth(returnRecords, 'returnDateRaw', 'totalAmount', 0);
    const returnsLastMonth = _summarizeByMonth(returnRecords, 'returnDateRaw', 'totalAmount', 1);
    const wastageThisMonth = _summarizeByMonth(wastageRecords, 'dateRaw', 'totalQty', 0);
    const wastageLastMonth = _summarizeByMonth(wastageRecords, 'dateRaw', 'totalQty', 1);
    const totalContractorPayableDue = contractorLedger.reduce((sum, c) => sum + Math.max(c.balanceDue, 0), 0);

    return buildResponse(true, {
      generatedAt: new Date().toISOString(),
      kpis: {
        openPoCount: openPoSummary.count,
        openPoValue: openPoSummary.value,
        billsThisMonthCount: billsThisMonth.count,
        billsThisMonthValue: billsThisMonth.value,
        billsLastMonthCount: billsLastMonth.count,
        billsLastMonthValue: billsLastMonth.value,
        returnsThisMonthCount: returnsThisMonth.count,
        returnsThisMonthValue: returnsThisMonth.value,
        returnsLastMonthCount: returnsLastMonth.count,
        returnsLastMonthValue: returnsLastMonth.value,
        wastageThisMonthCount: wastageThisMonth.count,
        wastageThisMonthQty: wastageThisMonth.value,
        wastageLastMonthCount: wastageLastMonth.count,
        wastageLastMonthQty: wastageLastMonth.value,
        lowStockCount: lowStockFull.length,
        lowStockTotalDeficit: lowStockTotalDeficit,
        pendingProductionCount: pendingProductionCount,
        oldestPendingProductionDays: oldestPendingProductionDays,
        readyToDispatchUnits: Math.round(readyToDispatchUnits * 100) / 100,
        readyToDispatchProductCount: readyToDispatchFull.length,
        contractorPayablesDue: Math.round(totalContractorPayableDue * 100) / 100,
        contractorPayablesCount: contractorPayablesFull.length
      },
      pipeline: _getPipelineData(productionLots),
      productionStatusBreakdown: _getProductionStatusBreakdown(productionLots),
      dispatchTrend: _getDispatchTrend(dispatchRecords),
      lowStockItems: lowStockItems,
      lowStockTotalCount: lowStockFull.length,
      contractorPayables: contractorPayablesDue,
      contractorPayablesTotalCount: contractorPayablesFull.length,
      readyToDispatchItems: readyToDispatchItems,
      readyToDispatchTotalCount: readyToDispatchFull.length
    });
  } catch (error) {
    console.error('[getDashboardData] Error:', error.message);
    return buildResponse(false, null, 'Failed to load dashboard data: ' + error.message);
  }
}

/**
 * Single aggregator backing the mobile Home tab: pending production lots,
 * today's dispatch count, low-stock alert count, and the 5 most recent
 * activity entries (production lots + dispatches, newest first) — one
 * round trip instead of one per stat card. Every number is derived from
 * existing getters, same principle as getDashboardData above.
 */
function getMobileDashboard() {
  try {
    const productionResp = getProductionData();
    const productionLots = (productionResp && productionResp.data) || [];

    const dispatchResp = getDispatchData();
    const dispatchRecords = (dispatchResp && dispatchResp.data) || [];

    const stockResp = getStockData();
    const stockRecords = (stockResp && stockResp.data) || [];

    const todayKey = Utilities.formatDate(new Date(), _getSpreadsheetTimeZone(), 'yyyy-MM-dd');

    const pendingProductionCount = productionLots.filter(l =>
      l.status === 'Pending' || l.status === 'In Progress').length;

    const todaysDispatchCount = dispatchRecords.filter(d =>
      d.dateRaw && d.dateRaw.slice(0, 10) === todayKey).length;

    const lowStockCount = stockRecords.filter(s => s.isLowStock).length;

    const activity = [];
    productionLots.forEach(l => {
      if (!l.dateRaw) return;
      activity.push({
        type: 'production',
        title: l.lotNumber || l.processId,
        subtitle: `${l.qty} unit(s) · ${l.status}`,
        dateRaw: l.dateRaw
      });
    });
    dispatchRecords.forEach(d => {
      if (!d.dateRaw) return;
      activity.push({
        type: 'dispatch',
        title: d.dispatchNumber,
        subtitle: `${d.clientName || 'Direct supply'} · ${d.qty} unit(s)`,
        dateRaw: d.dateRaw
      });
    });
    activity.sort((a, b) => new Date(b.dateRaw) - new Date(a.dateRaw));

    return buildResponse(true, {
      pendingProductionCount: pendingProductionCount,
      todaysDispatchCount: todaysDispatchCount,
      lowStockCount: lowStockCount,
      recentActivity: activity.slice(0, 5),
      // Piggybacked here (rather than a dedicated endpoint) for the mobile
      // More tab's About row — this is still the one new function this
      // app adds, not a second one.
      appVersion: APP_CONFIG.APP_VERSION,
      userEmail: _getActiveUserEmail()
    });
  } catch (error) {
    console.error('[getMobileDashboard] Error:', error.message);
    return buildResponse(false, null, 'Failed to load mobile dashboard: ' + error.message);
  }
}
