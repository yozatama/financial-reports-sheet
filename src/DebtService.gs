/**
 * DebtService.gs
 * -------------------------------------------------------------
 * CRUD and analytics for the Debts sheet:
 *  - utilization (outstanding / limit)
 *  - debt-to-income ratio
 *  - upcoming due-date list
 *  - payoff timeline given monthly installments
 * -------------------------------------------------------------
 */

var DebtService = (function () {

  function add(d) {
    require_(d, 'Debt payload required');
    require_(d.name, 'Debt name required');
    require_(d.type, 'Debt type required');
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.DEBTS);
    var row = [
      d.id || genId('DEB'),
      d.name,
      d.provider || '',
      d.type,
      parseFloat(d.outstanding) || 0,
      parseFloat(d.limit) || 0,
      parseFloat(d.minPayment) || 0,
      parseFloat(d.interest) || 0,
      parseInt(d.dueDay, 10) || '',
      parseInt(d.billingDay, 10) || '',
      parseFloat(d.installment) || 0,
      parseInt(d.tenor, 10) || 0,
      parseInt(d.remainingTenor, 10) || parseInt(d.tenor, 10) || 0,
      d.status || STATUS.ACTIVE,
      d.notes || '',
      new Date()
    ];
    sh.appendRow(row);
    return row;
  }

  function update(id, patch) {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.DEBTS);
    var rowIdx = findRowByValue_(sh, COLS.DEBTS.ID, id);
    if (rowIdx < 0) throw new Error('Debt not found: ' + id);
    var row = sh.getRange(rowIdx, 1, 1, 16).getValues()[0];
    var map = COLS.DEBTS;
    function set(col, val) { if (val !== undefined) row[col - 1] = val; }
    set(map.NAME, patch.name);
    set(map.PROVIDER, patch.provider);
    set(map.TYPE, patch.type);
    if (patch.outstanding !== undefined) set(map.OUTSTANDING, parseFloat(patch.outstanding) || 0);
    if (patch.limit !== undefined) set(map.LIMIT, parseFloat(patch.limit) || 0);
    if (patch.minPayment !== undefined) set(map.MIN_PAYMENT, parseFloat(patch.minPayment) || 0);
    if (patch.interest !== undefined) set(map.INTEREST, parseFloat(patch.interest) || 0);
    if (patch.dueDay !== undefined) set(map.DUE_DAY, parseInt(patch.dueDay, 10));
    if (patch.billingDay !== undefined) set(map.BILLING_DAY, parseInt(patch.billingDay, 10));
    if (patch.installment !== undefined) set(map.INSTALLMENT, parseFloat(patch.installment) || 0);
    if (patch.tenor !== undefined) set(map.TENOR, parseInt(patch.tenor, 10));
    if (patch.remainingTenor !== undefined) set(map.REMAINING_TENOR, parseInt(patch.remainingTenor, 10));
    set(map.STATUS, patch.status);
    set(map.NOTES, patch.notes);
    row[map.UPDATED_AT - 1] = new Date();
    sh.getRange(rowIdx, 1, 1, 16).setValues([row]);
    return row;
  }

  function remove(id) {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.DEBTS);
    var rowIdx = findRowByValue_(sh, COLS.DEBTS.ID, id);
    if (rowIdx < 0) throw new Error('Debt not found: ' + id);
    sh.deleteRow(rowIdx);
    return { ok: true };
  }

  function list(activeOnly) {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.DEBTS);
    if (!sh) return [];
    var rows = readObjects_(sh);
    return activeOnly ? rows.filter(function (r) { return r.Status === STATUS.ACTIVE; }) : rows;
  }

  function totalOutstanding() {
    return list(true).reduce(function (s, r) { return s + (parseFloat(r['Outstanding Balance']) || 0); }, 0);
  }

  function totalMonthlyObligation() {
    return list(true).reduce(function (s, r) {
      var min = parseFloat(r['Min Payment']) || 0;
      var inst = parseFloat(r['Monthly Installment']) || 0;
      return s + Math.max(min, inst);
    }, 0);
  }

  function utilizationByDebt() {
    return list(true).map(function (r) {
      var lim = parseFloat(r['Credit Limit']) || 0;
      var out = parseFloat(r['Outstanding Balance']) || 0;
      return {
        id: r['Debt ID'],
        name: r['Debt Name'],
        utilization: lim > 0 ? out / lim : 0,
        outstanding: out, limit: lim
      };
    });
  }

  function avgUtilization() {
    var u = utilizationByDebt().filter(function (r) { return r.limit > 0; });
    if (!u.length) return 0;
    return u.reduce(function (s, r) { return s + r.utilization; }, 0) / u.length;
  }

  /** Up to N upcoming due dates within next 30 days. */
  function upcomingDueDates(daysAhead) {
    daysAhead = daysAhead || 30;
    var today = new Date();
    var todayD = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    var horizon = new Date(todayD.getTime() + daysAhead * 86400000);
    var result = [];
    list(true).forEach(function (r) {
      var day = parseInt(r['Due Day'], 10);
      if (!day) return;
      // build the next due date >= today
      var candidate = new Date(today.getFullYear(), today.getMonth(), day);
      if (candidate < todayD) candidate = new Date(today.getFullYear(), today.getMonth() + 1, day);
      if (candidate <= horizon) {
        var diffDays = Math.round((candidate - todayD) / 86400000);
        result.push({
          name: r['Debt Name'],
          provider: r['Provider'],
          dueDate: candidate,
          daysUntil: diffDays,
          minPayment: parseFloat(r['Min Payment']) || 0,
          outstanding: parseFloat(r['Outstanding Balance']) || 0
        });
      }
    });
    result.sort(function (a, b) { return a.daysUntil - b.daysUntil; });
    return result;
  }

  /** Estimate months to clear each debt at current installment. */
  function payoffTimeline() {
    return list(true).map(function (r) {
      var out = parseFloat(r['Outstanding Balance']) || 0;
      var inst = Math.max(parseFloat(r['Monthly Installment']) || 0, parseFloat(r['Min Payment']) || 0);
      var months = inst > 0 ? Math.ceil(out / inst) : null;
      return { name: r['Debt Name'], outstanding: out, monthlyPayment: inst, monthsToPayoff: months };
    });
  }

  return {
    add: add, update: update, remove: remove, list: list,
    totalOutstanding: totalOutstanding,
    totalMonthlyObligation: totalMonthlyObligation,
    utilizationByDebt: utilizationByDebt,
    avgUtilization: avgUtilization,
    upcomingDueDates: upcomingDueDates,
    payoffTimeline: payoffTimeline
  };
})();
