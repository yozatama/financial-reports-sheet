/**
 * DashboardService.gs
 * -------------------------------------------------------------
 * Builds and refreshes the Dashboard sheet. Uses native Sheets
 * cells + formulas (no charts API churn) so the dashboard stays
 * fast and renders identically on mobile.
 *
 * Layout (12-column grid, columns A..L):
 *
 *   Row 1-2  : Title bar + subtitle
 *   Row 4-7  : KPI cards x 6  (Total Balance, Income, Expense,
 *              Net, Savings Rate, Debt Ratio)
 *   Row 9-10 : Section header "Spending Analysis"
 *   Row 11-22: Top categories table  (left)  +  Trend (right)
 *   Row 24-25: Section header "Debt Monitoring"
 *   Row 26-34: Outstanding by debt + Upcoming due dates
 *   Row 36-37: Section header "Cash Flow"
 *   Row 38-46: Inflow vs outflow table + Account breakdown
 *   Row 48-49: Section header "Insights"
 *   Row 50-58: Auto-generated bullets
 * -------------------------------------------------------------
 */

var DashboardService = (function () {

  function refresh() { layoutAndFill_(true); }
  // Lightweight refresh after writes. Re-uses the same redraw path
  // because partial updates would conflict with existing merged ranges.
  function refreshLightweight() { layoutAndFill_(true); }

  /** Static layout (called once during setup, mostly cosmetic). */
  function layout(sh) { layoutAndFill_(true, sh); }

  function layoutAndFill_(redraw, sheetArg) {
    var ss = SpreadsheetApp.getActive();
    var sh = sheetArg || ss.getSheetByName(SHEETS.DASHBOARD);
    if (!sh) return;

    if (redraw) {
      sh.clear();
      // sh.clear() does not break merged ranges; do it explicitly to avoid layout drift.
      try { sh.getRange(1, 1, 200, 12).breakApart(); } catch (e) { /* no merges yet */ }
      sh.setHiddenGridlines(true);
      for (var c = 1; c <= 12; c++) sh.setColumnWidth(c, 110);
      drawTitle_(sh);
    }

    var month = new Date();
    var totals = TransactionService.monthTotals(month);
    var totalBalance = AccountService.totalBalance();
    var incomeTarget = parseFloat(getSetting('monthly_income_target', 0)) || 0;
    var income = totals.income || 0;
    var expense = totals.expense || 0;
    var net = income - expense;
    var savingsRate = (income > 0) ? net / income : (incomeTarget > 0 ? net / incomeTarget : 0);
    var monthlyObligation = DebtService.totalMonthlyObligation();
    var totalDebt = DebtService.totalOutstanding();
    var debtRatio = (income > 0) ? monthlyObligation / income : (incomeTarget > 0 ? monthlyObligation / incomeTarget : 0);

    drawKPI_(sh, 4, 1, 'Total Balance', fmtMoney_(totalBalance), 'across active accounts', THEME.PRIMARY);
    drawKPI_(sh, 4, 3, 'Monthly Income', fmtMoney_(income), Utilities.formatDate(month, APP.DEFAULT_TIMEZONE, 'MMMM yyyy'), THEME.SUCCESS);
    drawKPI_(sh, 4, 5, 'Monthly Expense', fmtMoney_(expense), totals.count + ' transactions', THEME.DANGER);
    drawKPI_(sh, 4, 7, 'Net Cash Flow', fmtMoney_(net), net >= 0 ? 'Positive' : 'Negative', net >= 0 ? THEME.ACCENT : THEME.WARNING);
    drawKPI_(sh, 4, 9, 'Savings Rate', fmtPct_(savingsRate), savingsRate >= THRESHOLDS.HEALTHY_SAVINGS_RATE ? 'Healthy' : 'Below target', savingsRate >= THRESHOLDS.HEALTHY_SAVINGS_RATE ? THEME.SUCCESS : THEME.WARNING);
    drawKPI_(sh, 4, 11, 'Debt Ratio', fmtPct_(debtRatio), debtRatio <= THRESHOLDS.HEALTHY_DEBT_RATIO ? 'Within healthy range' : 'Above healthy threshold', debtRatio <= THRESHOLDS.HEALTHY_DEBT_RATIO ? THEME.SUCCESS : THEME.DANGER);

    drawSectionHeader_(sh, 9, 'Spending Analysis');
    drawTopCategories_(sh, 11);
    drawMonthTrend_(sh, 11);

    drawSectionHeader_(sh, 24, 'Debt Monitoring');
    drawDebtTable_(sh, 26);
    drawUpcomingDue_(sh, 26);

    drawSectionHeader_(sh, 36, 'Cash Flow');
    drawAccountBreakdown_(sh, 38);
    drawInflowOutflow_(sh, 38);

    drawSectionHeader_(sh, 48, 'Insights');
    drawInsights_(sh, 50, { income: income, expense: expense, net: net, debtRatio: debtRatio, savingsRate: savingsRate, totalDebt: totalDebt });

    drawSectionHeader_(sh, 60, 'Quick Actions');
    drawQuickActionsHelp_(sh, 62);
  }

  /* -----------------------------------------------------------
   * Drawing helpers
   * ----------------------------------------------------------- */

  function drawTitle_(sh) {
    sh.getRange(1, 1, 1, 12).merge()
      .setValue(APP.NAME + '  ·  Personal Finance Dashboard')
      .setFontWeight('bold').setFontSize(20)
      .setBackground(THEME.HEADER_BG).setFontColor(THEME.HEADER_FG)
      .setVerticalAlignment('middle').setHorizontalAlignment('left');
    sh.setRowHeight(1, 50);
    sh.getRange(2, 1, 1, 12).merge()
      .setValue('Live overview of your wallets, debts, and spending. Refreshed: ' + Utilities.formatDate(new Date(), APP.DEFAULT_TIMEZONE, 'yyyy-MM-dd HH:mm'))
      .setFontColor(THEME.MUTED).setFontSize(11).setHorizontalAlignment('left');
    sh.setRowHeight(2, 24);
  }

  function drawKPI_(sh, row, col, title, value, sub, color) {
    // Sub-merge layout: title (1 row), value (2 rows), sub (1 row) inside a 4x2 card.
    sh.getRange(row, col, 1, 2).merge()
      .setValue(title).setFontSize(10).setFontColor(THEME.MUTED)
      .setBackground(THEME.CARD).setVerticalAlignment('bottom');
    sh.getRange(row + 1, col, 2, 2).merge()
      .setValue(value).setFontSize(18).setFontWeight('bold').setFontColor(color || '#111827')
      .setBackground(THEME.CARD).setVerticalAlignment('middle');
    sh.getRange(row + 3, col, 1, 2).merge()
      .setValue(sub).setFontSize(10).setFontColor(THEME.MUTED)
      .setBackground(THEME.CARD).setVerticalAlignment('top');
    sh.getRange(row, col, 4, 2)
      .setBorder(true, true, true, true, false, false, THEME.BORDER, SpreadsheetApp.BorderStyle.SOLID);
    for (var i = 0; i < 4; i++) sh.setRowHeight(row + i, 24);
  }

  function drawSectionHeader_(sh, row, label) {
    sh.getRange(row, 1, 1, 12).merge()
      .setValue(label)
      .setFontWeight('bold').setFontSize(13)
      .setBackground('#F3F4F6').setFontColor('#111827')
      .setVerticalAlignment('middle').setHorizontalAlignment('left');
    sh.setRowHeight(row, 32);
  }

  /** Top categories table (left half). */
  function drawTopCategories_(sh, row) {
    var month = monthKey_(new Date());
    var data = aggregateCategoryThisMonth_();
    sh.getRange(row, 1, 1, 6).merge()
      .setValue('Top Spending Categories · ' + month)
      .setFontWeight('bold').setBackground(THEME.CARD).setFontColor('#111827');
    sh.getRange(row + 1, 1, 1, 6).setValues([['Category', 'Amount', 'Share', 'Bar', '', '']])
      .setFontWeight('bold').setBackground('#F9FAFB').setFontColor(THEME.MUTED);

    var top = data.slice(0, 8);
    var total = top.reduce(function (s, r) { return s + r.amount; }, 0) || 1;
    var rows = top.map(function (r) {
      var pct = r.amount / total;
      var bar = bar_(pct, 14);
      return [r.category, r.amount, pct, bar, '', ''];
    });
    if (!rows.length) rows = [['No expenses this month', 0, 0, '', '', '']];
    sh.getRange(row + 2, 1, rows.length, 6).setValues(rows);
    sh.getRange(row + 2, 2, rows.length).setNumberFormat('#,##0');
    sh.getRange(row + 2, 3, rows.length).setNumberFormat('0.0%');
    sh.getRange(row + 2, 4, rows.length).setFontFamily('Roboto Mono').setFontColor(THEME.PRIMARY);
    sh.getRange(row + 1, 1, rows.length + 1, 6).setBackground(THEME.CARD)
      .setBorder(true, true, true, true, false, false, THEME.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }

  /** Daily expense trend (right half). */
  function drawMonthTrend_(sh, row) {
    var trend = dailyExpenseTrend_();
    sh.getRange(row, 7, 1, 6).merge()
      .setValue('Daily Expense Trend (last 14 days)')
      .setFontWeight('bold').setBackground(THEME.CARD).setFontColor('#111827');
    sh.getRange(row + 1, 7, 1, 6).setValues([['Date', 'Amount', 'Bar', '', '', '']])
      .setFontWeight('bold').setBackground('#F9FAFB').setFontColor(THEME.MUTED);

    var max = trend.reduce(function (m, r) { return Math.max(m, r.amount); }, 0) || 1;
    var rows = trend.map(function (r) {
      return [fmtDate_(r.date), r.amount, bar_(r.amount / max, 18), '', '', ''];
    });
    sh.getRange(row + 2, 7, rows.length, 6).setValues(rows);
    sh.getRange(row + 2, 8, rows.length).setNumberFormat('#,##0');
    sh.getRange(row + 2, 9, rows.length).setFontFamily('Roboto Mono').setFontColor(THEME.DANGER);
    sh.getRange(row + 1, 7, rows.length + 1, 6).setBackground(THEME.CARD)
      .setBorder(true, true, true, true, false, false, THEME.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }

  function drawDebtTable_(sh, row) {
    sh.getRange(row, 1, 1, 6).merge()
      .setValue('Outstanding by Debt')
      .setFontWeight('bold').setBackground(THEME.CARD).setFontColor('#111827');
    sh.getRange(row + 1, 1, 1, 6).setValues([['Debt', 'Outstanding', 'Limit', 'Util.', 'Min Pay', 'Status']])
      .setFontWeight('bold').setBackground('#F9FAFB').setFontColor(THEME.MUTED);
    var debts = DebtService.list(true);
    var rows = debts.map(function (d) {
      var lim = parseFloat(d['Credit Limit']) || 0;
      var out = parseFloat(d['Outstanding Balance']) || 0;
      var util = lim > 0 ? out / lim : 0;
      return [d['Debt Name'], out, lim, util, parseFloat(d['Min Payment']) || 0, d['Status']];
    });
    if (!rows.length) rows = [['No active debts', 0, 0, 0, 0, '']];
    sh.getRange(row + 2, 1, rows.length, 6).setValues(rows);
    sh.getRange(row + 2, 2, rows.length, 2).setNumberFormat('#,##0');
    sh.getRange(row + 2, 4, rows.length).setNumberFormat('0%');
    sh.getRange(row + 2, 5, rows.length).setNumberFormat('#,##0');
    sh.getRange(row + 1, 1, rows.length + 1, 6).setBackground(THEME.CARD)
      .setBorder(true, true, true, true, false, false, THEME.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }

  function drawUpcomingDue_(sh, row) {
    sh.getRange(row, 7, 1, 6).merge()
      .setValue('Upcoming Due Dates')
      .setFontWeight('bold').setBackground(THEME.CARD).setFontColor('#111827');
    sh.getRange(row + 1, 7, 1, 6).setValues([['Debt', 'Due Date', 'In Days', 'Min Pay', 'Status', '']])
      .setFontWeight('bold').setBackground('#F9FAFB').setFontColor(THEME.MUTED);

    var dues = DebtService.upcomingDueDates(30);
    var rows = dues.map(function (d) {
      var status = d.daysUntil <= 3 ? '⚠ Soon' : (d.daysUntil <= 7 ? 'This Week' : 'Upcoming');
      return [d.name, fmtDate_(d.dueDate), d.daysUntil, d.minPayment, status, ''];
    });
    if (!rows.length) rows = [['No upcoming dues', '', '', 0, 'OK', '']];
    sh.getRange(row + 2, 7, rows.length, 6).setValues(rows);
    sh.getRange(row + 2, 10, rows.length).setNumberFormat('#,##0');
    sh.getRange(row + 1, 7, rows.length + 1, 6).setBackground(THEME.CARD)
      .setBorder(true, true, true, true, false, false, THEME.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }

  function drawAccountBreakdown_(sh, row) {
    sh.getRange(row, 1, 1, 6).merge()
      .setValue('Account Balances')
      .setFontWeight('bold').setBackground(THEME.CARD).setFontColor('#111827');
    sh.getRange(row + 1, 1, 1, 6).setValues([['Account', 'Type', 'Balance', 'Bar', '', '']])
      .setFontWeight('bold').setBackground('#F9FAFB').setFontColor(THEME.MUTED);

    var accs = AccountService.list(true);
    var max = accs.reduce(function (m, a) { return Math.max(m, parseFloat(a['Current Balance']) || 0); }, 0) || 1;
    var rows = accs.map(function (a) {
      var bal = parseFloat(a['Current Balance']) || 0;
      return [a['Account Name'], a['Account Type'], bal, bar_(Math.max(0, bal) / max, 18), '', ''];
    });
    if (!rows.length) rows = [['No accounts', '', 0, '', '', '']];
    sh.getRange(row + 2, 1, rows.length, 6).setValues(rows);
    sh.getRange(row + 2, 3, rows.length).setNumberFormat('#,##0');
    sh.getRange(row + 2, 4, rows.length).setFontFamily('Roboto Mono').setFontColor(THEME.PRIMARY);
    sh.getRange(row + 1, 1, rows.length + 1, 6).setBackground(THEME.CARD)
      .setBorder(true, true, true, true, false, false, THEME.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }

  function drawInflowOutflow_(sh, row) {
    sh.getRange(row, 7, 1, 6).merge()
      .setValue('Inflow vs Outflow (last 6 months)')
      .setFontWeight('bold').setBackground(THEME.CARD).setFontColor('#111827');
    sh.getRange(row + 1, 7, 1, 6).setValues([['Month', 'Income', 'Expense', 'Net', 'Trend', '']])
      .setFontWeight('bold').setBackground('#F9FAFB').setFontColor(THEME.MUTED);

    var months = sixMonthSummary_();
    var maxAmt = months.reduce(function (m, r) { return Math.max(m, r.income, r.expense); }, 0) || 1;
    var rows = months.map(function (r) {
      var trend = bar_(r.income / maxAmt, 6) + ' ' + bar_(r.expense / maxAmt, 6);
      return [r.month, r.income, r.expense, r.income - r.expense, trend, ''];
    });
    if (!rows.length) rows = [['No data', 0, 0, 0, '', '']];
    sh.getRange(row + 2, 7, rows.length, 6).setValues(rows);
    sh.getRange(row + 2, 8, rows.length, 3).setNumberFormat('#,##0');
    sh.getRange(row + 2, 11, rows.length).setFontFamily('Roboto Mono');
    sh.getRange(row + 1, 7, rows.length + 1, 6).setBackground(THEME.CARD)
      .setBorder(true, true, true, true, false, false, THEME.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }

  function drawInsights_(sh, row, m) {
    var lines = [];
    var prev = previousMonthTotals_();
    if (m.income > 0 && prev.expense > 0) {
      var diff = (m.expense - prev.expense) / prev.expense;
      if (Math.abs(diff) > 0.05) {
        lines.push((diff >= 0 ? '🔺 Spending increased ' : '🔻 Spending decreased ') +
          fmtPct_(Math.abs(diff)) + ' compared to last month.');
      }
    }
    if (m.savingsRate >= THRESHOLDS.HEALTHY_SAVINGS_RATE) {
      lines.push('✅ Your savings rate (' + fmtPct_(m.savingsRate) + ') is on track.');
    } else if (m.income > 0) {
      lines.push('⚠️ Savings rate (' + fmtPct_(m.savingsRate) + ') is below the ' + fmtPct_(THRESHOLDS.HEALTHY_SAVINGS_RATE) + ' target.');
    }
    if (m.debtRatio > THRESHOLDS.HEALTHY_DEBT_RATIO) {
      lines.push('🚨 Debt obligations (' + fmtPct_(m.debtRatio) + ' of income) exceed the healthy threshold (' + fmtPct_(THRESHOLDS.HEALTHY_DEBT_RATIO) + ').');
    } else if (m.totalDebt > 0) {
      lines.push('🟢 Debt obligations are within a healthy range.');
    }
    var topCat = aggregateCategoryThisMonth_()[0];
    if (topCat) lines.push('💡 Largest expense category this month: ' + topCat.category + ' (' + fmtMoney_(topCat.amount) + ').');
    var avgUtil = DebtService.avgUtilization();
    if (avgUtil > THRESHOLDS.HEALTHY_CREDIT_UTIL) {
      lines.push('💳 Average credit utilization is ' + fmtPct_(avgUtil) + ' — try to keep it below ' + fmtPct_(THRESHOLDS.HEALTHY_CREDIT_UTIL) + '.');
    }
    if (!lines.length) lines.push('Start adding transactions to see personalised insights.');

    var max = 8;
    for (var i = 0; i < max; i++) {
      var cell = sh.getRange(row + i, 1, 1, 12);
      cell.merge()
        .setValue(i < lines.length ? lines[i] : '')
        .setBackground(THEME.CARD).setFontColor('#111827').setFontSize(11)
        .setVerticalAlignment('middle').setHorizontalAlignment('left');
    }
    sh.getRange(row, 1, max, 12)
      .setBorder(true, true, true, true, false, false, THEME.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }

  function drawQuickActionsHelp_(sh, row) {
    var msg = 'Use the "' + APP.NAME + '" menu above to: Add Transaction · Quick Add (AI) · Add Account · Add Debt · Refresh Dashboard.';
    sh.getRange(row, 1, 1, 12).merge()
      .setValue(msg)
      .setBackground(THEME.CARD).setFontColor(THEME.MUTED).setFontSize(11)
      .setHorizontalAlignment('left').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, THEME.BORDER, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row, 36);
  }

  /* -----------------------------------------------------------
   * Aggregation helpers
   * ----------------------------------------------------------- */

  function aggregateCategoryThisMonth_() {
    var d = new Date(), from = startOfMonth_(d), to = endOfMonth_(d);
    var map = {};
    TransactionService.list().forEach(function (r) {
      if (r.Type !== TYPES.EXPENSE) return;
      if (!r.Date) return;
      var dt = new Date(r.Date);
      if (dt < from || dt > to) return;
      var cat = r.Category || 'Other';
      map[cat] = (map[cat] || 0) + (parseFloat(r.Amount) || 0);
    });
    return Object.keys(map).map(function (k) { return { category: k, amount: map[k] }; })
      .sort(function (a, b) { return b.amount - a.amount; });
  }

  function dailyExpenseTrend_() {
    var days = 14;
    var today = new Date();
    var byDay = {};
    var cutoff = new Date(today.getTime() - days * 86400000);
    TransactionService.list().forEach(function (r) {
      if (r.Type !== TYPES.EXPENSE || !r.Date) return;
      var dt = new Date(r.Date);
      if (dt < cutoff) return;
      var k = fmtDate_(dt);
      byDay[k] = (byDay[k] || 0) + (parseFloat(r.Amount) || 0);
    });
    var out = [];
    for (var i = days - 1; i >= 0; i--) {
      var d = new Date(today.getTime() - i * 86400000);
      var k = fmtDate_(d);
      out.push({ date: d, amount: byDay[k] || 0 });
    }
    return out;
  }

  function previousMonthTotals_() {
    var d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - 1);
    return TransactionService.monthTotals(d);
  }

  function sixMonthSummary_() {
    var out = [];
    for (var i = 5; i >= 0; i--) {
      var d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - i);
      var t = TransactionService.monthTotals(d);
      out.push({
        month: Utilities.formatDate(d, APP.DEFAULT_TIMEZONE, 'MMM yyyy'),
        income: t.income, expense: t.expense
      });
    }
    return out;
  }

  /** ASCII bar of length n (block characters) for in-cell visualisation. */
  function bar_(ratio, n) {
    if (!isFinite(ratio) || ratio < 0) ratio = 0;
    if (ratio > 1) ratio = 1;
    var filled = Math.round(ratio * n);
    return new Array(filled + 1).join('█') + new Array(n - filled + 1).join('░');
  }

  function fmtPct_(n) {
    if (!isFinite(n)) return '0%';
    return (n * 100).toFixed(n < 0.1 ? 1 : 1) + '%';
  }

  return { refresh: refresh, refreshLightweight: refreshLightweight, layout: layout };
})();
