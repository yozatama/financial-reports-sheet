/**
 * Triggers.gs
 * -------------------------------------------------------------
 * Time-based automations:
 *  - dailyMaintenance:     refresh balances, budgets, dashboard, send due reminders
 *  - monthlyRollover:      append previous month summary, copy recurring tx, reset budgets
 *  - applyRecurring:       create transactions for items flagged Recurring=Yes
 *
 * Use installTriggers() once after deployment. uninstallTriggers()
 * removes everything we created.
 * -------------------------------------------------------------
 */

function installTriggers() {
  uninstallTriggers();
  ScriptApp.newTrigger('dailyMaintenance')
    .timeBased().atHour(7).everyDays(1).create();
  ScriptApp.newTrigger('monthlyRollover')
    .timeBased().onMonthDay(1).atHour(1).create();
  Logger_.info('Triggers installed');
  return { ok: true };
}

function uninstallTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'dailyMaintenance' || fn === 'monthlyRollover') {
      ScriptApp.deleteTrigger(t);
    }
  });
  return { ok: true };
}

/** Daily housekeeping. */
function dailyMaintenance() {
  try {
    AccountService.recomputeAllBalances();
    BudgetService.refreshAll();
    applyRecurring();
    DashboardService.refresh();
    sendDueDateReminders_();
  } catch (err) {
    Logger_.error('dailyMaintenance failed', err);
  }
}

/** Append last-month rollup to Monthly Summary, then refresh dashboard. */
function monthlyRollover() {
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(SHEETS.MONTHLY);
    if (!sh) return;
    var prev = new Date(); prev.setDate(1); prev.setMonth(prev.getMonth() - 1);
    var totals = TransactionService.monthTotals(prev);
    var topCat = topCategoryFor_(prev);
    var savingsRate = totals.income > 0 ? (totals.income - totals.expense) / totals.income : 0;
    sh.appendRow([
      Utilities.formatDate(prev, APP.DEFAULT_TIMEZONE, 'yyyy-MM'),
      totals.income, totals.expense, totals.income - totals.expense,
      savingsRate, topCat || '', totals.count
    ]);
    DashboardService.refresh();
  } catch (err) {
    Logger_.error('monthlyRollover failed', err);
  }
}

function topCategoryFor_(monthDate) {
  var from = startOfMonth_(monthDate), to = endOfMonth_(monthDate);
  var map = {};
  TransactionService.list().forEach(function (r) {
    if (r.Type !== TYPES.EXPENSE || !r.Date) return;
    var dt = new Date(r.Date);
    if (dt < from || dt > to) return;
    var k = r.Category || 'Other';
    map[k] = (map[k] || 0) + (parseFloat(r.Amount) || 0);
  });
  var top = Object.keys(map).sort(function (a, b) { return map[b] - map[a]; })[0];
  return top || '';
}

/**
 * Recurring transactions: any Transactions row flagged Yes is
 * cloned for the current day if no clone exists yet for today.
 * Lightweight implementation - good enough for personal use.
 */
function applyRecurring() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.TRANSACTIONS);
  if (!sh || sh.getLastRow() < 2) return;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 16).getValues();
  var today = fmtDate_(new Date());
  // Build a set of (description+amount) added today to avoid double-applying
  var todaySet = {};
  rows.forEach(function (r) {
    if (fmtDate_(r[COLS.TRANSACTIONS.DATE - 1]) === today) {
      var key = r[COLS.TRANSACTIONS.DESCRIPTION - 1] + '|' + r[COLS.TRANSACTIONS.AMOUNT - 1];
      todaySet[key] = true;
    }
  });

  rows.forEach(function (r) {
    var recurring = String(r[COLS.TRANSACTIONS.RECURRING - 1]).toLowerCase() === 'yes';
    if (!recurring) return;
    var desc = r[COLS.TRANSACTIONS.DESCRIPTION - 1];
    var amt = r[COLS.TRANSACTIONS.AMOUNT - 1];
    var key = desc + '|' + amt;
    if (todaySet[key]) return;
    var origDate = r[COLS.TRANSACTIONS.DATE - 1];
    if (!origDate) return;
    var od = new Date(origDate);
    if (od.getDate() !== new Date().getDate()) return; // only on matching day-of-month
    TransactionService.add({
      type: r[COLS.TRANSACTIONS.TYPE - 1],
      category: r[COLS.TRANSACTIONS.CATEGORY - 1],
      subcategory: r[COLS.TRANSACTIONS.SUBCATEGORY - 1],
      amount: amt,
      account: r[COLS.TRANSACTIONS.ACCOUNT - 1],
      accountTo: r[COLS.TRANSACTIONS.ACCOUNT_TO - 1],
      merchant: r[COLS.TRANSACTIONS.MERCHANT - 1],
      description: desc + ' (recurring)',
      tags: 'recurring',
      recurring: false
    });
  });
}

function sendDueDateReminders_() {
  if (String(getSetting('email_reminders', 'No')).toLowerCase() !== 'yes') return;
  var to = String(getSetting('notify_email', '')).trim();
  if (!to) return;
  var dues = DebtService.upcomingDueDates(7);
  if (!dues.length) return;
  var rows = dues.map(function (d) {
    return '<tr><td>' + htmlEscape_(d.name) + '</td><td>' + fmtDate_(d.dueDate) +
      '</td><td>' + d.daysUntil + ' days</td><td style="text-align:right">' +
      fmtMoney_(d.minPayment) + '</td></tr>';
  }).join('');
  var html =
    '<h3>Upcoming Debt Due Dates</h3>' +
    '<table cellpadding="6" cellspacing="0" border="1" style="border-collapse:collapse;font-family:Arial">' +
    '<tr style="background:#1F2937;color:#fff"><th>Debt</th><th>Due</th><th>In</th><th>Min Pay</th></tr>' +
    rows + '</table>';
  MailApp.sendEmail({
    to: to,
    subject: '[' + APP.NAME + '] Upcoming debt due dates',
    htmlBody: html
  });
}
