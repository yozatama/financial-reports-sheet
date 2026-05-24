/**
 * BudgetService.gs
 * -------------------------------------------------------------
 * Recomputes the Budget Tracking sheet from current-month
 * transactions. Status can be: On Track / Warning (>=80%) /
 * Over Budget (>=100%).
 * -------------------------------------------------------------
 */

var BudgetService = (function () {

  function add(b) {
    require_(b.category, 'Category required');
    require_(b.limit, 'Limit required');
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.BUDGETS);
    var row = [
      genId('BDG'),
      b.month || monthKey_(new Date()),
      b.category,
      parseFloat(b.limit) || 0,
      0, parseFloat(b.limit) || 0, 0, 'On Track'
    ];
    sh.appendRow(row);
    refreshAll();
    return row;
  }

  function remove(id) {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.BUDGETS);
    var rowIdx = findRowByValue_(sh, COLS.BUDGETS.ID, id);
    if (rowIdx < 0) throw new Error('Budget not found');
    sh.deleteRow(rowIdx);
    return { ok: true };
  }

  function list() {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.BUDGETS);
    return sh ? readObjects_(sh) : [];
  }

  /** Recompute spent / remaining / status for every budget row. */
  function refreshAll() {
    var ss = SpreadsheetApp.getActive();
    var bSh = ss.getSheetByName(SHEETS.BUDGETS);
    var txSh = ss.getSheetByName(SHEETS.TRANSACTIONS);
    if (!bSh || bSh.getLastRow() < 2) return;

    var rows = bSh.getRange(2, 1, bSh.getLastRow() - 1, 8).getValues();
    var spentByKey = {};

    if (txSh && txSh.getLastRow() >= 2) {
      var txs = txSh.getRange(2, 1, txSh.getLastRow() - 1, 16).getValues();
      txs.forEach(function (r) {
        if (r[COLS.TRANSACTIONS.TYPE - 1] !== TYPES.EXPENSE) return;
        var date = r[COLS.TRANSACTIONS.DATE - 1];
        if (!date) return;
        var month = monthKey_(new Date(date));
        var cat = r[COLS.TRANSACTIONS.CATEGORY - 1];
        var amt = parseFloat(r[COLS.TRANSACTIONS.AMOUNT - 1]) || 0;
        var key = month + '||' + cat;
        spentByKey[key] = (spentByKey[key] || 0) + amt;
      });
    }

    rows = rows.map(function (row) {
      var month = row[COLS.BUDGETS.MONTH - 1];
      var cat = row[COLS.BUDGETS.CATEGORY - 1];
      var limit = parseFloat(row[COLS.BUDGETS.LIMIT - 1]) || 0;
      var spent = spentByKey[month + '||' + cat] || 0;
      var remaining = limit - spent;
      var pct = limit > 0 ? spent / limit : 0;
      var status = pct >= 1 ? 'Over Budget' : (pct >= 0.8 ? 'Warning' : 'On Track');
      row[COLS.BUDGETS.SPENT - 1] = spent;
      row[COLS.BUDGETS.REMAINING - 1] = remaining;
      row[COLS.BUDGETS.PERCENT - 1] = pct;
      row[COLS.BUDGETS.STATUS - 1] = status;
      return row;
    });
    bSh.getRange(2, 1, rows.length, 8).setValues(rows);
  }

  return { add: add, remove: remove, list: list, refreshAll: refreshAll };
})();
