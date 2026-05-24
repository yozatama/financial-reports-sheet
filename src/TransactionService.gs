/**
 * TransactionService.gs
 * -------------------------------------------------------------
 * CRUD for the Transactions sheet. Adding or editing a
 * transaction automatically updates affected account balances
 * via AccountService.recomputeAllBalances() so the books always
 * stay consistent.
 * -------------------------------------------------------------
 */

var TransactionService = (function () {

  /**
   * Add a new transaction. `tx` is an object with keys matching
   * the column names. Required: type, amount, account.
   * Returns the persisted row as an object.
   */
  function add(tx) {
    require_(tx, 'Transaction payload is required');
    require_(tx.type, 'Transaction type is required');
    require_(tx.amount, 'Amount is required');
    require_(tx.account, 'Account is required');

    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(SHEETS.TRANSACTIONS);
    if (!sh) throw new Error('Transactions sheet not found. Run Setup first.');

    var now = new Date();
    var date = tx.date ? new Date(tx.date) : now;
    var time = tx.time ? new Date(tx.time) : now;

    var amount = Math.abs(parseFloat(tx.amount) || 0);
    if (amount <= 0) throw new Error('Amount must be greater than 0');

    if (tx.type === TYPES.TRANSFER) {
      require_(tx.accountTo, 'Destination account is required for transfers');
      if (tx.accountTo === tx.account) throw new Error('Source and destination must differ');
    }

    var row = [
      tx.id || genId('TX'),
      date,
      time,
      tx.type,
      tx.category || '',
      tx.subcategory || '',
      amount,
      tx.account,
      tx.accountTo || '',
      tx.merchant || '',
      tx.description || '',
      Array.isArray(tx.tags) ? tx.tags.join(', ') : (tx.tags || ''),
      tx.recurring ? 'Yes' : 'No',
      now,
      tx.rawInput || '',
      typeof tx.confidence === 'number' ? tx.confidence : ''
    ];

    sh.appendRow(row);

    // Side-effects: recompute balances and refresh budgets/dashboard
    AccountService.recomputeAllBalances();
    BudgetService.refreshAll();
    DashboardService.refreshLightweight();

    Logger_.info('Transaction added', { id: row[0], type: tx.type, amount: amount });
    return rowToObject_(sh, row);
  }

  /** Update existing transaction by ID. Partial payload allowed. */
  function update(id, patch) {
    require_(id, 'Transaction ID is required');
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.TRANSACTIONS);
    var rowIdx = findRowByValue_(sh, COLS.TRANSACTIONS.ID, id);
    if (rowIdx < 0) throw new Error('Transaction not found: ' + id);

    var width = sh.getLastColumn();
    var current = sh.getRange(rowIdx, 1, 1, width).getValues()[0];
    var headers = sh.getRange(1, 1, 1, width).getValues()[0];
    var headerMap = {};
    headers.forEach(function (h, i) { headerMap[h] = i; });

    function setCol(name, value) {
      if (name in headerMap && value !== undefined) current[headerMap[name]] = value;
    }
    setCol('Date', patch.date ? new Date(patch.date) : current[headerMap['Date']]);
    setCol('Time', patch.time ? new Date(patch.time) : current[headerMap['Time']]);
    setCol('Type', patch.type);
    setCol('Category', patch.category);
    setCol('Subcategory', patch.subcategory);
    if (patch.amount !== undefined) setCol('Amount', Math.abs(parseFloat(patch.amount) || 0));
    setCol('Account', patch.account);
    setCol('Account To', patch.accountTo);
    setCol('Merchant', patch.merchant);
    setCol('Description', patch.description);
    if (patch.tags !== undefined) {
      setCol('Tags', Array.isArray(patch.tags) ? patch.tags.join(', ') : patch.tags);
    }
    if (patch.recurring !== undefined) setCol('Recurring', patch.recurring ? 'Yes' : 'No');

    sh.getRange(rowIdx, 1, 1, width).setValues([current]);

    AccountService.recomputeAllBalances();
    BudgetService.refreshAll();
    DashboardService.refreshLightweight();

    Logger_.info('Transaction updated', { id: id });
    return rowToObject_(sh, current);
  }

  /** Delete a transaction by ID. */
  function remove(id) {
    require_(id, 'Transaction ID is required');
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.TRANSACTIONS);
    var rowIdx = findRowByValue_(sh, COLS.TRANSACTIONS.ID, id);
    if (rowIdx < 0) throw new Error('Transaction not found: ' + id);
    sh.deleteRow(rowIdx);

    AccountService.recomputeAllBalances();
    BudgetService.refreshAll();
    DashboardService.refreshLightweight();

    Logger_.info('Transaction deleted', { id: id });
    return { ok: true };
  }

  /** List transactions, newest first. Optional filter object. */
  function list(filter) {
    filter = filter || {};
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.TRANSACTIONS);
    if (!sh || sh.getLastRow() < 2) return [];
    var rows = readObjects_(sh).reverse();
    return rows.filter(function (r) {
      if (filter.type && r.Type !== filter.type) return false;
      if (filter.account && r.Account !== filter.account && r['Account To'] !== filter.account) return false;
      if (filter.category && r.Category !== filter.category) return false;
      if (filter.fromDate && r.Date && new Date(r.Date) < new Date(filter.fromDate)) return false;
      if (filter.toDate && r.Date && new Date(r.Date) > new Date(filter.toDate)) return false;
      return true;
    });
  }

  function getById(id) {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.TRANSACTIONS);
    var rowIdx = findRowByValue_(sh, COLS.TRANSACTIONS.ID, id);
    if (rowIdx < 0) return null;
    var row = sh.getRange(rowIdx, 1, 1, sh.getLastColumn()).getValues()[0];
    return rowToObject_(sh, row);
  }

  /** Sum amounts for the current month, grouped by type. */
  function monthTotals(monthDate) {
    var d = monthDate ? new Date(monthDate) : new Date();
    var from = startOfMonth_(d), to = endOfMonth_(d);
    var totals = { income: 0, expense: 0, transfer: 0, count: 0 };
    list().forEach(function (r) {
      if (!r.Date) return;
      var dt = new Date(r.Date);
      if (dt < from || dt > to) return;
      totals.count++;
      var amt = parseFloat(r.Amount) || 0;
      if (r.Type === TYPES.INCOME) totals.income += amt;
      else if (r.Type === TYPES.EXPENSE) totals.expense += amt;
      else if (r.Type === TYPES.TRANSFER) totals.transfer += amt;
    });
    return totals;
  }

  function rowToObject_(sh, row) {
    var headers = sh.getRange(1, 1, 1, row.length).getValues()[0];
    var o = {};
    headers.forEach(function (h, i) { o[h] = row[i]; });
    return o;
  }

  return {
    add: add, update: update, remove: remove,
    list: list, getById: getById, monthTotals: monthTotals
  };
})();
