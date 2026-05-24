/**
 * AccountService.gs
 * -------------------------------------------------------------
 * CRUD for the Accounts sheet plus the central balance recompute
 * routine. Balance is always derived from initial balance plus
 * the sum of transactions, never edited directly. This avoids
 * "drift" between transactions and account totals.
 * -------------------------------------------------------------
 */

var AccountService = (function () {

  function add(acc) {
    require_(acc, 'Account payload is required');
    require_(acc.name, 'Account name is required');
    require_(acc.type, 'Account type is required');

    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.ACCOUNTS);
    if (findRowByValue_(sh, COLS.ACCOUNTS.NAME, acc.name) > 0) {
      throw new Error('Account already exists: ' + acc.name);
    }
    var row = [
      acc.id || genId('ACC'),
      acc.name,
      acc.type,
      acc.institution || '',
      parseFloat(acc.initialBalance) || 0,
      parseFloat(acc.initialBalance) || 0,
      acc.currency || APP.DEFAULT_CURRENCY,
      new Date(),
      acc.status || STATUS.ACTIVE,
      acc.notes || ''
    ];
    sh.appendRow(row);
    recomputeAllBalances();
    Logger_.info('Account added', { name: acc.name });
    return row;
  }

  function update(id, patch) {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.ACCOUNTS);
    var rowIdx = findRowByValue_(sh, COLS.ACCOUNTS.ID, id);
    if (rowIdx < 0) throw new Error('Account not found: ' + id);
    var current = sh.getRange(rowIdx, 1, 1, 10).getValues()[0];
    if (patch.name) current[COLS.ACCOUNTS.NAME - 1] = patch.name;
    if (patch.type) current[COLS.ACCOUNTS.TYPE - 1] = patch.type;
    if (patch.institution !== undefined) current[COLS.ACCOUNTS.INSTITUTION - 1] = patch.institution;
    if (patch.initialBalance !== undefined) current[COLS.ACCOUNTS.INITIAL_BALANCE - 1] = parseFloat(patch.initialBalance) || 0;
    if (patch.currency) current[COLS.ACCOUNTS.CURRENCY - 1] = patch.currency;
    if (patch.status) current[COLS.ACCOUNTS.STATUS - 1] = patch.status;
    if (patch.notes !== undefined) current[COLS.ACCOUNTS.NOTES - 1] = patch.notes;
    current[COLS.ACCOUNTS.LAST_UPDATED - 1] = new Date();
    sh.getRange(rowIdx, 1, 1, 10).setValues([current]);
    recomputeAllBalances();
    return current;
  }

  function remove(id) {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.ACCOUNTS);
    var rowIdx = findRowByValue_(sh, COLS.ACCOUNTS.ID, id);
    if (rowIdx < 0) throw new Error('Account not found: ' + id);
    sh.deleteRow(rowIdx);
    return { ok: true };
  }

  function list(activeOnly) {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.ACCOUNTS);
    if (!sh) return [];
    var rows = readObjects_(sh);
    return activeOnly ? rows.filter(function (r) { return r.Status === STATUS.ACTIVE; }) : rows;
  }

  function names(activeOnly) {
    return list(activeOnly).map(function (r) { return r['Account Name']; });
  }

  function totalBalance() {
    return list(true).reduce(function (sum, r) {
      return sum + (parseFloat(r['Current Balance']) || 0);
    }, 0);
  }

  /**
   * Recompute current balance for every account.
   * balance = initial + income(in) - expense(out) + transfers(in - out)
   * Single batched read/write for performance on big sheets.
   */
  function recomputeAllBalances() {
    var ss = SpreadsheetApp.getActive();
    var accSh = ss.getSheetByName(SHEETS.ACCOUNTS);
    var txSh = ss.getSheetByName(SHEETS.TRANSACTIONS);
    if (!accSh) return;

    var accLast = accSh.getLastRow();
    if (accLast < 2) return;

    var accs = accSh.getRange(2, 1, accLast - 1, 10).getValues();
    var balances = {};
    accs.forEach(function (r) {
      balances[r[COLS.ACCOUNTS.NAME - 1]] = parseFloat(r[COLS.ACCOUNTS.INITIAL_BALANCE - 1]) || 0;
    });

    if (txSh && txSh.getLastRow() >= 2) {
      var txs = txSh.getRange(2, 1, txSh.getLastRow() - 1, 16).getValues();
      txs.forEach(function (r) {
        var type = r[COLS.TRANSACTIONS.TYPE - 1];
        var amount = parseFloat(r[COLS.TRANSACTIONS.AMOUNT - 1]) || 0;
        var src = r[COLS.TRANSACTIONS.ACCOUNT - 1];
        var dst = r[COLS.TRANSACTIONS.ACCOUNT_TO - 1];
        if (!type || amount <= 0) return;
        if (type === TYPES.INCOME && src in balances) balances[src] += amount;
        else if (type === TYPES.EXPENSE && src in balances) balances[src] -= amount;
        else if (type === TYPES.TRANSFER) {
          if (src in balances) balances[src] -= amount;
          if (dst in balances) balances[dst] += amount;
        }
      });
    }

    var now = new Date();
    var updates = accs.map(function (r) {
      var name = r[COLS.ACCOUNTS.NAME - 1];
      r[COLS.ACCOUNTS.CURRENT_BALANCE - 1] = balances[name] !== undefined ? balances[name] : r[COLS.ACCOUNTS.CURRENT_BALANCE - 1];
      r[COLS.ACCOUNTS.LAST_UPDATED - 1] = now;
      return r;
    });
    accSh.getRange(2, 1, updates.length, 10).setValues(updates);
  }

  function recomputeAllBalances_safe() {
    try { recomputeAllBalances(); } catch (e) { Logger_.error('recompute failed', e); }
  }

  return {
    add: add, update: update, remove: remove,
    list: list, names: names, totalBalance: totalBalance,
    recomputeAllBalances: recomputeAllBalances,
    recomputeAllBalances_safe: recomputeAllBalances_safe
  };
})();
