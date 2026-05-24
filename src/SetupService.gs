/**
 * SetupService.gs
 * -------------------------------------------------------------
 * Builds (or repairs) every sheet the app needs. Idempotent:
 * running it twice will not duplicate data, only re-apply the
 * structure (headers, formats, validations, conditional rules).
 *
 * Public entry point: setupSpreadsheet()
 * -------------------------------------------------------------
 */

function setupSpreadsheet() {
  var ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(APP.DEFAULT_TIMEZONE);
  ss.setSpreadsheetLocale(APP.DEFAULT_LOCALE);

  setupSettings_(ss);
  setupCategories_(ss);
  setupAccounts_(ss);
  setupDebts_(ss);
  setupTransactions_(ss);
  setupBudgets_(ss);
  setupMonthlySummary_(ss);
  setupAILogs_(ss);
  setupDashboard_(ss);
  reorderSheets_(ss);

  // Refresh all derived data
  AccountService.recomputeAllBalances();
  BudgetService.refreshAll();
  DashboardService.refresh();

  return { ok: true, message: 'Setup complete' };
}

/* -------------------------------------------------------------
 * Per-sheet builders
 * ------------------------------------------------------------- */

function ensureSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function writeHeader_(sh, headers, opts) {
  opts = opts || {};
  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground(THEME.HEADER_BG)
    .setFontColor(THEME.HEADER_FG)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment(opts.headerAlign || 'left');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 36);
}

function applyZebra_(sh, startRow, startCol, numRows, numCols) {
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=ISEVEN(ROW())')
    .setBackground(THEME.ZEBRA)
    .setRanges([sh.getRange(startRow, startCol, numRows, numCols)])
    .build();
  var rules = sh.getConditionalFormatRules();
  rules.push(rule);
  sh.setConditionalFormatRules(rules);
}

/* ---------------- Settings ---------------- */
function setupSettings_(ss) {
  var sh = ensureSheet_(ss, SHEETS.SETTINGS);
  if (sh.getLastRow() === 0) {
    writeHeader_(sh, ['Key', 'Value', 'Description']);
    var defaults = [
      ['currency', APP.DEFAULT_CURRENCY, 'Default currency code (IDR, USD, ...)'],
      ['locale', APP.DEFAULT_LOCALE, 'Spreadsheet locale'],
      ['timezone', APP.DEFAULT_TIMEZONE, 'Spreadsheet timezone'],
      ['monthly_income_target', 10000000, 'Used for savings rate insight'],
      ['openai_api_key', '', 'Optional. Paste your OpenAI key to enable AI fallback parser'],
      ['openai_model', 'gpt-4o-mini', 'OpenAI model used by the AI fallback'],
      ['email_reminders', 'No', 'Send debt due-date reminders by email (Yes/No)'],
      ['notify_email', Session.getEffectiveUser().getEmail() || '', 'Email address for reminders'],
      ['theme', 'light', 'UI theme: light or dark']
    ];
    sh.getRange(2, 1, defaults.length, 3).setValues(defaults);
  }
  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 320);
  sh.setColumnWidth(3, 420);
}

/* ---------------- Categories ---------------- */
function setupCategories_(ss) {
  var sh = ensureSheet_(ss, SHEETS.CATEGORIES);
  if (sh.getLastRow() === 0) {
    writeHeader_(sh, ['ID', 'Type', 'Category', 'Subcategory', 'Icon', 'Color', 'Active']);
    var seed = SEED_CATEGORIES_();
    sh.getRange(2, 1, seed.length, 7).setValues(seed);
  }
  // Validation: Type column
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([TYPES.INCOME, TYPES.EXPENSE, TYPES.TRANSFER], true).build();
  sh.getRange(2, COLS.CATEGORIES.TYPE, 2000).setDataValidation(typeRule);
  var activeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true).build();
  sh.getRange(2, COLS.CATEGORIES.ACTIVE, 2000).setDataValidation(activeRule);

  [120, 110, 180, 220, 60, 110, 80].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
}

function SEED_CATEGORIES_() {
  // [ID, Type, Category, Subcategory, Icon, Color, Active]
  var rows = [];
  function add(type, cat, sub, icon, color) {
    rows.push([genId('CAT'), type, cat, sub, icon, color, 'Yes']);
  }
  add(TYPES.EXPENSE, 'Food & Beverage', 'Coffee', '☕', '#F59E0B');
  add(TYPES.EXPENSE, 'Food & Beverage', 'Restaurant', '🍽️', '#F59E0B');
  add(TYPES.EXPENSE, 'Food & Beverage', 'Groceries', '🛒', '#F59E0B');
  add(TYPES.EXPENSE, 'Food & Beverage', 'Snacks', '🍪', '#F59E0B');
  add(TYPES.EXPENSE, 'Transport', 'Fuel', '⛽', '#3B82F6');
  add(TYPES.EXPENSE, 'Transport', 'Ride-hailing', '🚕', '#3B82F6');
  add(TYPES.EXPENSE, 'Transport', 'Public Transport', '🚆', '#3B82F6');
  add(TYPES.EXPENSE, 'Transport', 'Parking & Toll', '🅿️', '#3B82F6');
  add(TYPES.EXPENSE, 'Bills & Utilities', 'Electricity', '💡', '#10B981');
  add(TYPES.EXPENSE, 'Bills & Utilities', 'Water', '💧', '#10B981');
  add(TYPES.EXPENSE, 'Bills & Utilities', 'Internet', '🌐', '#10B981');
  add(TYPES.EXPENSE, 'Bills & Utilities', 'Phone', '📱', '#10B981');
  add(TYPES.EXPENSE, 'Shopping', 'Clothes', '👕', '#EC4899');
  add(TYPES.EXPENSE, 'Shopping', 'Electronics', '🔌', '#EC4899');
  add(TYPES.EXPENSE, 'Shopping', 'Online Shopping', '📦', '#EC4899');
  add(TYPES.EXPENSE, 'Health', 'Medical', '🏥', '#EF4444');
  add(TYPES.EXPENSE, 'Health', 'Pharmacy', '💊', '#EF4444');
  add(TYPES.EXPENSE, 'Entertainment', 'Movies', '🎬', '#A855F7');
  add(TYPES.EXPENSE, 'Entertainment', 'Subscriptions', '📺', '#A855F7');
  add(TYPES.EXPENSE, 'Entertainment', 'Games', '🎮', '#A855F7');
  add(TYPES.EXPENSE, 'Education', 'Courses', '🎓', '#0EA5E9');
  add(TYPES.EXPENSE, 'Education', 'Books', '📚', '#0EA5E9');
  add(TYPES.EXPENSE, 'Family', 'Kids', '👶', '#F472B6');
  add(TYPES.EXPENSE, 'Family', 'Parents', '👨‍👩‍👧', '#F472B6');
  add(TYPES.EXPENSE, 'Debt Payment', 'Credit Card', '💳', '#6B7280');
  add(TYPES.EXPENSE, 'Debt Payment', 'PayLater', '🧾', '#6B7280');
  add(TYPES.EXPENSE, 'Debt Payment', 'Loan', '🏦', '#6B7280');
  add(TYPES.EXPENSE, 'Other', 'Miscellaneous', '🔖', '#9CA3AF');

  add(TYPES.INCOME, 'Salary', 'Monthly Salary', '💼', '#22D3A6');
  add(TYPES.INCOME, 'Salary', 'Bonus', '🎁', '#22D3A6');
  add(TYPES.INCOME, 'Freelance', 'Project', '🧑‍💻', '#22D3A6');
  add(TYPES.INCOME, 'Investment', 'Dividend', '📈', '#22D3A6');
  add(TYPES.INCOME, 'Investment', 'Interest', '🏦', '#22D3A6');
  add(TYPES.INCOME, 'Other Income', 'Refund', '↩️', '#22D3A6');
  add(TYPES.INCOME, 'Other Income', 'Gift', '🎀', '#22D3A6');

  add(TYPES.TRANSFER, 'Transfer', 'Between Accounts', '🔁', '#5B5BD6');
  add(TYPES.TRANSFER, 'Transfer', 'Top Up', '⬆️', '#5B5BD6');
  add(TYPES.TRANSFER, 'Transfer', 'Withdraw', '⬇️', '#5B5BD6');
  return rows;
}

/* ---------------- Accounts ---------------- */
function setupAccounts_(ss) {
  var sh = ensureSheet_(ss, SHEETS.ACCOUNTS);
  if (sh.getLastRow() === 0) {
    writeHeader_(sh, [
      'Account ID', 'Account Name', 'Account Type', 'Institution',
      'Initial Balance', 'Current Balance', 'Currency',
      'Last Updated', 'Status', 'Notes'
    ]);
    var seed = [
      [genId('ACC'), 'Cash Wallet', 'Cash', 'Self', 500000, 500000, 'IDR', new Date(), 'Active', 'Physical cash'],
      [genId('ACC'), 'BCA', 'Bank', 'BCA', 5000000, 5000000, 'IDR', new Date(), 'Active', 'Primary bank'],
      [genId('ACC'), 'Mandiri', 'Bank', 'Mandiri', 2000000, 2000000, 'IDR', new Date(), 'Active', ''],
      [genId('ACC'), 'SeaBank', 'Digital Bank', 'SeaBank', 1500000, 1500000, 'IDR', new Date(), 'Active', ''],
      [genId('ACC'), 'GoPay', 'E-Wallet', 'Gojek', 250000, 250000, 'IDR', new Date(), 'Active', ''],
      [genId('ACC'), 'OVO', 'E-Wallet', 'OVO', 150000, 150000, 'IDR', new Date(), 'Active', ''],
      [genId('ACC'), 'DANA', 'E-Wallet', 'DANA', 100000, 100000, 'IDR', new Date(), 'Active', ''],
      [genId('ACC'), 'ShopeePay', 'E-Wallet', 'Shopee', 80000, 80000, 'IDR', new Date(), 'Active', '']
    ];
    sh.getRange(2, 1, seed.length, 10).setValues(seed);
  }

  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(ACCOUNT_TYPES, true).build();
  sh.getRange(2, COLS.ACCOUNTS.TYPE, 2000).setDataValidation(typeRule);
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([STATUS.ACTIVE, STATUS.INACTIVE], true).build();
  sh.getRange(2, COLS.ACCOUNTS.STATUS, 2000).setDataValidation(statusRule);

  sh.getRange(2, COLS.ACCOUNTS.INITIAL_BALANCE, 2000, 2).setNumberFormat('#,##0');
  sh.getRange(2, COLS.ACCOUNTS.LAST_UPDATED, 2000).setNumberFormat('yyyy-mm-dd hh:mm');

  // Highlight low balance
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#FEE2E2').setFontColor('#991B1B')
    .setRanges([sh.getRange(2, COLS.ACCOUNTS.CURRENT_BALANCE, 2000)])
    .build();
  sh.setConditionalFormatRules((sh.getConditionalFormatRules() || []).concat([rule]));

  [120, 180, 130, 140, 130, 140, 90, 160, 100, 240]
    .forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
}

/* ---------------- Debts ---------------- */
function setupDebts_(ss) {
  var sh = ensureSheet_(ss, SHEETS.DEBTS);
  if (sh.getLastRow() === 0) {
    writeHeader_(sh, [
      'Debt ID', 'Debt Name', 'Provider', 'Type', 'Outstanding Balance',
      'Credit Limit', 'Min Payment', 'Interest %', 'Due Day',
      'Billing Day', 'Monthly Installment', 'Tenor', 'Remaining Tenor',
      'Status', 'Notes', 'Updated At'
    ]);
    var seed = [
      [genId('DEB'), 'BCA Credit Card', 'BCA', 'Credit Card', 4500000, 15000000, 450000, 2.25, 18, 5, 0, 0, 0, 'Active', 'Primary CC', new Date()],
      [genId('DEB'), 'ShopeePay PayLater', 'Shopee', 'PayLater', 800000, 2500000, 80000, 2.95, 25, 1, 0, 0, 0, 'Active', '', new Date()]
    ];
    sh.getRange(2, 1, seed.length, 16).setValues(seed);
  }
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DEBT_TYPES, true).build();
  sh.getRange(2, COLS.DEBTS.TYPE, 2000).setDataValidation(typeRule);
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([STATUS.ACTIVE, STATUS.PAID_OFF, STATUS.OVERDUE], true).build();
  sh.getRange(2, COLS.DEBTS.STATUS, 2000).setDataValidation(statusRule);

  sh.getRange(2, COLS.DEBTS.OUTSTANDING, 2000, 2).setNumberFormat('#,##0');
  sh.getRange(2, COLS.DEBTS.MIN_PAYMENT, 2000).setNumberFormat('#,##0');
  sh.getRange(2, COLS.DEBTS.INSTALLMENT, 2000).setNumberFormat('#,##0');
  sh.getRange(2, COLS.DEBTS.INTEREST, 2000).setNumberFormat('0.00"%"');
  sh.getRange(2, COLS.DEBTS.UPDATED_AT, 2000).setNumberFormat('yyyy-mm-dd hh:mm');

  // Conditional formatting: utilization > 30% red, > 50% bright red, paid off green
  var utilRange = sh.getRange(2, COLS.DEBTS.OUTSTANDING, 2000);
  var rules = sh.getConditionalFormatRules() || [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($F2>0,$E2/$F2>0.5)')
    .setBackground('#FEE2E2').setFontColor('#991B1B')
    .setRanges([utilRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($F2>0,$E2/$F2>0.3,$E2/$F2<=0.5)')
    .setBackground('#FEF3C7').setFontColor('#92400E')
    .setRanges([utilRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(STATUS.PAID_OFF)
    .setBackground('#D1FAE5').setFontColor('#065F46')
    .setRanges([sh.getRange(2, COLS.DEBTS.STATUS, 2000)]).build());
  sh.setConditionalFormatRules(rules);

  [120, 200, 130, 130, 160, 140, 130, 100, 90, 110, 160, 90, 140, 110, 220, 160]
    .forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
}

/* ---------------- Transactions ---------------- */
function setupTransactions_(ss) {
  var sh = ensureSheet_(ss, SHEETS.TRANSACTIONS);
  if (sh.getLastRow() === 0) {
    writeHeader_(sh, [
      'Transaction ID', 'Date', 'Time', 'Type', 'Category', 'Subcategory',
      'Amount', 'Account', 'Account To', 'Merchant', 'Description', 'Tags',
      'Recurring', 'Created At', 'Raw Input', 'Confidence'
    ]);
  }

  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([TYPES.INCOME, TYPES.EXPENSE, TYPES.TRANSFER], true).build();
  sh.getRange(2, COLS.TRANSACTIONS.TYPE, 5000).setDataValidation(typeRule);
  var recurRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true).build();
  sh.getRange(2, COLS.TRANSACTIONS.RECURRING, 5000).setDataValidation(recurRule);

  // Account dropdown sourced from Accounts sheet
  var accSh = ss.getSheetByName(SHEETS.ACCOUNTS);
  if (accSh) {
    var accRange = accSh.getRange('B2:B1000');
    var accRule = SpreadsheetApp.newDataValidation().requireValueInRange(accRange, true).build();
    sh.getRange(2, COLS.TRANSACTIONS.ACCOUNT, 5000).setDataValidation(accRule);
    sh.getRange(2, COLS.TRANSACTIONS.ACCOUNT_TO, 5000).setDataValidation(accRule);
  }
  // Category dropdown
  var catSh = ss.getSheetByName(SHEETS.CATEGORIES);
  if (catSh) {
    var catRange = catSh.getRange('C2:C1000');
    var catRule = SpreadsheetApp.newDataValidation().requireValueInRange(catRange, true).build();
    sh.getRange(2, COLS.TRANSACTIONS.CATEGORY, 5000).setDataValidation(catRule);
  }

  sh.getRange(2, COLS.TRANSACTIONS.DATE, 5000).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, COLS.TRANSACTIONS.TIME, 5000).setNumberFormat('hh:mm');
  sh.getRange(2, COLS.TRANSACTIONS.AMOUNT, 5000).setNumberFormat('#,##0');
  sh.getRange(2, COLS.TRANSACTIONS.CREATED_AT, 5000).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sh.getRange(2, COLS.TRANSACTIONS.CONFIDENCE, 5000).setNumberFormat('0%');

  // Color rows by type
  var rules = sh.getConditionalFormatRules() || [];
  var fullRange = sh.getRange(2, 1, 5000, 16);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$D2="' + TYPES.INCOME + '"')
    .setBackground('#ECFDF5')
    .setRanges([fullRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$D2="' + TYPES.EXPENSE + '"')
    .setBackground('#FEF2F2')
    .setRanges([fullRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$D2="' + TYPES.TRANSFER + '"')
    .setBackground('#EEF2FF')
    .setRanges([fullRange]).build());
  sh.setConditionalFormatRules(rules);

  [150, 100, 80, 100, 160, 160, 130, 150, 150, 180, 280, 180, 100, 160, 280, 100]
    .forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
}

/* ---------------- Budgets ---------------- */
function setupBudgets_(ss) {
  var sh = ensureSheet_(ss, SHEETS.BUDGETS);
  if (sh.getLastRow() === 0) {
    writeHeader_(sh, ['ID', 'Month', 'Category', 'Limit', 'Spent', 'Remaining', '% Used', 'Status']);
    var month = monthKey_(new Date());
    var seed = [
      [genId('BDG'), month, 'Food & Beverage', 1500000, 0, 1500000, 0, 'On Track'],
      [genId('BDG'), month, 'Transport', 800000, 0, 800000, 0, 'On Track'],
      [genId('BDG'), month, 'Bills & Utilities', 1200000, 0, 1200000, 0, 'On Track'],
      [genId('BDG'), month, 'Shopping', 1000000, 0, 1000000, 0, 'On Track'],
      [genId('BDG'), month, 'Entertainment', 500000, 0, 500000, 0, 'On Track']
    ];
    sh.getRange(2, 1, seed.length, 8).setValues(seed);
  }
  sh.getRange(2, COLS.BUDGETS.LIMIT, 2000, 3).setNumberFormat('#,##0');
  sh.getRange(2, COLS.BUDGETS.PERCENT, 2000).setNumberFormat('0%');

  var rules = sh.getConditionalFormatRules() || [];
  var statusRange = sh.getRange(2, COLS.BUDGETS.STATUS, 2000);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Over Budget')
    .setBackground('#FEE2E2').setFontColor('#991B1B').setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Warning')
    .setBackground('#FEF3C7').setFontColor('#92400E').setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('On Track')
    .setBackground('#D1FAE5').setFontColor('#065F46').setRanges([statusRange]).build());
  sh.setConditionalFormatRules(rules);

  [120, 100, 200, 140, 140, 140, 100, 130]
    .forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
}

/* ---------------- Monthly summary ---------------- */
function setupMonthlySummary_(ss) {
  var sh = ensureSheet_(ss, SHEETS.MONTHLY);
  if (sh.getLastRow() === 0) {
    writeHeader_(sh, ['Month', 'Income', 'Expense', 'Net', 'Savings Rate', 'Top Category', 'Transactions']);
  }
  sh.getRange(2, 2, 200, 3).setNumberFormat('#,##0');
  sh.getRange(2, 5, 200).setNumberFormat('0.0%');
  [110, 160, 160, 160, 130, 220, 130].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
}

/* ---------------- AI logs ---------------- */
function setupAILogs_(ss) {
  var sh = ensureSheet_(ss, SHEETS.AI_LOGS);
  if (sh.getLastRow() === 0) {
    writeHeader_(sh, ['Timestamp', 'Raw Input', 'Method', 'Confidence', 'Parsed Type',
      'Parsed Category', 'Parsed Amount', 'Parsed Account', 'Status', 'Result JSON']);
  }
  sh.getRange(2, 1, 5000).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sh.getRange(2, 4, 5000).setNumberFormat('0%');
  sh.getRange(2, 7, 5000).setNumberFormat('#,##0');
  [160, 320, 100, 100, 110, 200, 130, 150, 110, 360]
    .forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
}

/* ---------------- Dashboard placeholder ---------------- */
function setupDashboard_(ss) {
  var sh = ensureSheet_(ss, SHEETS.DASHBOARD);
  sh.clear();
  sh.setHiddenGridlines(true);
  for (var c = 1; c <= 12; c++) sh.setColumnWidth(c, 110);
  // Layout will be drawn by DashboardService.refresh()
  DashboardService.layout(sh);
}

/* ---------------- Order ---------------- */
function reorderSheets_(ss) {
  var order = [
    SHEETS.DASHBOARD, SHEETS.TRANSACTIONS, SHEETS.ACCOUNTS, SHEETS.DEBTS,
    SHEETS.BUDGETS, SHEETS.MONTHLY, SHEETS.CATEGORIES, SHEETS.SETTINGS, SHEETS.AI_LOGS
  ];
  for (var i = 0; i < order.length; i++) {
    var sh = ss.getSheetByName(order[i]);
    if (!sh) continue;
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
  }
  var dash = ss.getSheetByName(SHEETS.DASHBOARD);
  if (dash) ss.setActiveSheet(dash);
}
