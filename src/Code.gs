/**
 * Code.gs
 * -------------------------------------------------------------
 * Main entry point: custom menu, dialog/sidebar launchers, and
 * the server-side endpoints invoked from HTML via google.script.run.
 *
 * Every server endpoint is wrapped in safeCall_ so the UI can
 * always read a stable { ok, data | error } envelope.
 * -------------------------------------------------------------
 */

/* ---------------- Menu ---------------- */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(APP.NAME)
    .addItem('⚡ Quick Add (AI)', 'openQuickAdd')
    .addItem('＋ Add Transaction', 'openAddTransaction')
    .addSeparator()
    .addItem('🏦 Add Account', 'openAddAccount')
    .addItem('💳 Add Debt', 'openAddDebt')
    .addSeparator()
    .addItem('🔄 Refresh Dashboard', 'refreshDashboard')
    .addItem('📊 Open Sidebar', 'openSidebar')
    .addSeparator()
    .addSubMenu(ui.createMenu('Setup')
      .addItem('🛠 Run Initial Setup', 'runSetup')
      .addItem('🔁 Recompute Balances', 'runRecomputeBalances')
      .addItem('📅 Install Daily Triggers', 'runInstallTriggers')
      .addItem('🗑 Remove Triggers', 'runUninstallTriggers'))
    .addToUi();
}

/** Convenience wrappers so menu items map cleanly to services. */
function runSetup() {
  setupSpreadsheet();
  SpreadsheetApp.getActive().toast('Setup complete.', APP.NAME, 4);
}
function runRecomputeBalances() {
  AccountService.recomputeAllBalances();
  BudgetService.refreshAll();
  DashboardService.refreshLightweight();
  SpreadsheetApp.getActive().toast('Balances recomputed.', APP.NAME, 3);
}
function runInstallTriggers() {
  installTriggers();
  SpreadsheetApp.getActive().toast('Daily trigger installed (07:00).', APP.NAME, 4);
}
function runUninstallTriggers() {
  uninstallTriggers();
  SpreadsheetApp.getActive().toast('Triggers removed.', APP.NAME, 3);
}
function refreshDashboard() {
  DashboardService.refresh();
  SpreadsheetApp.getActive().toast('Dashboard refreshed.', APP.NAME, 2);
}

/* ---------------- Dialogs / Sidebar ---------------- */

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function openQuickAdd() {
  var html = HtmlService.createTemplateFromFile('QuickAdd').evaluate()
    .setTitle('Quick Add')
    .setWidth(560).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Quick Add (AI)');
}

function openAddTransaction() {
  var html = HtmlService.createTemplateFromFile('AddTransaction').evaluate()
    .setTitle('Add Transaction')
    .setWidth(560).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Transaction');
}

function openAddAccount() {
  var html = HtmlService.createTemplateFromFile('AddAccount').evaluate()
    .setTitle('Add Account')
    .setWidth(520).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Account');
}

function openAddDebt() {
  var html = HtmlService.createTemplateFromFile('AddDebt').evaluate()
    .setTitle('Add Debt')
    .setWidth(560).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Debt');
}

function openSidebar() {
  var html = HtmlService.createTemplateFromFile('Sidebar').evaluate()
    .setTitle(APP.NAME);
  SpreadsheetApp.getUi().showSidebar(html);
}

/* ---------------- Server endpoints (called from HTML) ---------------- */

/** Lookups consumed by every form to fill dropdowns. */
function uiBootstrap() {
  // Return raw data; google.script.run.withFailureHandler will catch any throws.
  return {
    accounts: AccountService.names(true),
    expenseCategories: CategoryService.topLevelNames(TYPES.EXPENSE),
    incomeCategories: CategoryService.topLevelNames(TYPES.INCOME),
    transferCategories: CategoryService.topLevelNames(TYPES.TRANSFER),
    currency: getSetting('currency', APP.DEFAULT_CURRENCY),
    hasOpenAI: OpenAIClient.isConfigured()
  };
}

function aiParse(text)   { return safeCall_(function () { return AIParser.parse(text); }); }
function aiCommit(payload) {
  return safeCall_(function () {
    return AIParser.commit(payload);
  });
}

function saveTransaction(tx) {
  return safeCall_(function () { return TransactionService.add(tx); });
}
function saveAccount(acc) {
  return safeCall_(function () { return AccountService.add(acc); });
}
function saveDebt(d) {
  return safeCall_(function () { return DebtService.add(d); });
}

/** Snapshot used by the sidebar UI. */
function sidebarSnapshot() {
  var totals = TransactionService.monthTotals(new Date());
  var net = totals.income - totals.expense;
  var recent = TransactionService.list().slice(0, 6).map(function (t) {
    var amount = (t.Type === TYPES.EXPENSE ? '-' : (t.Type === TYPES.INCOME ? '+' : '↔ ')) + fmtMoney_(t.Amount);
    return {
      title: (t.Category || '') + (t.Subcategory ? ' · ' + t.Subcategory : ''),
      subtitle: fmtDate_(t.Date) + ' · ' + (t.Account || '') + (t['Account To'] ? ' → ' + t['Account To'] : ''),
      amount: amount,
      kind: t.Type === TYPES.EXPENSE ? 'red' : (t.Type === TYPES.INCOME ? 'green' : 'gray')
    };
  });
  return {
    balance: AccountService.totalBalance(),
    balanceFmt: fmtMoney_(AccountService.totalBalance()),
    income: totals.income, incomeFmt: fmtMoney_(totals.income),
    expense: totals.expense, expenseFmt: fmtMoney_(totals.expense),
    net: net, netFmt: fmtMoney_(net),
    recent: recent
  };
}

/* ---------------- Smoke test (run from editor) ---------------- */
/**
 * Run this once after deploying to verify the parser end-to-end
 * without opening any UI. Logs results to View > Executions.
 */
function smokeTestParser() {
  var samples = [
    'jajan kopi 35rb',
    'bayar listrik 450k',
    'gajian 8jt',
    'isi bensin motor 50 ribu',
    'tf ke bca 1.2jt',
    'makan sushi 120rb pake gopay',
    'bayar cc bca 2jt',
    'topup ovo 100rb'
  ];
  samples.forEach(function (s) {
    var r = AIParser.parse(s);
    Logger_.info('parse: ' + s, {
      type: r.transaction_type, cat: r.category, sub: r.subcategory,
      amount: r.amount, account: r.account, to: r.account_to,
      conf: r.confidence
    });
  });
}
