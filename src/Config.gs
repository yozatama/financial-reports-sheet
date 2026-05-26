/**
 * Config.gs
 * -------------------------------------------------------------
 * Central configuration constants used across the project.
 * Anything that is "magic" (sheet name, column index, threshold)
 * should live here so the rest of the codebase stays declarative.
 * -------------------------------------------------------------
 */

const APP = Object.freeze({
  NAME: "Yoza's Financial Report",
  VERSION: '1.0.0',
  AUTHOR: 'Yoza Pratama',
  DEFAULT_CURRENCY: 'IDR',
  DEFAULT_LOCALE: 'id-ID',
  DEFAULT_TIMEZONE: 'Asia/Jakarta'
});

/** Centralised sheet names. Edit here to rename across the app. */
const SHEETS = Object.freeze({
  DASHBOARD: 'Dashboard',
  TRANSACTIONS: 'Transactions',
  ACCOUNTS: 'Accounts',
  DEBTS: 'Debts',
  CATEGORIES: 'Categories',
  BUDGETS: 'Budget Tracking',
  MONTHLY: 'Monthly Summary',
  SETTINGS: 'Settings',
  AI_LOGS: 'AI Logs'
});

/** Column definitions per sheet (1-indexed for Sheets API). */
const COLS = Object.freeze({
  TRANSACTIONS: {
    ID: 1, DATE: 2, TIME: 3, TYPE: 4, CATEGORY: 5, SUBCATEGORY: 6,
    AMOUNT: 7, ACCOUNT: 8, ACCOUNT_TO: 9, MERCHANT: 10, DESCRIPTION: 11,
    TAGS: 12, RECURRING: 13, CREATED_AT: 14, RAW_INPUT: 15, CONFIDENCE: 16
  },
  ACCOUNTS: {
    ID: 1, NAME: 2, TYPE: 3, INSTITUTION: 4, INITIAL_BALANCE: 5,
    CURRENT_BALANCE: 6, CURRENCY: 7, LAST_UPDATED: 8, STATUS: 9, NOTES: 10
  },
  DEBTS: {
    ID: 1, NAME: 2, PROVIDER: 3, TYPE: 4, OUTSTANDING: 5, LIMIT: 6,
    MIN_PAYMENT: 7, INTEREST: 8, DUE_DAY: 9, BILLING_DAY: 10,
    INSTALLMENT: 11, TENOR: 12, REMAINING_TENOR: 13, STATUS: 14,
    NOTES: 15, UPDATED_AT: 16
  },
  CATEGORIES: {
    ID: 1, TYPE: 2, CATEGORY: 3, SUBCATEGORY: 4, ICON: 5, COLOR: 6, ACTIVE: 7
  },
  BUDGETS: {
    ID: 1, MONTH: 2, CATEGORY: 3, LIMIT: 4, SPENT: 5, REMAINING: 6,
    PERCENT: 7, STATUS: 8
  }
});

/** Transaction & account type enums. */
const TYPES = Object.freeze({
  INCOME: 'Income',
  EXPENSE: 'Expense',
  TRANSFER: 'Transfer'
});

const ACCOUNT_TYPES = Object.freeze([
  'Cash', 'Bank', 'E-Wallet', 'Digital Bank', 'Credit Card', 'Investment', 'Savings'
]);

const DEBT_TYPES = Object.freeze([
  'Credit Card', 'PayLater', 'Personal Loan', 'Installment', 'Mortgage', 'Other'
]);

const STATUS = Object.freeze({
  ACTIVE: 'Active',
  INACTIVE: 'Inactive',
  PAID_OFF: 'Paid Off',
  OVERDUE: 'Overdue'
});

/** Visual theme - used for conditional formatting and HTML UIs. */
const THEME = Object.freeze({
  PRIMARY: '#5B5BD6',
  PRIMARY_DARK: '#3E3EBF',
  ACCENT: '#22D3A6',
  DANGER: '#EF4444',
  WARNING: '#F59E0B',
  INFO: '#3B82F6',
  SUCCESS: '#10B981',
  BG: '#0F1115',
  CARD: '#FFFFFF',
  MUTED: '#6B7280',
  BORDER: '#E5E7EB',
  HEADER_BG: '#1F2937',
  HEADER_FG: '#F9FAFB',
  ZEBRA: '#F9FAFB'
});

/** Health thresholds for insight engine. */
const THRESHOLDS = Object.freeze({
  HEALTHY_DEBT_RATIO: 0.36,        // monthly obligations / income
  HEALTHY_CREDIT_UTIL: 0.30,       // outstanding / limit
  HEALTHY_SAVINGS_RATE: 0.20,      // savings / income
  AI_CONFIDENCE_CONFIRM: 0.65,     // below -> ask user to confirm
  AI_CONFIDENCE_FALLBACK: 0.55     // below -> escalate to Gemini if available
});

/**
 * Helper: read a value from the Settings sheet by key.
 * Falls back to a default if the key is missing.
 */
function getSetting(key, defaultValue) {
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.SETTINGS);
    if (!sh) return defaultValue;
    var values = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).getValues();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === key) {
        var v = values[i][1];
        return (v === '' || v === null || v === undefined) ? defaultValue : v;
      }
    }
  } catch (err) {
    Logger_.error('getSetting failed', err);
  }
  return defaultValue;
}

function setSetting(key, value) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.SETTINGS);
  if (!sh) throw new Error('Settings sheet not found. Run Setup first.');
  var last = sh.getLastRow();
  var values = last >= 2 ? sh.getRange(2, 1, last - 1, 2).getValues() : [];
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      sh.getRange(i + 2, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value, '']);
}
