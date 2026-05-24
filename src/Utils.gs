/**
 * Utils.gs
 * -------------------------------------------------------------
 * Pure helper functions: ID generation, date helpers, formatting,
 * a tiny structured logger and a small Result helper.
 *
 * Nothing in this file should depend on any service - keep it
 * dependency free so any module can require it safely.
 * -------------------------------------------------------------
 */

/** Generate a short, time-sortable, human readable id. */
function genId(prefix) {
  var d = new Date();
  var pad = function (n, w) { return String(n).padStart(w || 2, '0'); };
  var stamp = d.getFullYear() +
    pad(d.getMonth() + 1) + pad(d.getDate()) +
    pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds());
  var rand = Math.floor(Math.random() * 9000 + 1000);
  return (prefix || 'ID') + '-' + stamp + '-' + rand;
}

/** Today as midnight Date in spreadsheet timezone. */
function today_() {
  var tz = APP.DEFAULT_TIMEZONE;
  var s = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  return new Date(s + 'T00:00:00');
}

function startOfMonth_(d) {
  d = d || new Date();
  return new Date(d.getFullYear(), d.getMonth(), 1);
}

function endOfMonth_(d) {
  d = d || new Date();
  return new Date(d.getFullYear(), d.getMonth() + 1, 0, 23, 59, 59);
}

function monthKey_(d) {
  d = d || new Date();
  return Utilities.formatDate(d, APP.DEFAULT_TIMEZONE, 'yyyy-MM');
}

function fmtDate_(d) {
  if (!(d instanceof Date)) d = new Date(d);
  return Utilities.formatDate(d, APP.DEFAULT_TIMEZONE, 'yyyy-MM-dd');
}

function fmtTime_(d) {
  if (!(d instanceof Date)) d = new Date(d);
  return Utilities.formatDate(d, APP.DEFAULT_TIMEZONE, 'HH:mm:ss');
}

function fmtMoney_(n, currency) {
  currency = currency || APP.DEFAULT_CURRENCY;
  if (typeof n !== 'number') n = parseFloat(n) || 0;
  // Lightweight IDR-friendly formatter (Apps Script has no Intl in some envs).
  var sign = n < 0 ? '-' : '';
  var abs = Math.abs(n).toFixed(0);
  var withSep = abs.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  var symbol = currency === 'IDR' ? 'Rp ' : (currency + ' ');
  return sign + symbol + withSep;
}

/**
 * Tiny structured logger. Writes to Logger and (best-effort) to AI Logs sheet.
 * Note: name underscored to avoid clashing with the global Logger.
 */
var Logger_ = (function () {
  function write(level, msg, data) {
    var line = '[' + level + '] ' + msg + (data ? ' :: ' + safeStringify_(data) : '');
    try { Logger.log(line); } catch (e) { /* ignore */ }
  }
  return {
    info: function (m, d) { write('INFO', m, d); },
    warn: function (m, d) { write('WARN', m, d); },
    error: function (m, d) { write('ERROR', m, d && d.stack ? d.stack : d); },
    debug: function (m, d) { write('DEBUG', m, d); }
  };
})();

function safeStringify_(o) {
  try { return JSON.stringify(o); } catch (e) { return String(o); }
}

/** Throw a friendly error if value is missing. */
function require_(value, message) {
  if (value === null || value === undefined || value === '') {
    throw new Error(message || 'Required value missing');
  }
  return value;
}

/** Unique values from a 1-D array, preserving order. */
function unique_(arr) {
  var seen = {}, out = [];
  for (var i = 0; i < arr.length; i++) {
    var k = String(arr[i]);
    if (!seen[k] && arr[i] !== '' && arr[i] !== null && arr[i] !== undefined) {
      seen[k] = 1;
      out.push(arr[i]);
    }
  }
  return out;
}

/** Find row index (1-indexed including header) of a value in a column, or -1. */
function findRowByValue_(sheet, col, value) {
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var values = sheet.getRange(2, col, last - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(value).trim()) return i + 2;
  }
  return -1;
}

/** Read a sheet to an array of objects keyed by header row. */
function readObjects_(sheet) {
  var last = sheet.getLastRow();
  var width = sheet.getLastColumn();
  if (last < 2) return [];
  var headers = sheet.getRange(1, 1, 1, width).getValues()[0];
  var rows = sheet.getRange(2, 1, last - 1, width).getValues();
  return rows.map(function (r) {
    var o = {};
    for (var i = 0; i < headers.length; i++) o[headers[i]] = r[i];
    return o;
  });
}

/** Robust HTML escape for embedding values into UI templates. */
function htmlEscape_(s) {
  if (s === null || s === undefined) return '';
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/**
 * Wrap a function call so it never throws across the
 * google.script.run boundary - returns a Result envelope.
 */
function safeCall_(fn) {
  try {
    var data = fn();
    return { ok: true, data: data };
  } catch (err) {
    Logger_.error('safeCall failed', err);
    return { ok: false, error: String(err && err.message || err) };
  }
}
