/**
 * AIParser.gs
 * -------------------------------------------------------------
 * Free-text -> structured transaction parser for Indonesian +
 * English casual language. Strategy:
 *
 *   1. Normalise input (lowercase, fix common typos, expand
 *      numeric abbreviations: rb / k / ribu / jt / juta / m).
 *   2. Detect amount via regex.
 *   3. Detect transaction type from action verb keywords.
 *   4. Match account via AccountService aliases.
 *   5. Match category via keyword dictionary.
 *   6. Score confidence based on what we matched.
 *   7. If confidence is below threshold AND OpenAI key is set
 *      -> escalate to OpenAIClient.parse() and merge.
 *   8. Always log the result to the AI Logs sheet.
 *
 * The parser returns a deterministic envelope:
 *   {
 *     transaction_type, category, subcategory, amount,
 *     account, account_to, description, merchant, tags,
 *     confidence, raw_input, method, reasons[]
 *   }
 * -------------------------------------------------------------
 */

var AIParser = (function () {

  /* -----------------------------------------------------------
   * Keyword dictionaries. Easy to extend without code changes.
   * ----------------------------------------------------------- */
  var ACTION_KEYWORDS = {
    income: [
      'gajian', 'gaji', 'salary', 'bonus', 'thr', 'dapet', 'dapat',
      'terima', 'masuk', 'income', 'cashback', 'refund', 'untung',
      'penghasilan', 'fee', 'honor', 'profit', 'dividen'
    ],
    transfer: [
      'tf', 'transfer', 'kirim', 'pindah', 'topup', 'top up', 'top-up',
      'isi saldo', 'tarik', 'withdraw', 'wd', 'setor'
    ],
    expense: [
      'bayar', 'beli', 'jajan', 'makan', 'minum', 'belanja', 'isi',
      'spend', 'keluar', 'cicil', 'cicilan', 'tagihan', 'pesen',
      'order', 'pesan', 'langganan', 'subscribe', 'sewa'
    ]
  };

  // Map of canonical category -> trigger keywords (Indonesian + English)
  var CATEGORY_KEYWORDS = {
    'Food & Beverage': {
      sub: {
        'Coffee': ['kopi', 'coffee', 'starbucks', 'kopken', 'kopken', 'janji jiwa', 'fore', 'kenangan', 'tuku', 'latte', 'americano'],
        'Restaurant': ['makan', 'resto', 'restoran', 'restaurant', 'sushi', 'ramen', 'pizza', 'burger', 'kfc', 'mcd', 'mcdonalds', 'hokben', 'warung', 'padang', 'soto', 'bakso', 'mie', 'ayam'],
        'Groceries': ['groceries', 'belanja bulanan', 'sayur', 'pasar', 'supermarket', 'indomaret', 'alfamart', 'hypermart', 'tokopedia mart', 'transmart'],
        'Snacks': ['snack', 'jajan', 'cemilan', 'cilok', 'gorengan', 'martabak', 'es', 'boba', 'chatime', 'kue', 'donat', 'donut']
      }
    },
    'Transport': {
      sub: {
        'Fuel': ['bensin', 'pertamax', 'pertalite', 'solar', 'fuel', 'spbu', 'shell', 'bp'],
        'Ride-hailing': ['gojek', 'grab', 'maxim', 'gocar', 'goride', 'grabcar', 'grabbike', 'taxi', 'taksi', 'blue bird', 'gobluebird'],
        'Public Transport': ['krl', 'mrt', 'lrt', 'bus', 'bis', 'transjakarta', 'kereta', 'busway', 'angkot'],
        'Parking & Toll': ['parkir', 'parking', 'tol', 'toll', 'e-toll', 'etoll']
      }
    },
    'Bills & Utilities': {
      sub: {
        'Electricity': ['listrik', 'pln', 'token listrik', 'electricity', 'tagihan listrik'],
        'Water': ['air', 'pdam', 'water'],
        'Internet': ['internet', 'wifi', 'indihome', 'biznet', 'iconnet', 'firstmedia', 'myrepublic'],
        'Phone': ['pulsa', 'paket data', 'kuota', 'telkomsel', 'simpati', 'xl', 'indosat', 'tri', '3', 'smartfren', 'axis', 'phone', 'pulsa hp']
      }
    },
    'Shopping': {
      sub: {
        'Clothes': ['baju', 'celana', 'kaos', 'sepatu', 'sandal', 'jaket', 'kemeja', 'fashion', 'uniqlo', 'zara', 'h&m'],
        'Electronics': ['elektronik', 'hp', 'laptop', 'tv', 'electronics', 'gadget', 'handphone', 'tablet', 'iphone', 'samsung'],
        'Online Shopping': ['shopee', 'tokopedia', 'lazada', 'tiktok shop', 'tiktokshop', 'bukalapak', 'blibli', 'amazon', 'belanja online', 'olshop']
      }
    },
    'Health': {
      sub: {
        'Medical': ['dokter', 'rumah sakit', 'rs ', 'klinik', 'medical', 'periksa', 'lab', 'rontgen', 'check up'],
        'Pharmacy': ['apotek', 'obat', 'kimia farma', 'guardian', 'watsons', 'pharmacy']
      }
    },
    'Entertainment': {
      sub: {
        'Movies': ['bioskop', 'cinema', 'xxi', 'cgv', 'movie', 'film'],
        'Subscriptions': ['netflix', 'spotify', 'disney', 'youtube premium', 'apple music', 'hbo', 'iqiyi', 'viu', 'wetv', 'subscription', 'langganan'],
        'Games': ['game', 'steam', 'mobile legend', 'ml', 'pubg', 'genshin', 'topup game', 'voucher game']
      }
    },
    'Education': {
      sub: {
        'Courses': ['kursus', 'course', 'udemy', 'coursera', 'pluralsight', 'kelas', 'training', 'workshop', 'bootcamp'],
        'Books': ['buku', 'book', 'ebook', 'gramedia', 'periplus']
      }
    },
    'Family': {
      sub: {
        'Kids': ['anak', 'sekolah', 'mainan', 'susu', 'popok', 'pampers', 'diaper'],
        'Parents': ['ortu', 'orang tua', 'mama', 'papa', 'ibu', 'bapak', 'nenek', 'kakek']
      }
    },
    'Debt Payment': {
      sub: {
        'Credit Card': ['cc', 'kartu kredit', 'credit card', 'tagihan cc', 'bayar cc'],
        'PayLater': ['paylater', 'pay later', 'spaylater', 'shopee paylater', 'gopaylater', 'kredivo', 'akulaku'],
        'Loan': ['cicilan', 'pinjol', 'pinjaman', 'kta', 'kredit', 'loan']
      }
    },
    'Salary': {
      sub: {
        'Monthly Salary': ['gaji', 'gajian', 'salary', 'payroll'],
        'Bonus': ['bonus', 'thr', 'tunjangan']
      }
    },
    'Freelance': {
      sub: { 'Project': ['freelance', 'project', 'proyek', 'fee', 'honor', 'kontrak'] }
    },
    'Investment': {
      sub: {
        'Dividend': ['dividen', 'dividend'],
        'Interest': ['bunga', 'interest']
      }
    },
    'Other Income': {
      sub: {
        'Refund': ['refund', 'pengembalian', 'cashback'],
        'Gift': ['hadiah', 'gift', 'angpao']
      }
    },
    'Transfer': {
      sub: {
        'Top Up': ['topup', 'top up', 'top-up', 'isi saldo'],
        'Withdraw': ['tarik', 'withdraw', 'wd', 'cash out'],
        'Between Accounts': ['transfer', 'tf', 'kirim', 'pindah dana']
      }
    }
  };

  /* -----------------------------------------------------------
   * Public entry point
   * ----------------------------------------------------------- */
  function parse(text) {
    require_(text, 'Input text required');
    var raw = String(text);
    var input = normalise_(raw);

    var reasons = [];
    var amount = detectAmount_(input, reasons);
    var type = detectType_(input, reasons);
    var accounts = detectAccounts_(input, reasons);
    var catMatch = detectCategory_(input, type, reasons);
    var tags = detectTags_(input);

    // Sensible defaults so the form is always populated
    if (!type) {
      // If we found two accounts -> transfer, else expense
      type = (accounts.from && accounts.to) ? TYPES.TRANSFER : TYPES.EXPENSE;
      reasons.push('Defaulted type to ' + type);
    }
    if (type === TYPES.TRANSFER && !accounts.to) {
      // try to interpret "ke <account>"
      var to = matchAccountAfter_(input, /ke\s+/);
      if (to) accounts.to = to;
    }
    // Top-up / isi saldo: the named account is the DESTINATION, not the source.
    if (type === TYPES.TRANSFER && accounts.from && !accounts.to &&
        /\b(topup|top up|top-up|isi saldo)\b/.test(input)) {
      accounts.to = accounts.from;
      accounts.from = '';
      reasons.push('Top-up: swapped account to destination');
    }
    // Withdraw / tarik: the named account is the SOURCE - already correct, no swap needed.

    var confidence = scoreConfidence_({ amount: amount, type: type, accounts: accounts, cat: catMatch });

    var result = {
      transaction_type: type,
      category: catMatch.category || (type === TYPES.INCOME ? 'Other Income' : (type === TYPES.TRANSFER ? 'Transfer' : 'Other')),
      subcategory: catMatch.subcategory || '',
      amount: amount || 0,
      account: accounts.from || guessDefaultAccount_(),
      account_to: accounts.to || '',
      description: shortDescription_(raw, catMatch),
      merchant: catMatch.merchant || '',
      tags: tags,
      confidence: round2_(confidence),
      raw_input: raw,
      method: 'rules',
      reasons: reasons
    };

    // Optional escalation to OpenAI for low-confidence inputs
    if (result.confidence < THRESHOLDS.AI_CONFIDENCE_OPENAI) {
      var openAIResult = OpenAIClient.tryParse(raw);
      if (openAIResult && openAIResult.ok) {
        result = mergeAIResult_(result, openAIResult.data);
        result.method = 'rules+openai';
        result.confidence = Math.max(result.confidence, openAIResult.data.confidence || 0.75);
      }
    }

    logToSheet_(result);
    return result;
  }

  /** Save the parsed result as a real transaction. */
  function commit(parsed) {
    require_(parsed, 'Parsed payload required');
    return TransactionService.add({
      type: parsed.transaction_type,
      category: parsed.category,
      subcategory: parsed.subcategory,
      amount: parsed.amount,
      account: parsed.account,
      accountTo: parsed.account_to,
      merchant: parsed.merchant,
      description: parsed.description,
      tags: parsed.tags,
      rawInput: parsed.raw_input,
      confidence: parsed.confidence
    });
  }

  /* -----------------------------------------------------------
   * Helpers
   * ----------------------------------------------------------- */

  function normalise_(s) {
    return String(s)
      .toLowerCase()
      .replace(/[\u00A0\u2000-\u200B]/g, ' ')   // weird whitespace
      .replace(/[,;]/g, ' ')
      .replace(/\s+/g, ' ')
      .replace(/\b(rp\.?|idr)\b/g, ' ')         // strip currency markers
      .replace(/(\d)\s+(\d)/g, '$1$2')          // join "1 000" -> "1000"
      .trim();
  }

  /**
   * Detect amount. Supports:
   *   "35rb", "35 rb", "35 ribu", "1.2jt", "1,2 juta",
   *   "450k", "8jt", "100rb", "120000".
   */
  function detectAmount_(input, reasons) {
    // Try suffixed numbers first (most common in casual ID text)
    var suffixed = input.match(/(\d+(?:[.,]\d+)?)\s*(rb|ribu|k|jt|juta|jeti|m|mio)\b/);
    if (suffixed) {
      var n = parseFloat(suffixed[1].replace(',', '.'));
      var unit = suffixed[2];
      var mult = (unit === 'rb' || unit === 'ribu' || unit === 'k') ? 1000
        : (unit === 'jt' || unit === 'juta' || unit === 'jeti') ? 1000000
        : (unit === 'm' || unit === 'mio') ? 1000000 : 1;
      reasons.push('Amount via suffix "' + unit + '"');
      return Math.round(n * mult);
    }
    // Plain number with optional thousand separators
    var plain = input.match(/(\d{1,3}(?:[.,]\d{3})+|\d{4,})/);
    if (plain) {
      var v = parseInt(plain[1].replace(/[.,]/g, ''), 10);
      reasons.push('Amount via plain number');
      return v;
    }
    // Last resort: any number
    var any = input.match(/\b(\d+)\b/);
    if (any) {
      reasons.push('Amount via fallback');
      return parseInt(any[1], 10);
    }
    return 0;
  }

  function detectType_(input, reasons) {
    if (containsAny_(input, ACTION_KEYWORDS.transfer)) {
      reasons.push('Type=Transfer keyword');
      return TYPES.TRANSFER;
    }
    if (containsAny_(input, ACTION_KEYWORDS.income)) {
      reasons.push('Type=Income keyword');
      return TYPES.INCOME;
    }
    if (containsAny_(input, ACTION_KEYWORDS.expense)) {
      reasons.push('Type=Expense keyword');
      return TYPES.EXPENSE;
    }
    return null;
  }

  /**
   * Match a known account name (or alias) appearing in the text.
   * Returns { from, to } where `to` is filled if pattern "ke X" or "to X" found.
   */
  function detectAccounts_(input, reasons) {
    var accounts = AccountService.list(true);
    var aliases = {
      'cash': ['cash', 'tunai', 'dompet'],
      'bca': ['bca', 'mandiri bca'],
      'mandiri': ['mandiri'],
      'bni': ['bni'],
      'bri': ['bri'],
      'seabank': ['seabank', 'sea bank'],
      'jago': ['jago', 'bank jago'],
      'jenius': ['jenius'],
      'gopay': ['gopay', 'go pay'],
      'ovo': ['ovo'],
      'dana': ['dana'],
      'shopeepay': ['shopeepay', 'shopee pay', 'spay'],
      'linkaja': ['linkaja', 'link aja']
    };

    function matchOne(text) {
      // First, try direct match against actual account names
      for (var i = 0; i < accounts.length; i++) {
        var name = String(accounts[i]['Account Name']).toLowerCase();
        if (text.indexOf(name) !== -1) return accounts[i]['Account Name'];
      }
      // Then, alias lookup
      for (var key in aliases) {
        for (var j = 0; j < aliases[key].length; j++) {
          if (text.indexOf(aliases[key][j]) !== -1) {
            // find an account whose name contains the canonical key
            var found = accounts.find(function (a) {
              return String(a['Account Name']).toLowerCase().indexOf(key) !== -1;
            });
            if (found) return found['Account Name'];
          }
        }
      }
      return null;
    }

    var to = null, from = null;
    // "ke <account>" or "to <account>" indicates destination of a transfer
    var keMatch = input.match(/\b(?:ke|to)\s+([a-z0-9 ]+)$/);
    if (keMatch) {
      to = matchOne(keMatch[1]);
      if (to) reasons.push('Detected destination via "ke ' + to + '"');
    }
    from = matchOne(input);
    if (from && from === to) from = null;
    if (from) reasons.push('Detected source account ' + from);

    // "pake/pakai/dengan <account>" hints source account
    if (!from) {
      var pakeMatch = input.match(/\b(?:pake|pakai|dengan|via|with|dari|from)\s+([a-z0-9 ]+)/);
      if (pakeMatch) {
        var p = matchOne(pakeMatch[1]);
        if (p) { from = p; reasons.push('Source via "pake ' + p + '"'); }
      }
    }
    return { from: from, to: to };
  }

  function matchAccountAfter_(input, prefixRegex) {
    var m = input.match(new RegExp(prefixRegex.source + '([a-z0-9 ]+)'));
    if (!m) return null;
    var fragment = m[1];
    var accs = AccountService.list(true);
    for (var i = 0; i < accs.length; i++) {
      if (fragment.indexOf(String(accs[i]['Account Name']).toLowerCase()) !== -1) {
        return accs[i]['Account Name'];
      }
    }
    return null;
  }

  function detectCategory_(input, type, reasons) {
    var best = { score: 0 };
    Object.keys(CATEGORY_KEYWORDS).forEach(function (cat) {
      var subs = CATEGORY_KEYWORDS[cat].sub || {};
      Object.keys(subs).forEach(function (sub) {
        subs[sub].forEach(function (kw) {
          if (input.indexOf(kw) !== -1) {
            var score = kw.length; // longer match wins
            if (score > best.score) {
              best = { score: score, category: cat, subcategory: sub, merchant: titleCase_(kw) };
            }
          }
        });
      });
    });
    if (best.category) reasons.push('Category=' + best.category + ' / ' + best.subcategory);

    // If the type is Transfer but we matched a non-transfer category, prefer Transfer
    if (type === TYPES.TRANSFER && best.category !== 'Transfer') {
      best = { score: 1, category: 'Transfer', subcategory: 'Between Accounts' };
    }
    return best;
  }

  function detectTags_(input) {
    var tags = [];
    if (/\b(urgent|penting)\b/.test(input)) tags.push('urgent');
    if (/\b(refund|cashback)\b/.test(input)) tags.push('refund');
    if (/\b(monthly|bulanan|recurring)\b/.test(input)) tags.push('recurring');
    if (/\b(business|kantor|kerjaan)\b/.test(input)) tags.push('business');
    return tags;
  }

  function shortDescription_(raw, catMatch) {
    if (catMatch && catMatch.category) {
      return titleCase_(catMatch.subcategory || catMatch.category);
    }
    return String(raw).slice(0, 80);
  }

  function guessDefaultAccount_() {
    var accs = AccountService.list(true);
    if (!accs.length) return '';
    // Prefer Cash if available
    var cash = accs.find(function (a) { return a['Account Type'] === 'Cash'; });
    return (cash || accs[0])['Account Name'];
  }

  function scoreConfidence_(p) {
    var score = 0.0;
    if (p.amount && p.amount > 0) score += 0.35;
    if (p.type) score += 0.20;
    if (p.cat && p.cat.category) score += 0.25;
    if (p.accounts && p.accounts.from) score += 0.15;
    if (p.type === TYPES.TRANSFER && p.accounts && p.accounts.to) score += 0.05;
    return Math.min(1, score);
  }

  function mergeAIResult_(rule, ai) {
    return {
      transaction_type: ai.transaction_type || rule.transaction_type,
      category: ai.category || rule.category,
      subcategory: ai.subcategory || rule.subcategory,
      amount: ai.amount || rule.amount,
      account: ai.account || rule.account,
      account_to: ai.account_to || rule.account_to,
      description: ai.description || rule.description,
      merchant: ai.merchant || rule.merchant,
      tags: (ai.tags && ai.tags.length) ? ai.tags : rule.tags,
      confidence: ai.confidence || rule.confidence,
      raw_input: rule.raw_input,
      method: rule.method,
      reasons: rule.reasons.concat(['OpenAI used'])
    };
  }

  function logToSheet_(result) {
    try {
      var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.AI_LOGS);
      if (!sh) return;
      sh.appendRow([
        new Date(), result.raw_input, result.method, result.confidence,
        result.transaction_type, result.category, result.amount, result.account,
        result.confidence >= THRESHOLDS.AI_CONFIDENCE_CONFIRM ? 'OK' : 'NEEDS_CONFIRM',
        safeStringify_(result)
      ]);
    } catch (e) { Logger_.error('AI log failed', e); }
  }

  function containsAny_(text, words) {
    for (var i = 0; i < words.length; i++) {
      if (text.indexOf(words[i]) !== -1) return true;
    }
    return false;
  }

  function titleCase_(s) {
    return String(s || '').replace(/\b\w/g, function (c) { return c.toUpperCase(); });
  }

  function round2_(n) { return Math.round(n * 100) / 100; }

  return { parse: parse, commit: commit };
})();
