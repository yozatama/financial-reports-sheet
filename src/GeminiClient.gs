/**
 * GeminiClient.gs
 * -------------------------------------------------------------
 * Optional Google Gemini fallback for transaction parsing. Reads
 * `gemini_api_key` and `gemini_model` from the Settings sheet.
 * If no key is configured, tryParse() simply returns null and
 * the rule-based parser owns the result.
 *
 * Why Gemini for this use case:
 *  - Native JSON-mode (responseMimeType + responseSchema) gives
 *    deterministic, type-safe output without prompt gymnastics.
 *  - Gemini 2.5 Flash-Lite is one of the cheapest hosted models,
 *    well below US$0.001 per parse on typical inputs.
 *
 * Endpoint:
 *   POST https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key=API_KEY
 * -------------------------------------------------------------
 */

var GeminiClient = (function () {

  var ENDPOINT = 'https://generativelanguage.googleapis.com/v1beta/models/';

  function isConfigured() {
    var key = getSetting('gemini_api_key', '');
    return !!(key && String(key).trim());
  }

  /** Returns { ok, data } or null if not configured / failed. */
  function tryParse(text) {
    if (!isConfigured()) return null;
    try {
      var parsed = parse(text);
      return { ok: true, data: parsed };
    } catch (err) {
      Logger_.error('Gemini parse failed', err);
      return { ok: false, error: String(err) };
    }
  }

  function parse(text) {
    var apiKey = String(getSetting('gemini_api_key', '')).trim();
    var model = String(getSetting('gemini_model', 'gemini-2.5-flash-lite')).trim();
    if (!apiKey) throw new Error('Gemini API key not set');

    var accounts = AccountService.names(true);
    var categoriesByType = {
      Income: CategoryService.topLevelNames(TYPES.INCOME),
      Expense: CategoryService.topLevelNames(TYPES.EXPENSE),
      Transfer: CategoryService.topLevelNames(TYPES.TRANSFER)
    };

    var systemInstruction = [
      'You are a personal finance NLP parser.',
      'Input is casual Indonesian and/or English describing exactly one transaction.',
      'Indonesian numeric suffixes: rb / k / ribu => x1000; jt / juta / mio => x1000000.',
      'Pick category and account values from the lists below; if you cannot match, use the closest one.',
      'Allowed accounts: ' + JSON.stringify(accounts),
      'Allowed categories by type: ' + JSON.stringify(categoriesByType),
      'For Transfer: account_to is the destination; account is the source.',
      'For Top Up / isi saldo: the named wallet is the destination (account_to).',
      'For Withdraw / tarik: the named wallet is the source (account).',
      'confidence is your self-assessment between 0 and 1.'
    ].join('\n');

    // Strict response schema — Gemini guarantees JSON conforming to this.
    var responseSchema = {
      type: 'OBJECT',
      properties: {
        transaction_type: { type: 'STRING', enum: [TYPES.INCOME, TYPES.EXPENSE, TYPES.TRANSFER] },
        category: { type: 'STRING' },
        subcategory: { type: 'STRING' },
        amount: { type: 'NUMBER' },
        account: { type: 'STRING' },
        account_to: { type: 'STRING' },
        description: { type: 'STRING' },
        merchant: { type: 'STRING' },
        tags: { type: 'ARRAY', items: { type: 'STRING' } },
        confidence: { type: 'NUMBER' }
      },
      required: ['transaction_type', 'category', 'amount', 'account', 'confidence']
    };

    var payload = {
      systemInstruction: { parts: [{ text: systemInstruction }] },
      contents: [{ role: 'user', parts: [{ text: text }] }],
      generationConfig: {
        temperature: 0.1,
        responseMimeType: 'application/json',
        responseSchema: responseSchema
      }
    };

    var url = ENDPOINT + encodeURIComponent(model) + ':generateContent?key=' + encodeURIComponent(apiKey);
    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = resp.getResponseCode();
    var body = resp.getContentText();
    if (code < 200 || code >= 300) {
      throw new Error('Gemini API error ' + code + ': ' + body.slice(0, 500));
    }
    var data = JSON.parse(body);
    var candidates = data && data.candidates;
    if (!candidates || !candidates.length) {
      throw new Error('Gemini returned no candidates: ' + body.slice(0, 200));
    }
    var content = candidates[0] && candidates[0].content && candidates[0].content.parts &&
                  candidates[0].content.parts[0] && candidates[0].content.parts[0].text;
    if (!content) throw new Error('Gemini returned empty content');
    return JSON.parse(content);
  }

  return { isConfigured: isConfigured, tryParse: tryParse, parse: parse };
})();
