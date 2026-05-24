/**
 * OpenAIClient.gs
 * -------------------------------------------------------------
 * Optional OpenAI fallback for transaction parsing. Reads
 * `openai_api_key` and `openai_model` from the Settings sheet.
 * If no key is configured, tryParse() simply returns null and
 * the rule-based parser owns the result.
 *
 * The prompt is tuned for Indonesian + English casual finance
 * input and asks the model to return strict JSON.
 * -------------------------------------------------------------
 */

var OpenAIClient = (function () {

  var ENDPOINT = 'https://api.openai.com/v1/chat/completions';

  function isConfigured() {
    var key = getSetting('openai_api_key', '');
    return !!(key && String(key).trim());
  }

  /** Returns { ok, data } or null if not configured / failed. */
  function tryParse(text) {
    if (!isConfigured()) return null;
    try {
      var parsed = parse(text);
      return { ok: true, data: parsed };
    } catch (err) {
      Logger_.error('OpenAI parse failed', err);
      return { ok: false, error: String(err) };
    }
  }

  function parse(text) {
    var apiKey = String(getSetting('openai_api_key', '')).trim();
    var model = String(getSetting('openai_model', 'gpt-4o-mini')).trim();
    if (!apiKey) throw new Error('OpenAI API key not set');

    var accounts = AccountService.names(true);
    var categoriesByType = {
      Income: CategoryService.topLevelNames(TYPES.INCOME),
      Expense: CategoryService.topLevelNames(TYPES.EXPENSE),
      Transfer: CategoryService.topLevelNames(TYPES.TRANSFER)
    };

    var system = [
      'You are a personal finance NLP parser.',
      'Input is casual Indonesian and/or English describing one transaction.',
      'Always reply with STRICT JSON only, no prose, no markdown fences.',
      'Schema:',
      '{',
      '  "transaction_type": "Income|Expense|Transfer",',
      '  "category": "string",',
      '  "subcategory": "string",',
      '  "amount": number,',
      '  "account": "string",',
      '  "account_to": "string",',
      '  "description": "string",',
      '  "merchant": "string",',
      '  "tags": ["string"],',
      '  "confidence": 0.0-1.0',
      '}',
      'Indonesian numeric suffixes: rb/k/ribu = x1000; jt/juta = x1000000.',
      'Allowed accounts: ' + JSON.stringify(accounts),
      'Allowed categories: ' + JSON.stringify(categoriesByType)
    ].join('\n');

    var payload = {
      model: model,
      temperature: 0.1,
      messages: [
        { role: 'system', content: system },
        { role: 'user', content: text }
      ],
      response_format: { type: 'json_object' }
    };

    var resp = UrlFetchApp.fetch(ENDPOINT, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = resp.getResponseCode();
    var body = resp.getContentText();
    if (code < 200 || code >= 300) {
      throw new Error('OpenAI API error ' + code + ': ' + body);
    }
    var data = JSON.parse(body);
    var content = data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
    if (!content) throw new Error('No content in OpenAI response');
    return JSON.parse(content);
  }

  return { isConfigured: isConfigured, tryParse: tryParse, parse: parse };
})();
