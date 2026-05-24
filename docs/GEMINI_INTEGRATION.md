# Optional Google Gemini Integration

The rule-based parser handles the vast majority of casual Indonesian + English
inputs without any external API. For tricky inputs (creative slang, missing
keywords, novel merchants), you can enable a Gemini fallback that kicks in
**only when the rule-based confidence is below the configured threshold**.

## Get an API key

1. Open <https://aistudio.google.com/apikey> and click *"Create API key"*.
2. Pick a Google Cloud project (the prompt offers a free one if you don't have one).
3. Copy the key — it looks like `AIza...`.

## Enable in the spreadsheet

1. Open the `Settings` sheet.
2. Paste the key into the `gemini_api_key` row.
3. (Optional) Change `gemini_model`. Defaults to **`gemini-2.5-flash-lite`** —
   the cheapest GA model in the Gemini 2.5 family, perfect for short text
   classification. Other good choices:
   - `gemini-2.5-flash` — slightly higher quality, ~3x the price.
   - `gemini-3.5-flash` — newest fast tier (May 2026), strong on agentic prompts.

That's it. The next time you use **Quick Add (AI)** with input that scores
below `THRESHOLDS.AI_CONFIDENCE_FALLBACK` (0.55 by default), the parser will
escalate to Gemini, merge the result, and bump the confidence label. The
`AI Logs` sheet records every call's method as either `rules` or
`rules+gemini`.

## How it works

`GeminiClient.parse(text)` sends a `generateContent` request with:

- **`systemInstruction`** describing the strict output contract and the
  allowed account/category names *from your spreadsheet* — so the model
  can never invent a category that isn't yours.
- **`contents`** carrying the raw user input.
- **`generationConfig.responseMimeType: "application/json"`** plus a strict
  **`responseSchema`** so the response is guaranteed to be parseable JSON
  matching the parser's envelope shape.
- `temperature: 0.1` for deterministic output.

`AIParser.mergeAIResult_` then layers the Gemini result on top of the rule-based
result, preferring AI fields when present and keeping rule-based fallbacks
otherwise.

## Cost & rate limits

Every fallback is one `generateContent` call. With `gemini-2.5-flash-lite` at
~$0.10 / 1M input tokens and ~$0.40 / 1M output tokens (May 2026 pricing),
the typical cost per parse is well below US $0.0005. Free tier covers a
generous monthly quota for personal use.

If you log lots of low-confidence inputs, keep an eye on the `AI Logs`
sheet — you can lower the escalation threshold in
`Config.gs::THRESHOLDS.AI_CONFIDENCE_FALLBACK` to escalate less often.

## Privacy

- The raw input text is sent to Google's Gemini API when escalation triggers.
- Account names and category names from your spreadsheet are included in the
  system prompt as a closed vocabulary.
- Account **balances** and individual transactions are **not** sent.
- Disable the integration any time by clearing the `gemini_api_key` cell.

## Going further

- **Switch model:** edit the `gemini_model` cell in `Settings`. Any
  `generateContent`-capable model on the Gemini API works.
- **Vertex AI / enterprise auth:** replace `GeminiClient.gs` with a Vertex
  variant that uses an OAuth token instead of an API key — Apps Script can
  obtain one via `ScriptApp.getOAuthToken()` if you add the right scope.
- **Self-hosted LLM:** you can also point at any OpenAI-compatible endpoint
  (vLLM, LiteLLM, Ollama) by writing a small client; the parser only depends
  on `tryParse(text) -> { ok, data } | null`.
- You can also call `GeminiClient.parse()` directly from the editor for
  quick prompt iteration.
