# Optional OpenAI Integration

The rule-based parser handles the vast majority of casual Indonesian + English
inputs without any external API. For tricky inputs (creative slang, missing
keywords, novel merchants), you can enable an OpenAI fallback that kicks in
**only when the rule-based confidence is below the configured threshold**.

## Enable

1. Get an API key from <https://platform.openai.com/api-keys>.
2. Open the `Settings` sheet in your spreadsheet.
3. Paste the key into the `openai_api_key` row.
4. (Optional) Change `openai_model` (default: `gpt-4o-mini`). Any chat-completions
   model that supports `response_format: json_object` works.

That's it. The next time you use **Quick Add (AI)** with input that scores
below `THRESHOLDS.AI_CONFIDENCE_OPENAI` (0.55 by default), the parser will
escalate to OpenAI, merge the results, and bump the confidence label. The
`AI Logs` sheet records every call's method as either `rules` or `rules+openai`.

## How it works

`OpenAIClient.parse(text)` sends a chat completion with:

- **System prompt** describing the strict JSON schema and the allowed
  account/category names *from your spreadsheet* (so the model can never
  invent a category that isn't yours).
- **User prompt** is the raw input.
- `temperature: 0.1` and `response_format: { type: "json_object" }` to
  keep replies deterministic and parseable.

`AIParser.mergeAIResult_` then layers the AI result on top of the rule-based
result, preferring AI fields when present and keeping rule-based fallbacks
otherwise.

## Cost & rate limits

Every fallback is one chat-completion call. With `gpt-4o-mini` the typical
cost per parse is < US $0.0005. If you log lots of low-confidence inputs,
keep an eye on the `AI Logs` sheet — you can lower the threshold in
`Config.gs::THRESHOLDS.AI_CONFIDENCE_OPENAI` to escalate less often.

## Privacy

- The raw input text is sent to OpenAI when escalation triggers.
- Account names and category names from your spreadsheet are included in the
  system prompt as a closed vocabulary.
- Account **balances** and individual transactions are **not** sent.
- Disable the integration any time by clearing the `openai_api_key` cell.

## Going further

- Swap `OpenAIClient` for any other LLM provider — replace the
  `ENDPOINT` and request payload in `OpenAIClient.gs`.
- For privacy-sensitive setups, point the endpoint at a self-hosted LLM
  served via OpenAI-compatible APIs (e.g. vLLM, LiteLLM, Ollama with the
  OpenAI proxy).
- You can also call `OpenAIClient.parse()` directly from the editor for
  quick prompt iteration.
