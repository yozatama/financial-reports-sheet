# Kiro Finance — Personal Finance Spreadsheet System

A polished personal-finance management system that runs entirely on **Google
Sheets + Google Apps Script**, with an AI-assisted natural-language transaction
parser tuned for Indonesian + English casual input.

The whole thing feels like a small fintech SaaS, but lives in a single
spreadsheet you fully control.

---

## Highlights

- **Modern dashboard** — KPI cards, daily expense trend, top categories,
  debt monitoring, cash-flow heatmap, automatic insights.
- **Quick Add (AI)** — type *"jajan kopi 35rb pake gopay"* and the parser
  fills in type, category, amount, account, merchant, tags with a confidence
  score. Optional Google Gemini fallback for low-confidence inputs.
- **Wallets & banks** — Cash, BCA, Mandiri, SeaBank, GoPay, OVO, DANA,
  ShopeePay, etc. Balances auto-update from transactions.
- **Debt tracking** — Credit Cards, PayLater, loans, installments, with
  utilization, due-date radar, and payoff timeline.
- **Budgets** — per-category monthly limits with status (On Track / Warning /
  Over Budget) driven by conditional formatting.
- **Automation** — recurring transactions, monthly rollover, optional
  email reminders for upcoming due dates.

## Repository layout

```
.
├── README.md
├── docs/
│   ├── SETUP.md
│   └── GEMINI_INTEGRATION.md
└── src/
    ├── appsscript.json
    ├── Code.gs                  # menu + dialog launchers + server endpoints
    ├── Config.gs                # constants, sheet names, theme
    ├── Utils.gs                 # logger, helpers, formatters
    ├── SetupService.gs          # creates/repairs every sheet
    ├── TransactionService.gs    # transaction CRUD + balance side-effects
    ├── AccountService.gs        # account CRUD + balance recompute
    ├── DebtService.gs           # debt analytics (utilization, due dates)
    ├── CategoryService.gs       # category CRUD
    ├── BudgetService.gs         # budget recompute + status
    ├── AIParser.gs              # rule-based ID/EN NLP + confidence
    ├── GeminiClient.gs          # optional AI fallback (Google Gemini)
    ├── DashboardService.gs      # dashboard layout + insights
    ├── Triggers.gs              # daily / monthly automations
    ├── Styles.html              # shared CSS (light + dark)
    ├── Sidebar.html             # sidebar UI
    ├── QuickAdd.html            # AI quick-add dialog
    ├── AddTransaction.html      # full transaction form
    ├── AddAccount.html          # account form
    └── AddDebt.html             # debt form
```

## Quick start

See [`docs/SETUP.md`](docs/SETUP.md) for the full walk-through. TL;DR:

1. Create a new Google Sheet.
2. Open **Extensions → Apps Script** and paste each file from `src/` into
   matching script files of the same name (or use [clasp](https://github.com/google/clasp)).
3. Save, reload the spreadsheet, and choose **Kiro Finance → Setup → Run Initial Setup**.
4. Use **Kiro Finance → Quick Add (AI)** to start logging transactions in plain language.
5. (Optional) Add a Gemini API key in the *Settings* sheet — see
   [`docs/GEMINI_INTEGRATION.md`](docs/GEMINI_INTEGRATION.md).

## Example AI inputs

```
jajan kopi 35rb pake gopay         → Expense / Food & Beverage / Coffee / 35,000 / GoPay
bayar listrik 450k                 → Expense / Bills & Utilities / Electricity / 450,000
gajian 8jt                         → Income / Salary / Monthly Salary / 8,000,000
isi bensin motor 50 ribu           → Expense / Transport / Fuel / 50,000
tf ke bca 1.2jt                    → Transfer / → BCA / 1,200,000
makan sushi 120rb pake gopay       → Expense / Food & Beverage / Restaurant / 120,000 / GoPay
bayar cc bca 2jt                   → Expense / Debt Payment / Credit Card / 2,000,000
topup ovo 100rb                    → Transfer / Top Up / OVO / 100,000
```

## Design language

Inspired by **Notion Finance, YNAB, Copilot Money, Monarch Money, Revolut**.
Light + dark friendly, soft shadows, rounded cards, generous spacing, mobile-first
forms. Theme tokens live in `Styles.html` and `Config.gs::THEME`.

## License

MIT — go wild, fork it, make it yours.
