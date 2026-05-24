# Setup Guide

Two paths: **manual paste** (5 min, no tooling) or **clasp** (recommended for
ongoing updates).

---

## Option A — Manual paste

1. Create a new Google Sheet at <https://sheets.new>.
2. Rename it to whatever you like (e.g. *My Finance*).
3. **Extensions → Apps Script**. The script editor opens in a new tab.
4. Delete the default `Code.gs`. Then, for **every** file under `src/` in this
   repository, create a matching file in the editor with the **same name and
   extension** (`.gs` for scripts, `.html` for HTML) and paste the contents.
   - Apps Script does not support sub-folders. All files live at the root
     of the script project.
   - You can copy multiple files quickly by clicking the `+` icon next to
     "Files" in the editor.
5. Open the manifest (`appsscript.json`):
   - In the editor, click ⚙️ **Project Settings** and tick
     *"Show appsscript.json manifest file"*.
   - Replace its contents with `src/appsscript.json` from this repo.
6. **Save** the project (`⌘/Ctrl + S`).
7. Reload the Google Sheet.
8. A new menu **"Kiro Finance"** appears. Run it once — Google will ask you to
   authorise the script. Accept the permissions.
9. Click **Kiro Finance → Setup → Run Initial Setup**. This creates every
   sheet (Dashboard, Transactions, Accounts, Debts, Budgets, etc.) with seed
   data and validations.
10. (Optional) **Kiro Finance → Setup → Install Daily Triggers** to enable the
    07:00 maintenance job.

You're done. Try **Kiro Finance → Quick Add (AI)** and type something like
`jajan kopi 35rb pake gopay`.

## Option B — clasp (recommended)

```bash
npm i -g @google/clasp
clasp login
git clone https://github.com/<your-fork>/financial-reports-sheet.git
cd financial-reports-sheet/src

# Create a new container-bound script for an existing sheet:
clasp create --type sheets --title "Kiro Finance" --rootDir .

# Or attach to an existing script project (paste the script ID):
# echo '{"scriptId":"<SCRIPT_ID>","rootDir":"."}' > .clasp.json

clasp push
```

Then open the bound spreadsheet, reload, and follow steps 9–10 above.

## Permissions you will see on first run

- **Spreadsheet (current only)** — read/write the active sheet.
- **Script container UI** — show menus, dialogs, and the sidebar.
- **Send mail as you** — only used if you enable email reminders in
  `Settings`. Nothing leaves your account otherwise.
- **External request** — only used by the optional OpenAI fallback.
- **Run when you are not present** — needed by the daily / monthly triggers.

If you don't want any of these, edit `src/appsscript.json` and remove the
corresponding scope; just be aware features that require it will fail.

## Customising

- **Categories** — edit the `Categories` sheet directly. New categories
  immediately appear in the dropdowns thanks to data validation by range.
- **Accounts** — add a row in `Accounts` or use the *Add Account* dialog.
  Account names typed in transactions must match exactly (the AI parser
  already maps the common aliases: gopay/ovo/dana/shopeepay/seabank/jago/...).
- **Theme** — colours live in two places:
  - Sheets-side: `THEME` in `src/Config.gs`.
  - HTML-side: CSS variables in `src/Styles.html`.
- **Default month / locale / timezone** — change `APP.DEFAULT_*` in
  `Config.gs` and the `Settings` sheet.

## Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| Menu missing after install | Reload the spreadsheet tab; `onOpen` runs on load. |
| Dropdowns empty in dialogs | Run *Setup → Run Initial Setup* once so the lookup sheets exist. |
| Balances not updating | Run *Setup → Recompute Balances*. The recompute is also triggered automatically after every Add/Edit/Delete. |
| AI parsing returns low confidence | Add merchant/keyword to the dictionary in `AIParser.gs::CATEGORY_KEYWORDS`, **or** configure the OpenAI fallback. |
| Daily reminders not firing | *Setup → Install Daily Triggers* — required once per script project. |
