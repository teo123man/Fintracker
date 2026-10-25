# Harcamalarım — personal spending tracker

## Privacy rule (the user's explicit wish — overrides convenience)
The user does NOT want their financial data sent to Anthropic. Claude only changes the code.
- The real data lives OUTSIDE this folder, in the path written in the local, untracked `home.txt`:
  `data/`, `Belgeler/` with the statements, `Harcamalarim.html/.xlsx`). Never open, read, list, grep, print,
  screenshot or summarise anything there, or wherever `HARCAMA_HOME` points. This applies to every
  tool (Read, Grep, Bash/PowerShell `cat`/`Get-Content`, python one-liners that print rows, the browser pane).
- Develop and verify against the fake dataset instead:
  - `python demo.py` → rebuilds `demo/` (made-up statements, relative to today) and `demo/Harcamalarim.html`
  - `python app.py --demo` (launch config `harcama-demo`, port 8765) → the app on demo data
  - after template-only changes: `HARCAMA_HOME=demo python harcama.py --build --no-recalc`
- If a change must be checked on the real data, run the pipeline and report only pass/fail, error types and counts —
  never amounts, merchants, names or dates. Don't load the real dashboard in the browser pane.
- A new bank format needs a parser: ask the user for a redacted sample (or a description of the layout); don't ask
  for real statements.

## Layout
- `harcama.py` — parsers (Garanti/QNB card PDFs, Garanti/Ziraat .xlsx + QNB .pdf account exports), categorisation,
  FX, the Excel ledger (`Harcamalarim.xlsx`), then calls `dashboard.py`. Source of truth: `data/*.csv`.
  Import = every .pdf/.xlsx under `<home>/Belgeler/` (recursive); type detected from content; `data/files.csv`
  registers each file by SHA-256 so nothing is parsed twice; overlapping account exports are de-duplicated per row.
- `dashboard.py` — builds the JSON payload; `dashboard_template.html` — the whole UI (vanilla JS + SVG, no libraries,
  works offline). Placeholders: `/*__DATA__*/null` (payload, at build) and `/*__APP__*/null` (token, filled by app.py).
- `app.py` — local server on 127.0.0.1 only; token + Host check on every POST. Import-the-folder button, open folder,
  file list, drop-to-copy, categorise (rule or override), budgets. `Harcamalarim.bat` starts it.
- `demo.py` — the fake home.
- `data/confirmed_no_activity.csv` — optional explicit user confirmations, with columns
  `source,start,end,note` (ISO dates, inclusive; source must exactly match an imported bank account).
  These fill coverage gaps without creating transactions or changing statement dates. Never infer inactivity
  merely because an export is missing. The dashboard Data page and workbook show saved confirmations.
- `data/excluded_coverage_sources.csv` — optional user-directed exclusions from completeness, with
  columns `source,note`. Source must exactly match an imported bank account. Historical transactions
  remain included; no statement dates are changed. Exclusions appear in Data and the workbook.

## Dashboard rules from the user
- No all-time totals; everything follows a month-only period picker. Nothing before the first card statement month.
- Colourful, not monotone; no "Claude blue". Category colours live in `CAT_COLORS` in the template.
