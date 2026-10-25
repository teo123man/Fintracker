# Development and privacy

- Develop and verify only with the synthetic dataset: `python demo.py`, then `python app.py --demo`.
- Never read, list, search, print or display real financial data, statements or generated personal dashboards. This includes the private directory selected by `home.txt` or `HARCAMA_HOME`.
- Keep `home.txt`, `coverage_sources.txt`, local assistant settings, financial documents and generated outputs out of Git.
- Do not put real data in `demo/`: `python demo.py` deletes and recreates it.
- Run `python -m unittest discover -v` with `HARCAMA_HOME=demo`, then `python demo.py` for pipeline verification.
- New bank parsers require a redacted sample or a format description, never real statements.
- Preserve explicit coverage confirmations and exclusions; do not infer inactivity from missing exports.
- See `CLAUDE.md` for existing dashboard behavior and architecture details.
