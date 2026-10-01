# CLAUDE.md: operating rules for this repository

`RUNBOOK.md` is the controlling execution document. At the start of every session or phase:

1. Read this file, then `RUNBOOK.md` (current phase, open items, known issues).
2. Read the relevant spec: `docs/accounting-model.md` for any accounting change, `docs/formula-spec.md` for workbook mechanics.
3. Inspect the repo (`git log`, `git status`), then continue from the first incomplete task. Never redo completed work.

After every phase: validate, update docs, update `RUNBOOK.md` (status, tests run, known issues, next step) and `CHANGELOG.md`, then commit.

A phase is complete only when the implementation exists, tests pass, docs are updated, acceptance criteria are met, and known issues are recorded.

## Engineering rules

- **The workbook is generated.** Never hand-edit `dist/*.xlsx`. Change `src/simpleledger/`, rebuild (`python -m simpleledger.build`), and commit the rebuilt files.
- **Accounting rule changes touch three places together:** `docs/accounting-model.md`, `tests/reference.py`, and the generator. Add or adjust a scenario in `tests/scenarios.py`.
- **Formulas:** Excel-2010 functions only (SUMIFS, COUNTIFS, INDEX/MATCH, SMALL/LARGE, SUMPRODUCT, EDATE/EOMONTH, IFERROR). No XLOOKUP/FILTER/UNIQUE/SORT/SEQUENCE/LET/LAMBDA/OFFSET/INDIRECT (a test enforces this). No VBA. No hard-coded report numbers.
- Use defined names for ranges and settings. Coerce lookup keys with `&""`. Reference source tables from the JOURNAL with `INDEX(range, n)`, never direct links.
- Totals never depend on report slots. They come from `SUMIFS` over CALC by type.
- Keep the user's world simple: Type + Category, ⛔/⚠/✓ messages in plain language.
- `src/sample-data/sunrise_bakery.json` and `docs/user-guide.md` are generated (`tools/make_sample_data.py`, `tools/make_user_guide.py`). CI fails if they're stale.

## Testing

- `python -m pytest` (needs `libreoffice-calc`). About 4 min. Builds and recalculates 12 data sets with 2 parallel LibreOffice workers; more workers have exhausted memory in the cloud container.
- A green recalc only proves formulas *evaluate*. `tests/compare.py` proves they're *right* against `tests/reference.py`.
- Never claim Excel/Mac/Google Sheets compatibility that hasn't been tested by a person (see RUNBOOK Phase 10).
