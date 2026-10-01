# Changelog

Semantic versioning: **major** = workbook redesign or breaking compatibility · **minor** = new features/reports · **patch** = formula/formatting fixes.

## [1.0.0] — 2026-10-01 (release candidate, not yet published)

### Added
- Workbook generator (`src/simpleledger/`) producing 19 sheets: START_HERE, SETUP, DASHBOARD, TRANSACTIONS, AR, AP, CUSTOMERS, VENDORS, COA, PNL, BALANCE_SHEET, CASH_FLOW, MONTHLY_SUMMARY, TRIAL_BALANCE, LEDGER, HELP, LISTS, JOURNAL, CALC.
- Double-entry engine behind a non-accountant interface: 8 transaction types; AR/AP registers with partial payments, terms, aging; optional VAT/GST; opening balances with automatic Opening Balance Equity; any fiscal-year start month.
- Row-level validation (⛔ errors excluded from the books, ⚠ warnings) and a Dashboard data-health panel (trial balance, balance sheet, cash-flow reconciliation, AR/AP register vs ledger, capacity).
- Dashboard: 8 KPI cards, 4 charts, cash summary, due-soon and top debtors/creditors.
- Sample business "Sunrise Bakery & Café" (deterministic generator).
- Test suite: independent reference ledger, 35 hand-written scenarios, validation messages, fiscal and calendar boundaries, 3 randomized books; all recalculated in LibreOffice.
- CI workflow; packaging script; user guide generated from the in-workbook HELP.
- Product, accounting, formula, QA, market-research, positioning, launch and support docs; marketing drafts.

### Fixed (during QA)
- Account Ledger: the closing-balance cell was overwritten by the entries header row (found by the reference comparison).
- Dashboard: charts overlapped each other and the chart-data block after switching to Arial; resized and shortened titles.
- Sample data: added quarterly VAT remittances so Tax Payable is realistic.
- Builder: text typed in a date field is kept as text (and flagged "Invalid date") instead of crashing the build.
