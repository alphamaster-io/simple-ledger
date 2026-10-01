# QA Plan

## 1. Strategy

The workbook is generated from code, so it is tested like code:

1. **Build** each data set into an `.xlsx` (`simpleledger.build`).
2. **Recalculate** it headlessly in LibreOffice Calc (`tools/recalc.py`). openpyxl stores formulas without values.
3. **Compare** every report figure with an **independent Python implementation** of the accounting model (`tests/reference.py`), plus hand-written expectations.

Two independent implementations (Excel formulas and Python) must agree to the cent on every figure.

Run it: `pip install -e ".[dev]" && python -m pytest`. This needs `libreoffice-calc`, and takes about 4 minutes. CI runs the same on every push (`.github/workflows/ci.yml`).

## 2. Automated test inventory

| Module | What it proves | Count |
|---|---|---|
| `test_structure.py` | Sheet order, no hidden sheets, no macros/external links, table names & capacities, defined names resolve, **no Excel-365-only/volatile functions**, protection (no password) on reports, unlocked inputs, dropdowns, default COA, recalc-on-open, 4 charts | 13 |
| `test_reports.py` | For **12 data sets**: zero formula errors; TB balanced + journal integrity; Assets = L + E; cash flow reconciles; **every P&L line × 16 columns, BS, CF, TB, AR/AP per document (paid/balance/status/bucket/due), dashboard KPIs, aging, monthly summary, ledger opening/closing = reference model** | 38 |
| `test_scenarios.py` | 35 hand-written business scenarios post exactly the expected debit/credit/tax entry; every entry balances; invoice/bill status after partial payments; VAT payable = output − input − remitted; tax-off mode | 39 |
| `test_validation.py` | 26 transaction, 10 invoice and 3 bill input mistakes → exact ⛔/⚠ message; ⛔ rows never posted; Dashboard counts errors; duplicate account names flagged, not double-counted | 42 |
| `test_edge_cases.py` | Prior-year entries → prior-year profit; Jan 1 / Jun 30 / Jul 1 / Dec 31 boundaries; future-dated entries excluded from balances; 987,654,321.99 and 0.01 exact; added account appears; inactive account hidden from dropdown; opening-balance equity; fiscal year starting in July; first fiscal month uses prior-year previous month; single transaction; ledger date order | 10 |

**Data sets:** blank · sample (Sunrise Bakery, 345 transactions, 22 invoices, 45 bills) · scenarios · validation · boundaries · fy_july · fy_july_first_month · tax_off · one_transaction · random_1..3 (about 140 random rows each, including deliberately broken rows, across two fiscal years with random fiscal-year start and report month).

Edge cases from the master plan (§31) map to these: zero transactions → blank; one transaction; same-day → ledger order & sample; backdated → boundaries (before start date and before FY); duplicate references → validation; large and decimal amounts → boundaries; negative adjustments → scenario 30; missing fields → validation; deleted transactions → empty row in boundaries; added accounts/categories → boundaries; month-end & year-end → boundaries/fy_july.

## 3. Release checklist (automated part)

- [x] Opens without errors in LibreOffice; all formulas calculate (0 error values)
- [x] No broken references; no external links; no macros; no hidden sheets
- [x] No test data in the production (blank) workbook
- [x] Trial balance balances; balance sheet balances; cash flow reconciles
- [x] P&L, AR, AP, opening balances, month-end and year-end boundaries verified against the reference model

## 4. Manual compatibility pass (Phase 10, not done yet)

For each of Excel 365 Windows, Excel for Mac, Excel Online, and Google Sheets (only to decide whether to claim support):

1. Open the blank and sample files: no repair prompt, no `#NAME?`, values match the LibreOffice-calculated sample (Dashboard KPIs: revenue Sep 326,522.88; net profit YTD 561,097.18; cash at month end 660,779.77 at the time of writing; re-check against the test output).
2. Dropdowns work (Type, Category, Paid From, Customer, Vendor, Dashboard month).
3. Add a transaction at the next free row: Check shows ✓, Dashboard updates.
4. Sort and filter TRANSACTIONS by date: reports unchanged.
5. Delete a transaction row: no `#REF!`; totals drop accordingly.
6. Protected sheets block typing in grey cells; Dashboard month/year still editable.
7. Charts render; print preview of PNL, BALANCE_SHEET, CASH_FLOW fits the page width.
8. Calculation speed acceptable with the sample file (< 2 s per edit on a mid-range laptop).

## 5. UAT script (Phase 12)

Testers: one small-business owner, one bookkeeper, one ordinary Excel user. No help given. Record completion, time, mistakes, questions and confusing screens.

1. Set up the business (SETUP). 2. Add a customer. 3. Add a vendor. 4. Record a cash sale. 5. Record an expense. 6. Record a credit sale (AR) and its payment. 7. Open the Dashboard for this month. 8. State this month's profit. 9. Find which customers owe money. 10. Find unpaid supplier bills.

Pass: ≥ 2 of 3 testers complete all 10 tasks, each core entry in < 60 s, no tester leaves bad data undetected.
