# Product Specification — SimpleLedger v1.0 (Standard)

> **SimpleLedger: a simple small-business finance system in Excel.**
> Know where your money goes, know how much you earn, and understand your business finances from one simple Excel workbook.

Working product name: **SimpleLedger** (trademark and domain check still pending; see RUNBOOK).

## 1. Personas

| Persona | Situation | Job to be done | Success looks like |
|---|---|---|---|
| **Maria, micro-business owner** (bakery, shop, online seller) | Records in a notebook or chat; mixes personal and business money | "Tell me if I'm actually making money and where it goes" | Records a sale in < 30 s; reads monthly profit without help |
| **Jon, freelancer / contractor** | Invoices clients, chases late payments | "Who owes me what, and what's overdue?" | Sees overdue invoices and aging at a glance |
| **Liza, bookkeeper / VA** | Keeps books for 3–10 micro clients who won't pay for software | "A consistent, auditable file per client" | Trial balance and ledger she trusts; one workbook per client |

## 2. User stories (MVP)

1. As an owner, I want to record an expense in under 30 seconds, so my records stay current without accounting software.
2. As an owner, I want to pick a type and category instead of thinking in debits and credits.
3. As an owner, I want monthly and year-to-date profit, so I know whether the business is working.
4. As an owner, I want to see my cash at month end and why it changed.
5. As an owner, I want to know who owes me money and which invoices are overdue.
6. As an owner, I want to know which supplier bills are due soon.
7. As an owner, I want the workbook to tell me when I made a mistake, and never silently report wrong numbers.
8. As a bookkeeper, I want a trial balance, account ledger and balance sheet that always balance.
9. As an owner, I want print-ready reports to give my accountant or bank.
10. As a VAT-registered business, I want VAT collected and paid tracked separately from my income.

## 3. Workbook sheets

| # | Sheet | Purpose | Editable? |
|---|---|---|---|
| 1 | START_HERE | Welcome, workflow, navigation, colour guide | No |
| 2 | SETUP | Business details, start date, fiscal year, currency, opening cash/bank, payment terms, tax settings; shows the calculated report period | Yellow cells |
| 3 | DASHBOARD | Month/year selector, 8 KPI cards, data health, 4 charts, cash summary, who-owes-whom | Month + year |
| 4 | TRANSACTIONS | `tblTransactions`: 2,000 rows | White columns |
| 5 | AR | `tblInvoices`: 500 customer invoices with paid/balance/status/aging | White columns |
| 6 | AP | `tblBills`: 500 supplier bills with paid/balance/status/aging | White columns |
| 7 | CUSTOMERS | `tblCustomers`: 300 rows with invoiced/paid/outstanding/overdue | White columns |
| 8 | VENDORS | `tblVendors`: 300 rows, same as CUSTOMERS | White columns |
| 9 | COA | `tblAccounts`: 44 default accounts, room for 150 | White columns |
| 10 | PNL | Profit & Loss: month, previous month, YTD, full year, 12 monthly columns | No |
| 11 | BALANCE_SHEET | As of report date and previous month end, with balance check | No |
| 12 | CASH_FLOW | Direct-method cash flow, month and YTD, cash by account, reconciliation | No |
| 13 | MONTHLY_SUMMARY | 12-month table; report month highlighted | No |
| 14 | TRIAL_BALANCE | Debits/credits per account, BALANCED / CHECK DATA | No |
| 15 | LEDGER | Pick an account and date range → every posting with running balance | Account + dates |
| 16 | HELP | How-to, reading reports, fixing warnings, FAQ | No |
| 17 | LISTS | Transaction-type rules, system accounts, dropdown lists | Yellow cells |
| 18 | JOURNAL | Engine: every posting stacked (visible, protected) | No |
| 19 | CALC | Engine: balances per account and fiscal month (visible, protected) | No |

Field-level definitions live in `docs/accounting-model.md` (rules) and `docs/formula-spec.md` (implementation).

## 4. Dashboard

KPI cards: revenue, expenses and net profit (each this month, with YTD), cash at month end (with change), receivables (with overdue), payables (with overdue), gross margin and net margin (YTD, with this month).

Data health: trial balance, balance sheet, cash-flow reconciliation, entries not posted, entries with warnings, AR register vs ledger, AP register vs ledger, setup/capacity.

Charts: revenue vs expenses by month, net profit by month, top 5 expense categories (YTD), receivables and payables aging.

Tables: cash flow this month, receivables and payables due in the next 14 days, top 3 customers owing you and top 3 vendors you owe.

## 5. Validation and error prevention

- Dropdowns for type, category, customer, vendor, payment method, money account, account type, yes/no, month.
- Every register row has a **Check** column (⛔ error / ⚠ warning / ✓ OK); see `accounting-model.md §4`.
- ⛔ rows are excluded from the books, and the Dashboard counts them.
- Balance checks on the Trial Balance, Balance Sheet and Cash Flow, plus AR/AP reconciliation to the ledger.

## 6. UX rules

- One font (Arial), one palette (navy / teal / slate). Yellow = settings, white table columns = data entry, grey = automatic.
- Data-entry tables are **not** protected, so sorting, filtering and adding rows work. Report, engine and settings sheets are protected **without a password**.
- Frozen headers, report print areas set to landscape/portrait fit-to-width, gridlines hidden on report sheets.
- Helper columns are grouped and collapsed, never hidden sheets.

## 7. Compatibility requirements

- `.xlsx`, **no macros**, no external links, no Power Query.
- Formulas are limited to functions available since Excel 2010: SUMIFS, COUNTIFS, INDEX/MATCH, SMALL/LARGE, SUMPRODUCT, EDATE/EOMONTH, IFERROR. No XLOOKUP, FILTER, UNIQUE, LET, LAMBDA, OFFSET or INDIRECT, which the test suite enforces. **Why:** maximum reach (Excel 2010–365, Mac, and better odds with Google Sheets/LibreOffice) at no loss of function.
- Defined names over A1 ranges instead of structured references inside formulas. They behave identically in Excel and are more portable.
- **Verified:** LibreOffice Calc 24.2 (automated, every build). **Not yet verified:** Excel 365 Windows/Mac, Excel Online, Google Sheets. This is the Phase 10 release gate. Do not claim support until tested.

## 8. Version scope

| In v1.0 Standard | Later (see roadmap) |
|---|---|
| Everything above | Budget vs actual, forecast, inventory, fixed-asset register/depreciation schedule, multi-currency, multi-business, CSV/bank import, recurring transactions |

Capacities in v1.0: 2,000 transactions, 500 invoices, 500 bills, 150 accounts, 300 customers, 300 vendors per workbook. That is one fiscal year for a typical micro business. Capacities are constants in `src/simpleledger/model.py`.

## 9. Tiers (proposed, Phase 13)

- **Free / lead magnet:** income & expense tracker + mini dashboard (to build from this generator by switching sheets off).
- **Standard:** this specification.
- **Premium:** Standard + budget vs actual, forecast, inventory, ratios/KPI scorecard.
