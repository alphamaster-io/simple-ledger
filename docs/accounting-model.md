# Accounting Model

The rules the workbook implements. `tests/reference.py` re-implements this document independently in Python and every release is checked against it. If you change a rule, change this document, `tests/reference.py` and the workbook together.

## 1. Principles

1. **One source of truth.** Users type into three registers only: `TRANSACTIONS`, `AR` (customer invoices) and `AP` (supplier bills), plus master data (`SETUP`, `COA`, `CUSTOMERS`, `VENDORS`).
2. **Real double entry, hidden from the user.** Every valid source row becomes one balanced journal entry (debit = credit). The user never types "debit" or "credit". They pick a **Type** and a **Category**.
3. **Bad rows never reach the books.** A row with an error (⛔) is not posted. The row stays visible and the Dashboard counts it. Warnings (⚠) are posted.
4. **Reports never contain typed numbers.** Every figure is a formula over the journal.

```text
SETUP / COA / CUSTOMERS / VENDORS        (master data)
TRANSACTIONS  AR  AP                     (user input; each row computes its own posting)
        │      │   │
        └──────┴───┴──► JOURNAL           (all postings stacked, one row per source row)
                           │
                           ▼
                         CALC             (balance per account × period)
                           │
        ┌──────────┬───────┼───────────┬──────────────┬───────────┐
        ▼          ▼       ▼           ▼              ▼           ▼
       PNL   BALANCE_SHEET CASH_FLOW TRIAL_BALANCE MONTHLY_SUMMARY LEDGER
                           │
                           ▼
                       DASHBOARD
```

## 2. Chart of accounts

Each account has: Code, Name (the key), Type, Normal Balance, Cash Flow Section, Payment Account?, Active?, Opening Balance, Description.

| Type | Statement | Shown as | Default normal |
|---|---|---|---|
| Current Asset | Balance sheet | debit − credit | Debit |
| Fixed Asset | Balance sheet | debit − credit | Debit |
| Current Liability | Balance sheet | credit − debit | Credit |
| Long-term Liability | Balance sheet | credit − debit | Credit |
| Equity | Balance sheet | credit − debit | Credit |
| Revenue | P&L | credit − debit | Credit |
| Cost of Sales | P&L | debit − credit | Debit |
| Operating Expense | P&L | debit − credit | Debit |
| Other Income | P&L | credit − debit | Credit |
| Other Expense | P&L | debit − credit | Debit |

The statement sign comes from the **type**, not the normal balance. So contra accounts show naturally as negatives: Accumulated Depreciation under Fixed Assets, Owner Drawings under Equity, Refunds & Discounts under Revenue.

**Cash Flow Section**: `Cash` marks cash-equivalent accounts (Cash on Hand, Bank, E-Wallet). Every other account is `Operating`, `Investing` or `Financing`, and that classifies the cash that moves against it.

**Payment Account? = Yes** puts the account in the *Paid From / Deposited To* dropdown. The Credit Card account is a payment account but **not** cash (it is a liability). Paying an expense by card therefore moves no cash. Paying the card bill from the bank does.

**System accounts** (`LISTS` sheet, editable names): Accounts Receivable, Accounts Payable, Tax Payable, Owner Capital, Owner Drawings, Opening Balance Equity.

Duplicate account names are flagged ⛔ on `COA`. Only the first occurrence is used, so nothing is double-counted.

## 3. Opening balances

- Cash on Hand and Bank Account opening balances come from `SETUP`. Every other account's opening balance is typed on `COA`, as a positive number in its normal direction.
- Opening balances are not dated. They sit before every transaction.
- **Opening Balance Equity** gets the balancing figure automatically: `OBE = −Σ(signed openings of all other accounts)`. The opening position therefore always balances.
- An opening balance on a P&L account is allowed but flagged ⚠. It is treated as profit of earlier years.

## 4. Transaction types (TRANSACTIONS sheet)

Columns the user fills: Date, Type, Category, Description, Customer, Vendor, Payment Method, Reference, **Paid From / Deposited To** (the *money account*), Amount (incl. tax), Tax.

| Type | Debit | Credit | Tax allowed | Typical use |
|---|---|---|---|---|
| Money In | Money account | Category | Yes | Cash sales, fees, loan received, other income |
| Money Out | Category | Money account | Yes | Expenses, purchases, equipment, loan repayment, VAT remittance |
| Customer Payment | Money account | Accounts Receivable | No | Customer pays an AR invoice (Reference = invoice #) |
| Bill Payment | Accounts Payable | Money account | No | You pay an AP bill (Reference = bill #, same vendor) |
| Transfer | Category (to) | Money account (from) | No | Cash → bank, bank → credit card, e-wallet → bank |
| Owner Investment | Money account | Owner Capital | No | Owner puts money in |
| Owner Draw | Owner Drawings | Money account | No | Owner takes money out |
| Adjustment | Category | Money account column (any account) | No | Corrections, depreciation. Negative amounts allowed |

These rules live as data in the `LISTS` sheet ("Transaction types" table), not hard-coded in formulas.

### Tax (VAT / GST / sales tax)

Tax is only used when `SETUP → Track VAT / sales tax = Yes`, and only on Money In / Money Out (and on AR/AP documents). Amount is the **gross** amount, Tax is the tax portion, Net = Amount − Tax.

- Money In with tax: Dr money account *gross* · Cr category *net* · Cr Tax Payable *tax*. This is output tax.
- Money Out with tax: Dr category *net* · Dr Tax Payable *tax* · Cr money account *gross*. This is input tax, recoverable VAT style.
- Remitting tax to the authority is `Money Out` with Category = Tax Payable.
- If tax tracking is off, or the type doesn't allow tax, a typed Tax value is ignored and flagged ⚠.

Users whose purchase tax is **not** recoverable (most US sales-tax situations) should leave Tax blank on expenses. The tax then stays part of the cost.

### Validation (first matching rule wins)

Errors (⛔, row not posted): missing date · invalid date (text, or outside 2000–2099) · missing type · unknown type · missing category for category-based types · unknown account · missing money account · unknown money account · missing / non-numeric / zero amount · |tax| > |amount|.

Warnings (⚠, row posted): duplicate transaction ID · negative amount (except Adjustment) · Customer Payment without customer · invoice # not found on AR · Bill Payment without vendor · bill # + vendor not found on AP · tax ignored · same account on both sides · date before accounting start date · money account isn't a payment account (except Adjustment).

A row whose Date, Type, Category, Amount and Money account are all empty is blank: no message, nothing posted.

## 5. Accounts receivable (AR sheet)

One row per customer invoice: Invoice #, Date, Customer, Description, Revenue Account, Amount (gross), Tax, Due Date (optional).

- Posting: Dr Accounts Receivable *gross* · Cr revenue account *net* · Cr Tax Payable *tax*.
- **Due date used** = Due Date if typed, otherwise invoice date + the customer's payment terms, otherwise + the default terms from SETUP.
- **Paid** = sum of posted `Customer Payment` transactions whose Reference = this invoice # (case-insensitive), dated on or before the report date.
- **Balance** = Amount − Paid, or 0 if the invoice is dated after the report date.
- **Status**: Future (dated after the report date) → Overpaid (balance < 0) → Paid (balance = 0) → Overdue (report date > due date) → Partial (some paid) → Unpaid.
- **Days outstanding** = report date − invoice date. **Days overdue** = max(0, report date − due date). **Aging bucket**: Current, 1–30, 31–60, 61–90, 90+ days overdue.
- Errors: missing invoice #, missing/invalid date, missing/unknown account, non-positive amount, tax > amount. Warnings: duplicate invoice #, missing customer, overpaid, account not Revenue/Other Income, tax ignored, before start date.

## 6. Accounts payable (AP sheet)

The mirror image of AR, with one difference: payments match on **bill # and vendor**, because two suppliers can use the same bill number. So the vendor is required (⛔ if missing).

- Posting: Dr expense/asset account *net* · Dr Tax Payable *tax* · Cr Accounts Payable *gross*.

## 7. Periods

The report period is chosen on the **Dashboard** (month + year; blank means the current month).

- Report period = that calendar month. **As-of date** = its last day.
- Fiscal year starts in the month set on SETUP. The fiscal year containing the report month runs from `FYStart` to `FYEnd`.
- Each journal row gets a **fiscal period**: 0 = before FYStart, 1–12 = month of the fiscal year, 13 = after FYEnd.
- **Previous month** comes from the monthly column, or, when the report month is the first fiscal month, from explicit dates in the prior fiscal year.

## 8. Reports

- **Balance at as-of date** = opening + all activity in fiscal periods 0..report month. Entries dated after the report date are excluded (they show in "Full Year" P&L columns only).
- **Profit & Loss**: this month, previous month, year to date, full fiscal year, plus 12 monthly columns. Revenue − Cost of Sales = Gross Profit. − Operating Expenses = Net Operating Income. + Other Income − Other Expenses = **Net Profit**.
- **Balance Sheet**: Assets = Liabilities + Equity, where Equity = equity accounts + *profit from earlier years* (all P&L activity and openings before FYStart) + *current-year profit* (year to date). There are no closing entries, so a workbook can span several years. A status line shows "✓ Balanced" or "⚠ Balance Sheet is out of balance by …".
- **Cash Flow (direct method)**: for every posting with exactly one cash-type side, the cash effect is classified by the other account's Cash Flow Section. Tax portions are always Operating. Transfers between two cash accounts are ignored. Ending cash must equal Balance-Sheet cash (status line).
- **Trial Balance**: every account's balance at the as-of date in a debit or credit column. BALANCED requires total debits = total credits **and** the journal integrity sum (Σ debits − Σ credits of every posting) = 0.
- **Account Ledger**: every posting to one account between two dates (default: fiscal year start → report date), in date order, with opening balance and running balance in the account's normal direction.
- **Monthly Summary**: 12 fiscal months of revenue, costs, profit, cash in/out, and month-end cash, receivables and payables.

## 9. Sample scenarios

The 35 hand-written scenarios in `tests/scenarios.py` are the executable version of this section. They cover: cash sale, VAT sale, credit sale, full/partial customer payment, cash/bank/card expenses, supplier bill, full/partial bill payment, owner investment and draw, transfers, card repayment, equipment with VAT, loan received/repaid/interest, refund, discount, recurring rent, miscellaneous income, reclassification, depreciation, negative reversal, VAT remittance, decimals, inventory, e-wallet cash-out, and tax typed on a non-taxable type.
