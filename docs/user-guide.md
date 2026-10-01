# SimpleLedger — User Guide (v1.0.0)

_Know where your money goes, know how much you earn, and understand your business finances from one simple Excel workbook._

> Generated from the workbook's HELP sheet by `tools/make_user_guide.py` — edit `src/simpleledger/ui.py`, not this file.

## The sheets at a glance

| Sheet | What it's for |
|---|---|
| START_HERE | Welcome page, workflow and buttons to every sheet |
| SETUP | Your business details, start date, currency, opening cash and bank, tax settings |
| DASHBOARD | Pick the report month; KPIs, data health, charts |
| TRANSACTIONS | Every sale, expense, payment and transfer (up to 2,000 rows) |
| AR / AP | Customer invoices you'll be paid for later / supplier bills you'll pay later |
| CUSTOMERS / VENDORS | Your customers and suppliers, with what they owe / you owe |
| COA | Chart of accounts: your categories and opening balances |
| PNL, BALANCE_SHEET, CASH_FLOW, MONTHLY_SUMMARY | Reports (automatic, print-ready) |
| TRIAL_BALANCE, LEDGER | For bookkeepers: balances per account; every posting to one account |
| HELP | This guide |
| LISTS, JOURNAL, CALC | Settings and the calculation engine — look, don't touch |

## Transaction types

| Type | Use it for |
|---|---|
| Money In | Any money received: cash sales, service fees, loans received, refunds received. |
| Money Out | Any money paid: expenses, purchases, equipment, loan repayments. |
| Customer Payment | A customer pays an invoice from the AR sheet. Put the invoice # in Reference. |
| Bill Payment | You pay a supplier bill from the AP sheet. Put the bill # in Reference. |
| Transfer | Move money between your own accounts. Category = account receiving the money. |
| Owner Investment | Owner puts personal money into the business. |
| Owner Draw | Owner takes money out for personal use. |
| Adjustment | Advanced: Category is debited, Paid From is credited. For corrections and depreciation. |

## How to start

- 1. Open SETUP and fill in the yellow cells: business name, accounting start date, currency, and how much cash and bank money you had on the start date.
- 2. Optional: on the COA (Chart of Accounts) sheet, enter other opening balances — e.g. a loan you already owe, or equipment you already own. The 'Opening Balance Equity' account balances them automatically.
- 3. Add your regular customers (CUSTOMERS) and suppliers (VENDORS).
- 4. Start recording on TRANSACTIONS. Pick the month on the DASHBOARD to see results.

## How to record income (a sale paid immediately)

- TRANSACTIONS → Type: Money In · Category: Product Sales (or Service Income) · Paid From / Deposited To: where the money went (Cash on Hand, Bank Account, E-Wallet) · Amount: total received.
- If you track VAT / sales tax (SETUP), type the tax portion in Tax. Example: 1,120 received incl. 120 VAT → Amount 1,120, Tax 120. Revenue is 1,000; 120 goes to Tax Payable.

## How to record a sale on credit (customer pays later)

- AR sheet → new row: Invoice #, date, customer, revenue account, amount. Leave Due Date blank to use the customer's payment terms.
- When the customer pays: TRANSACTIONS → Type: Customer Payment · Customer · Reference: the invoice # · Deposited To: Bank/Cash · Amount paid. Leave Category blank. Partial payments are fine — record each one.

## How to record expenses

- Paid now: TRANSACTIONS → Type: Money Out · Category: the expense (Rent, Utilities …) · Paid From: Cash, Bank or Credit Card · Amount.
- Paid later (supplier bill): AP sheet → bill #, date, vendor, expense account, amount. When you pay it: TRANSACTIONS → Type: Bill Payment · Vendor · Reference: bill # · Paid From · Amount.
- Buying equipment: Money Out with Category = Equipment. It goes to the Balance Sheet, not the P&L.
- Paying back a loan: Money Out with Category = Loans Payable (interest part: Interest Expense, on a separate row).

## How to record transfers and owner money

- Moving money between your own accounts (e.g. depositing cash to the bank, paying the credit card from the bank): Type: Transfer · Paid From: where it left · Category: where it arrived.
- Owner puts in personal money: Type: Owner Investment · Deposited To: Bank/Cash. Category is automatic.
- Owner takes money for personal use: Type: Owner Draw · Paid From: Bank/Cash. This is not an expense.
- Loan received: Money In · Category: Loans Payable.

## How to add customers, suppliers and accounts

- Type a new row at the bottom of CUSTOMERS, VENDORS or COA. Dropdowns update automatically.
- New account? Give it a unique name, a type (e.g. Operating Expense), Normal Balance and Cash Flow Section. It appears on the reports automatically.
- Stopped using an account? Set Active? = No to hide it from dropdowns. Don't delete accounts that have transactions.

## How to read the Profit & Loss

- Revenue − Cost of Sales = Gross Profit (what you keep after the direct cost of what you sold).
- Gross Profit − Operating Expenses = Net Operating Income. ± Other income/expenses = NET PROFIT.
- Columns: this month, previous month, year to date, full year, and each month of the fiscal year.
- Profit is not the same as cash: credit sales count as revenue before the money arrives; equipment purchases use cash but are not expenses.

## How to read the Balance Sheet

- What the business owns (Assets) = what it owes (Liabilities) + the owner's stake (Equity).
- Equity includes money you put in, minus drawings, plus all profits so far.
- The status line must say ✓ Balanced. If not, see 'How to fix warnings'.

## How to read the Cash Flow

- Shows where cash came from and went, in three groups: Operating (day-to-day business), Investing (equipment and assets), Financing (owners and loans).
- Cash at the end = cash at the start + net change. It must match the cash on the Balance Sheet.
- Transfers between your own cash/bank accounts are not inflows or outflows — they cancel out.

## How to fix warnings

- ⛔ on a row: that row is NOT included in reports. Read the message in the Check column — e.g. 'Missing transaction date', 'Select an account', 'Enter an amount' — and fix the white cells on that row.
- ⚠ on a row: included, but unusual — e.g. a negative amount, a Customer Payment whose invoice # can't be found, or a duplicate ID. Fix it if it's a mistake.
- Receivables/Payables register differs from the ledger: a Customer/Bill Payment probably has the wrong Reference, or someone used Category = Accounts Receivable/Payable on a normal Money In/Out row.
- Trial balance 'CHECK DATA' or Balance Sheet out of balance: a grey formula cell was overwritten. Copy the formula from the row above, or restore from your backup.

## What not to edit

- Grey cells and grey columns; the JOURNAL and CALC sheets (the engine); the report sheets.
- Don't rename an account that already has transactions unless you also update those rows (the Check column will show 'Unknown account').
- Don't insert columns inside the input tables. Adding rows at the bottom is fine up to the capacity.

## Frequently asked questions

- Do I need accounting knowledge? No. Pick a Type and a Category; the workbook does the debits and credits.
- Does it calculate my income tax? No. It produces the profit figure your accountant or tax form needs. VAT/sales tax collected and paid is tracked in Tax Payable if you switch it on.
- Can I use it for several businesses? Use one copy of the workbook per business.
- Does it work in Google Sheets? It is built and supported for Microsoft Excel. See the product page for the current compatibility list.
- Are there macros? No. It is a plain .xlsx file — safe to open, no security prompts.
- Can I print reports? Yes. Every report sheet is set up to print or Save as PDF on standard paper.
- A new fiscal year? Start a fresh copy, enter your closing balances from the old Balance Sheet as opening balances (cash and bank on SETUP, the rest on COA).
