"""Accounting model constants: chart of accounts, transaction types, lists, capacities.

This module is the single source of truth for *defaults*. The workbook copies
these into editable sheets (COA, LISTS); formulas then read the sheets, never
these Python values. See docs/accounting-model.md for the reasoning.
"""

# --- Capacities -------------------------------------------------------------
# Pre-sized table rows. The JOURNAL stacks all three sources, so its size is the
# sum. Increasing a capacity is a one-line change here + a rebuild.
TX_ROWS = 2000
INV_ROWS = 500
BILL_ROWS = 500
JOURNAL_ROWS = TX_ROWS + INV_ROWS + BILL_ROWS
COA_ROWS = 150
CUSTOMER_ROWS = 300
VENDOR_ROWS = 300
LEDGER_LINES = 400

# Ledger keys pack (date serial * KEY_MULT + journal row); must exceed JOURNAL_ROWS.
KEY_MULT = 10000
assert JOURNAL_ROWS < KEY_MULT

# --- Account types ------------------------------------------------------------
# name, statement, display sign (+1 = debit-positive), default normal balance, section label
ACCOUNT_TYPES = [
    ("Current Asset", "BS", 1, "Debit", "Current Assets"),
    ("Fixed Asset", "BS", 1, "Debit", "Fixed Assets"),
    ("Current Liability", "BS", -1, "Credit", "Current Liabilities"),
    ("Long-term Liability", "BS", -1, "Credit", "Long-term Liabilities"),
    ("Equity", "BS", -1, "Credit", "Equity"),
    ("Revenue", "PL", -1, "Credit", "Revenue"),
    ("Cost of Sales", "PL", 1, "Debit", "Cost of Sales"),
    ("Operating Expense", "PL", 1, "Debit", "Operating Expenses"),
    ("Other Income", "PL", -1, "Credit", "Other Income"),
    ("Other Expense", "PL", 1, "Debit", "Other Expenses"),
]
EXPENSE_TYPES = ("Cost of Sales", "Operating Expense", "Other Expense")
CF_SECTIONS = ["Cash", "Operating", "Investing", "Financing"]

# --- Default chart of accounts -----------------------------------------------
# code, name, type, normal, cash-flow section, payment account?, description
DEFAULT_ACCOUNTS = [
    (1000, "Cash on Hand", "Current Asset", "Debit", "Cash", "Yes", "Physical cash / cash register / petty cash"),
    (1010, "Bank Account", "Current Asset", "Debit", "Cash", "Yes", "Main business bank account"),
    (1020, "E-Wallet / Online Payments", "Current Asset", "Debit", "Cash", "Yes", "GCash, PayPal, Stripe, Maya, etc."),
    (1200, "Accounts Receivable", "Current Asset", "Debit", "Operating", "No", "Money customers owe you (from AR invoices)"),
    (1300, "Inventory", "Current Asset", "Debit", "Operating", "No", "Goods held for sale"),
    (1400, "Other Current Assets", "Current Asset", "Debit", "Operating", "No", "Deposits, prepaid rent, advances"),
    (1500, "Equipment", "Fixed Asset", "Debit", "Investing", "No", "Machines, computers, furniture"),
    (1510, "Vehicles", "Fixed Asset", "Debit", "Investing", "No", "Cars, motorcycles, delivery vans"),
    (1590, "Accumulated Depreciation", "Fixed Asset", "Credit", "Investing", "No", "Wear and tear of fixed assets (negative)"),
    (2000, "Accounts Payable", "Current Liability", "Credit", "Operating", "No", "Money you owe suppliers (from AP bills)"),
    (2100, "Credit Card", "Current Liability", "Credit", "Operating", "Yes", "Business credit card balance"),
    (2200, "Tax Payable", "Current Liability", "Credit", "Operating", "No", "VAT / GST / sales tax collected minus tax paid"),
    (2300, "Other Current Liabilities", "Current Liability", "Credit", "Operating", "No", "Customer deposits, accrued expenses"),
    (2500, "Loans Payable", "Long-term Liability", "Credit", "Financing", "No", "Bank or private loans"),
    (3000, "Owner Capital", "Equity", "Credit", "Financing", "No", "Money the owner put into the business"),
    (3100, "Owner Drawings", "Equity", "Debit", "Financing", "No", "Money the owner took out for personal use"),
    (3200, "Retained Earnings", "Equity", "Credit", "Financing", "No", "Profits kept from earlier years"),
    (3900, "Opening Balance Equity", "Equity", "Credit", "Financing", "No", "Automatic balancing figure for opening balances"),
    (4000, "Product Sales", "Revenue", "Credit", "Operating", "No", "Sales of goods"),
    (4100, "Service Income", "Revenue", "Credit", "Operating", "No", "Fees for services"),
    (4900, "Refunds & Discounts", "Revenue", "Debit", "Operating", "No", "Refunds and discounts given to customers (reduces revenue)"),
    (5000, "Cost of Goods Sold", "Cost of Sales", "Debit", "Operating", "No", "Purchase cost of goods you sold"),
    (5100, "Direct Materials", "Cost of Sales", "Debit", "Operating", "No", "Ingredients, raw materials"),
    (5200, "Direct Labor", "Cost of Sales", "Debit", "Operating", "No", "Wages of staff who make the product / deliver the service"),
    (5300, "Other Direct Costs", "Cost of Sales", "Debit", "Operating", "No", "Packaging, shipping to customer, platform fees per sale"),
    (6000, "Rent", "Operating Expense", "Debit", "Operating", "No", "Shop, office or storage rent"),
    (6100, "Utilities", "Operating Expense", "Debit", "Operating", "No", "Electricity, water, gas"),
    (6150, "Internet & Phone", "Operating Expense", "Debit", "Operating", "No", "Internet, mobile load, landline"),
    (6200, "Salaries & Wages", "Operating Expense", "Debit", "Operating", "No", "Office / admin staff pay"),
    (6300, "Advertising & Marketing", "Operating Expense", "Debit", "Operating", "No", "Ads, flyers, boosted posts"),
    (6400, "Transportation & Delivery", "Operating Expense", "Debit", "Operating", "No", "Fuel, fares, couriers"),
    (6500, "Office Supplies", "Operating Expense", "Debit", "Operating", "No", "Paper, ink, small tools"),
    (6550, "Software & Subscriptions", "Operating Expense", "Debit", "Operating", "No", "Apps, SaaS, website hosting"),
    (6600, "Bank & Payment Fees", "Operating Expense", "Debit", "Operating", "No", "Bank charges, card and e-wallet fees"),
    (6700, "Professional Fees", "Operating Expense", "Debit", "Operating", "No", "Accountant, lawyer, consultants"),
    (6800, "Repairs & Maintenance", "Operating Expense", "Debit", "Operating", "No", "Fixing equipment and premises"),
    (6850, "Insurance", "Operating Expense", "Debit", "Operating", "No", "Business insurance premiums"),
    (6870, "Taxes & Licenses", "Operating Expense", "Debit", "Operating", "No", "Permits, business licenses, non-income taxes"),
    (6880, "Depreciation", "Operating Expense", "Debit", "Operating", "No", "Wear and tear of equipment (non-cash)"),
    (6990, "Miscellaneous Expense", "Operating Expense", "Debit", "Operating", "No", "Small items that fit nowhere else"),
    (7000, "Interest Income", "Other Income", "Credit", "Operating", "No", "Bank interest earned"),
    (7100, "Other Income", "Other Income", "Credit", "Operating", "No", "Income outside your main business"),
    (8000, "Interest Expense", "Other Expense", "Debit", "Operating", "No", "Interest paid on loans"),
    (8100, "Other Expenses", "Other Expense", "Debit", "Operating", "No", "Penalties, losses, unusual costs"),
]

# System roles: (named range, label, default account name)
SYSTEM_ACCOUNTS = [
    ("acctAR", "Accounts Receivable account", "Accounts Receivable"),
    ("acctAP", "Accounts Payable account", "Accounts Payable"),
    ("acctTax", "Tax (VAT / sales tax) account", "Tax Payable"),
    ("acctCapital", "Owner Capital account", "Owner Capital"),
    ("acctDrawings", "Owner Drawings account", "Owner Drawings"),
    ("acctOBE", "Opening Balance Equity account", "Opening Balance Equity"),
]

# --- Transaction types ----------------------------------------------------------
# type, direction, other side ("Category" = the Category column, else a system role),
# tax allowed, help text
# Direction "In": Debit = Paid From/Deposited To, Credit = other side.
# Direction "Out": Debit = other side, Credit = Paid From/Deposited To.
TRANSACTION_TYPES = [
    ("Money In", "In", "Category", "Yes", "Any money received: cash sales, service fees, loans received, refunds received."),
    ("Money Out", "Out", "Category", "Yes", "Any money paid: expenses, purchases, equipment, loan repayments."),
    ("Customer Payment", "In", "acctAR", "No", "A customer pays an invoice from the AR sheet. Put the invoice # in Reference."),
    ("Bill Payment", "Out", "acctAP", "No", "You pay a supplier bill from the AP sheet. Put the bill # in Reference."),
    ("Transfer", "Out", "Category", "No", "Move money between your own accounts. Category = account receiving the money."),
    ("Owner Investment", "In", "acctCapital", "No", "Owner puts personal money into the business."),
    ("Owner Draw", "Out", "acctDrawings", "No", "Owner takes money out for personal use."),
    ("Adjustment", "Out", "Category", "No", "Advanced: Category is debited, Paid From is credited. For corrections and depreciation."),
]

PAYMENT_METHODS = ["Cash", "Bank Transfer", "Check", "Debit Card", "Credit Card", "E-Wallet", "Online Payment", "Other"]
BUSINESS_TYPES = [
    "Retail Shop", "Online Seller", "Food & Beverage", "Services", "Freelancer",
    "Contractor", "Agency", "Manufacturing", "Other",
]
MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]
YES_NO = ["Yes", "No"]

AGING_BUCKETS = ["Current", "1-30", "31-60", "61-90", "90+"]
DUE_SOON_DAYS = 14
DEFAULT_TERMS_DAYS = 30

# P&L / balance-sheet detail slots per section (totals never depend on slots).
PL_SLOTS = {"Revenue": 6, "Cost of Sales": 6, "Operating Expense": 22, "Other Income": 4, "Other Expense": 4}
BS_SLOTS = {"Current Asset": 10, "Fixed Asset": 6, "Current Liability": 8, "Long-term Liability": 4, "Equity": 6}
CASH_SLOTS = 6
TOP_N = 5
