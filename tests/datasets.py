"""Data sets used by the test suite. Each function returns a build() data dict."""

from __future__ import annotations

import datetime as dt
import json
import random
from pathlib import Path

import scenarios as SC

ROOT = Path(__file__).resolve().parents[1]
CASH, BANK, EW, CC = SC.CASH, SC.BANK, SC.EW, SC.CC


def blank():
    return {"setup": {}, "report": {"year": 2026, "month": "September"}}


def sample():
    return json.loads((ROOT / "src" / "sample-data" / "sunrise_bakery.json").read_text(encoding="utf-8"))


def scenarios():
    tx, inv, bills = [], [], []
    for _id, _title, src, rec, _exp, _chk in SC.SCENARIOS:
        {"tx": tx, "ar": inv, "ap": bills}[src].append(dict(rec))
    return {
        "setup": {"BizName": "Scenario Co", "StartDate": "2026-01-01", "OpenCash": 20000, "OpenBank": 100000,
                  "TaxOn": "Yes"},
        "report": {"year": 2026, "month": "March"},
        "customers": SC.CUSTOMERS, "vendors": SC.VENDORS,
        "transactions": tx, "invoices": inv, "bills": bills,
    }


# (record, expected message fragment in Check column) — "" means no message (blank row)
VALIDATION_TX = [
    (dict(Type="Money Out", Category="Rent", Money=BANK, Amount=100), "⛔ Missing transaction date"),
    (dict(Date="next week", Type="Money Out", Category="Rent", Money=BANK, Amount=100), "⛔ Invalid date"),
    (dict(Date="1990-05-01", Type="Money Out", Category="Rent", Money=BANK, Amount=100), "⛔ Invalid date"),
    (dict(Date="2026-02-01", Category="Rent", Money=BANK, Amount=100), "⛔ Select a type"),
    (dict(Date="2026-02-01", Type="Sale", Category="Rent", Money=BANK, Amount=100), "⛔ Unknown type: Sale"),
    (dict(Date="2026-02-01", Type="Money Out", Money=BANK, Amount=100), "⛔ Select an account (Category)"),
    (dict(Date="2026-02-01", Type="Money Out", Category="Snacks", Money=BANK, Amount=100),
     "⛔ Unknown account: Snacks"),
    (dict(Date="2026-02-01", Type="Money Out", Category="Rent", Amount=100), "⛔ Select Paid From / Deposited To"),
    (dict(Date="2026-02-01", Type="Money Out", Category="Rent", Money="Piggy Bank", Amount=100),
     "⛔ Unknown account: Piggy Bank"),
    (dict(Date="2026-02-01", Type="Money Out", Category="Rent", Money=BANK), "⛔ Enter an amount"),
    (dict(Date="2026-02-01", Type="Money Out", Category="Rent", Money=BANK, Amount=0), "⛔ Enter an amount"),
    (dict(Date="2026-02-01", Type="Money Out", Category="Rent", Money=BANK, Amount="abc"), "⛔ Enter an amount"),
    (dict(Date="2026-02-01", Type="Money In", Category="Product Sales", Money=CASH, Amount=100, Tax=150),
     "⛔ Tax is larger than the amount"),
    (dict(ID="DUP-1", Date="2026-02-02", Type="Money In", Category="Product Sales", Money=CASH, Amount=10),
     "⚠ Duplicate transaction ID"),
    (dict(ID="DUP-1", Date="2026-02-02", Type="Money In", Category="Product Sales", Money=CASH, Amount=20),
     "⚠ Duplicate transaction ID"),
    (dict(Date="2026-02-03", Type="Money In", Category="Product Sales", Money=CASH, Amount=-50),
     "⚠ Negative amount"),
    (dict(Date="2026-02-03", Type="Customer Payment", Ref="INV-1", Money=BANK, Amount=50), "⚠ Add the customer"),
    (dict(Date="2026-02-03", Type="Customer Payment", Customer="Brightpath Academy", Ref="INV-404", Money=BANK,
          Amount=50), "⚠ Invoice # in Reference not found"),
    (dict(Date="2026-02-03", Type="Bill Payment", Ref="B-1", Money=BANK, Amount=50), "⚠ Add the vendor"),
    (dict(Date="2026-02-03", Type="Bill Payment", Vendor="Fresh Dairy Co.", Ref="B-1", Money=BANK, Amount=50),
     "⚠ Bill # + vendor not found"),
    (dict(Date="2026-02-03", Type="Transfer", Category=BANK, Money=CASH, Amount=50, Tax=5), "⚠ Tax ignored"),
    (dict(Date="2026-02-03", Type="Transfer", Category=BANK, Money=BANK, Amount=50), "⚠ Same account on both sides"),
    (dict(Date="2025-12-15", Type="Money Out", Category="Rent", Money=BANK, Amount=50),
     "⚠ Date is before the accounting start date"),
    (dict(Date="2026-02-03", Type="Money Out", Category="Rent", Money="Inventory", Amount=50),
     "⚠ Paid From / Deposited To is not a payment account"),
    (dict(Desc="just a note, nothing else"), ""),
    (dict(Date="2026-02-04", Type="Money Out", Category="Rent", Money=BANK, Amount=1000), "✓ OK"),
]
VALIDATION_INV = [
    (dict(No="INV-1", Date="2026-01-10", Party="Brightpath Academy", Account="Service Income", Amount=1000),
     "✓ OK"),
    (dict(Date="2026-01-10", Party="Brightpath Academy", Account="Service Income", Amount=1000),
     "⛔ Enter an invoice #"),
    (dict(No="INV-3", Party="Brightpath Academy", Account="Service Income", Amount=1000), "⛔ Missing invoice date"),
    (dict(No="INV-4", Date="2026-01-10", Party="Brightpath Academy", Account="Gizmos", Amount=1000),
     "⛔ Unknown account: Gizmos"),
    (dict(No="INV-5", Date="2026-01-10", Party="Brightpath Academy", Account="Service Income", Amount=-5),
     "⛔ Amount must be positive"),
    (dict(No="INV-6", Date="2026-01-10", Party="Brightpath Academy", Account="Service Income", Amount=500),
     "⚠ Duplicate invoice #"),
    (dict(No="INV-6", Date="2026-01-11", Party="Brightpath Academy", Account="Service Income", Amount=500),
     "⚠ Duplicate invoice #"),
    (dict(No="INV-7", Date="2026-01-10", Account="Service Income", Amount=500), "⚠ Select a customer"),
    (dict(No="INV-8", Date="2026-01-10", Party="Brightpath Academy", Account="Rent", Amount=500),
     "⚠ Account is not a revenue account"),
    (dict(No="INV-9", Date="2026-01-10", Party="Northwind Offices", Account="Service Income", Amount=300),
     "⚠ Overpaid by 200.00"),
]
VALIDATION_BILL = [
    (dict(No="B-1", Date="2026-01-10", Party="Golden Grain Flour Mill", Account="Direct Materials", Amount=900),
     "✓ OK"),
    (dict(No="B-2", Date="2026-01-10", Account="Direct Materials", Amount=900), "⛔ Select a vendor"),
    (dict(No="B-3", Date="2026-01-10", Party="Fresh Dairy Co.", Account="Product Sales", Amount=900),
     "⚠ Account is not an expense or asset account"),
]


def validation():
    tx = [r for r, _ in VALIDATION_TX]
    tx.append(dict(Date="2026-02-05", Type="Customer Payment", Customer="Northwind Offices", Ref="INV-9",
                   Money=BANK, Amount=500))
    return {
        "setup": {"StartDate": "2026-01-01", "TaxOn": "Yes"},
        "report": {"year": 2026, "month": "February"},
        "customers": SC.CUSTOMERS, "vendors": SC.VENDORS,
        "transactions": tx,
        "invoices": [r for r, _ in VALIDATION_INV],
        "bills": [r for r, _ in VALIDATION_BILL],
        "extra_accounts": [(6000, "Rent", "Operating Expense", "Debit", "Operating", "No", "duplicate name")],
    }


def _boundary_tx():
    return [
        dict(Date="2025-12-31", Type="Money In", Category="Product Sales", Money=CASH, Amount=700),   # prior FY
        dict(Date="2026-01-01", Type="Money In", Category="Product Sales", Money=CASH, Amount=1000),
        dict(Date="2026-01-31", Type="Money Out", Category="Rent", Money=BANK, Amount=300),
        dict(Date="2026-02-01", Type="Money Out", Category="Rent", Money=BANK, Amount=310),
        dict(Date="2026-06-30", Type="Money In", Category="Service Income", Money=BANK, Amount=2000),
        dict(Date="2026-07-01", Type="Money In", Category="Service Income", Money=BANK, Amount=2100),
        dict(Date="2026-09-30", Type="Money Out", Category="Packaging Design", Money=CASH, Amount=123.45),
        dict(Date="2026-09-30", Type="Money In", Category="Product Sales", Money=BANK, Amount=987654321.99),
        dict(Date="2026-09-30", Type="Money In", Category="Product Sales", Money=BANK, Amount=0.01),
        dict(),  # an empty row in the middle (deleted contents)
        dict(Date="2026-10-01", Type="Money In", Category="Product Sales", Money=BANK, Amount=5555),  # after as-of
        dict(Date="2026-12-31", Type="Money Out", Category="Rent", Money=BANK, Amount=444),           # FY end
        dict(Date="2027-01-01", Type="Money Out", Category="Rent", Money=BANK, Amount=999),            # next FY
    ]


def boundaries():
    return {
        "setup": {"StartDate": "2025-12-01", "OpenCash": 1000, "OpenBank": 2000, "TaxOn": "No"},
        "report": {"year": 2026, "month": "September"},
        "opening": {"Equipment": 5000, "Loans Payable": 3000, "Retained Earnings": 1500, "Rent": 250},
        "extra_accounts": [(6155, "Packaging Design", "Operating Expense", "Debit", "Operating", "No", "added"),
                           (6999, "Old Account", "Operating Expense", "Debit", "Operating", "No", "inactive", "No")],
        "transactions": _boundary_tx(),
    }


def fy_july():
    d = boundaries()
    d["setup"] = dict(d["setup"], FYStartMonthName="July")
    d["report"] = {"year": 2027, "month": "January"}
    return d


def fy_july_first_month():
    d = fy_july()
    d["report"] = {"year": 2026, "month": "July"}
    return d


def tax_off():
    d = scenarios()
    d["setup"] = dict(d["setup"], TaxOn="No")
    return d


def one_transaction():
    return {"setup": {"OpenBank": 0}, "report": {"year": 2026, "month": "May"},
            "transactions": [dict(Date="2026-05-15", Type="Money In", Category="Service Income", Money=BANK,
                                  Amount=2500)]}


def randomized(seed):
    """Random but valid-ish books across two fiscal years, including some invalid rows."""
    rnd = random.Random(seed)
    income = ["Product Sales", "Service Income", "Other Income", "Interest Income"]
    expense = ["Rent", "Utilities", "Salaries & Wages", "Advertising & Marketing", "Direct Materials",
               "Cost of Goods Sold", "Interest Expense", "Bank & Payment Fees", "Refunds & Discounts"]
    money = [CASH, BANK, EW, CC]
    customers = [{"Name": f"Customer {i}", "Terms": rnd.choice([0, 7, 15, 30, None])} for i in range(6)]
    vendors = [{"Name": f"Vendor {i}", "Terms": rnd.choice([15, 30, 45])} for i in range(4)]
    start = dt.date(2025, 4, 1)

    def day():
        return (start + dt.timedelta(days=rnd.randint(0, 600))).isoformat()

    def amt():
        return round(rnd.choice([rnd.uniform(1, 500), rnd.uniform(500, 20000), rnd.uniform(0.01, 5)]), 2)

    tx, inv, bills = [], [], []
    for _ in range(140):
        kind = rnd.random()
        a = amt()
        if kind < 0.3:
            tx.append(dict(Date=day(), Type="Money In", Category=rnd.choice(income), Money=rnd.choice(money),
                           Amount=round(a * 1.12, 2), Tax=round(a * 0.12, 2) if rnd.random() < 0.6 else None))
        elif kind < 0.6:
            tx.append(dict(Date=day(), Type="Money Out", Category=rnd.choice(expense + ["Equipment", "Loans Payable"]),
                           Money=rnd.choice(money), Amount=round(a * 1.12, 2),
                           Tax=round(a * 0.12, 2) if rnd.random() < 0.5 else None))
        elif kind < 0.7:
            src, dst = rnd.sample(money, 2)
            tx.append(dict(Date=day(), Type="Transfer", Category=dst, Money=src, Amount=a))
        elif kind < 0.78:
            tx.append(dict(Date=day(), Type=rnd.choice(["Owner Investment", "Owner Draw"]), Money=BANK, Amount=a))
        elif kind < 0.82:
            tx.append(dict(Date=day(), Type="Adjustment", Category="Depreciation", Money="Accumulated Depreciation",
                           Amount=a))
        elif kind < 0.86:  # deliberately broken rows
            bad = rnd.choice([dict(Type="Money In", Category="Product Sales", Money=CASH, Amount=a),
                              dict(Date=day(), Type="Money In", Category="Nope", Money=CASH, Amount=a),
                              dict(Date=day(), Type="Money Out", Category="Rent", Money=CASH, Amount="x")])
            tx.append(bad)
        else:
            if rnd.random() < 0.5:
                no = f"INV-{len(inv) + 1}"
                c = rnd.choice(customers)["Name"]
                d0 = day()
                inv.append(dict(No=no, Date=d0, Party=c, Account=rnd.choice(income[:2]), Amount=round(a * 1.12, 2),
                                Tax=round(a * 0.12, 2)))
                for _p in range(rnd.choice([0, 1, 1, 2])):
                    pd = (dt.date.fromisoformat(d0) + dt.timedelta(days=rnd.randint(1, 120))).isoformat()
                    tx.append(dict(Date=pd, Type="Customer Payment", Customer=c, Ref=no, Money=BANK,
                                   Amount=round(a * rnd.choice([0.5, 0.56, 1.12]), 2)))
            else:
                no = f"BL-{len(bills) + 1}"
                v = rnd.choice(vendors)["Name"]
                d0 = day()
                bills.append(dict(No=no, Date=d0, Party=v, Account=rnd.choice(expense[:6] + ["Inventory"]),
                                  Amount=round(a * 1.12, 2), Tax=round(a * 0.12, 2) if rnd.random() < 0.7 else None))
                if rnd.random() < 0.7:
                    pd = (dt.date.fromisoformat(d0) + dt.timedelta(days=rnd.randint(1, 90))).isoformat()
                    tx.append(dict(Date=pd, Type="Bill Payment", Vendor=v, Ref=no, Money=rnd.choice([BANK, CC]),
                                   Amount=round(a * rnd.choice([0.5, 1.12]), 2)))
    fy = rnd.choice(["January", "April", "July", "October"])
    month = rnd.choice(["January", "March", "June", "September", "December"])
    return {
        "setup": {"StartDate": start.isoformat(), "FYStartMonthName": fy, "OpenCash": rnd.randint(0, 5000),
                  "OpenBank": rnd.randint(0, 50000), "TaxOn": "Yes"},
        "report": {"year": 2026, "month": month},
        "opening": {"Equipment": 12000, "Loans Payable": 8000, "Inventory": 1500},
        "customers": customers, "vendors": vendors,
        "transactions": tx, "invoices": inv, "bills": bills,
    }


ALL = {
    "blank": blank,
    "sample": sample,
    "scenarios": scenarios,
    "validation": validation,
    "boundaries": boundaries,
    "fy_july": fy_july,
    "fy_july_first_month": fy_july_first_month,
    "tax_off": tax_off,
    "one_transaction": one_transaction,
    "random_1": lambda: randomized(1),
    "random_2": lambda: randomized(2),
    "random_3": lambda: randomized(3),
}
