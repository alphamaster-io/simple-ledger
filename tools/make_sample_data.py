"""Generate src/sample-data/sunrise_bakery.json — the demo business shipped with the product.

Deterministic (fixed seed) so tests and screenshots are reproducible.
Fictional business; all names are invented.

    python tools/make_sample_data.py
"""

from __future__ import annotations

import datetime as dt
import json
import random
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / "src" / "sample-data" / "sunrise_bakery.json"
VAT = 0.12
END = dt.date(2026, 9, 30)


def vat_of(gross: float) -> float:
    return round(gross * VAT / (1 + VAT), 2)


def d(y, m, day):
    return dt.date(y, m, day).isoformat()


def main():
    rnd = random.Random(20260101)
    tx, inv, bills = [], [], []

    def add(date, typ, amount, money, category="", desc="", tax=None, **kw):
        rec = {"Date": date, "Type": typ, "Category": category, "Desc": desc, "Money": money,
               "Amount": round(amount, 2)}
        if tax:
            rec["Tax"] = tax
        rec.update(kw)
        tx.append(rec)

    customers = [
        {"Name": "Brightpath Academy", "Contact": "Ana Reyes", "Email": "admin@brightpath.example",
         "Phone": "0917 555 0101", "Terms": 30},
        {"Name": "Northwind Offices", "Contact": "Carlo Dizon", "Email": "events@northwind.example",
         "Phone": "0917 555 0102", "Terms": 15},
        {"Name": "Casa Verde Events", "Contact": "Bea Lim", "Email": "bea@casaverde.example",
         "Phone": "0917 555 0103", "Terms": 30},
        {"Name": "Luntian Coworking", "Contact": "Paolo Cruz", "Email": "ops@luntian.example",
         "Phone": "0917 555 0104", "Terms": 15},
        {"Name": "St. Jude Clinic", "Contact": "Dr. Rosa Tan", "Email": "office@stjude.example",
         "Phone": "0917 555 0105", "Terms": 30},
    ]
    vendors = [
        {"Name": "Golden Grain Flour Mill", "Contact": "Mr. Uy", "Phone": "02 8555 0201", "Terms": 30},
        {"Name": "Fresh Dairy Co.", "Contact": "Liza", "Phone": "02 8555 0202", "Terms": 15},
        {"Name": "PackRight Supplies", "Contact": "Jun", "Phone": "02 8555 0203", "Terms": 30},
    ]

    inv_no = 1001
    ads_prev = 0.0
    for m in range(1, 10):
        # Weekly counter sales (cash) and e-wallet sales, deposits of cash to bank.
        for week, day in enumerate((3, 10, 17, 24)):
            cash = rnd.uniform(42000, 62000) * (1 + 0.04 * m / 9)
            add(d(2026, m, day + 1), "Money In", cash, "Cash on Hand", "Product Sales",
                f"Counter sales week {week + 1}", vat_of(cash), Method="Cash")
            ew = rnd.uniform(14000, 24000)
            add(d(2026, m, day + 2), "Money In", ew, "E-Wallet / Online Payments", "Product Sales",
                f"GCash / online orders week {week + 1}", vat_of(ew), Method="E-Wallet")
            add(d(2026, m, day + 3), "Transfer", round(cash * 0.85, -2), "Cash on Hand", "Bank Account",
                "Cash deposit to bank", Method="Cash")
        add(d(2026, m, 28), "Transfer", round(rnd.uniform(40000, 60000), -2), "E-Wallet / Online Payments",
            "Bank Account", "E-wallet cash-out to bank", Method="Online Payment")
        # Fixed monthly costs
        add(d(2026, m, 1), "Money Out", 28000, "Bank Account", "Rent", "Shop rent", vat_of(28000),
            Method="Bank Transfer", Ref=f"RENT-{m:02d}")
        util = rnd.uniform(9000, 13500)
        add(d(2026, m, 12), "Money Out", util, "Bank Account", "Utilities", "Electricity & water", vat_of(util),
            Method="Bank Transfer")
        add(d(2026, m, 15), "Money Out", 2499, "Bank Account", "Internet & Phone", "Fiber internet",
            vat_of(2499), Method="Bank Transfer")
        add(d(2026, m, 15), "Money Out", 38000, "Bank Account", "Direct Labor", "Bakers' salaries (1st-15th)",
            Method="Bank Transfer")
        add(d(2026, m, 15), "Money Out", 9000, "Bank Account", "Salaries & Wages", "Cashier salary (1st-15th)",
            Method="Bank Transfer")
        add(d(2026, m, 28), "Money Out", 38000, "Bank Account", "Direct Labor", "Bakers' salaries (16th-end)",
            Method="Bank Transfer")
        add(d(2026, m, 28), "Money Out", 9000, "Bank Account", "Salaries & Wages", "Cashier salary (16th-end)",
            Method="Bank Transfer")
        ads = rnd.uniform(3000, 6000)
        add(d(2026, m, 8), "Money Out", ads, "Credit Card", "Advertising & Marketing", "Facebook / Instagram ads",
            vat_of(ads), Method="Credit Card")
        add(d(2026, m, 9), "Money Out", 1450, "Credit Card", "Software & Subscriptions", "POS app subscription",
            vat_of(1450), Method="Credit Card")
        pack = rnd.uniform(2500, 4500)
        add(d(2026, m, 6), "Money Out", pack, "Cash on Hand", "Other Direct Costs", "Boxes & paper bags (market)",
            vat_of(pack), Method="Cash")
        fuel = rnd.uniform(3500, 5500)
        add(d(2026, m, 20), "Money Out", fuel, "Cash on Hand", "Transportation & Delivery", "Delivery fuel",
            vat_of(fuel), Method="Cash")
        add(d(2026, m, 27), "Money Out", 150, "Bank Account", "Bank & Payment Fees", "Monthly bank charge",
            Method="Bank Transfer")
        if m > 1:  # pay last month's credit card statement
            add(d(2026, m, 5), "Transfer", round(ads_prev + 1450, 2), "Bank Account", "Credit Card",
                "Credit card payment", Method="Bank Transfer")
        ads_prev = ads
        add(d(2026, m, 25), "Owner Draw", 20000, "Bank Account", desc="Owner salary / personal draw",
            Method="Bank Transfer")
        add(d(2026, m, 10), "Money Out", 8000, "Bank Account", "Loans Payable", "Loan principal",
            Method="Bank Transfer", Ref="LOAN")
        add(d(2026, m, 10), "Money Out", round(1500 - 60 * (m - 1), 2), "Bank Account", "Interest Expense",
            "Loan interest", Method="Bank Transfer", Ref="LOAN")
        add(d(2026, m, 30 if m != 2 else 28), "Adjustment", 3000, "Accumulated Depreciation", "Depreciation",
            "Monthly depreciation of equipment")
        if m % 3 == 0:
            add(d(2026, m, 30), "Money In", round(rnd.uniform(70, 110), 2), "Bank Account", "Interest Income",
                "Bank interest", Method="Bank Transfer")

        # Catering invoices (AR) and their payments
        for _ in range(rnd.choice((2, 3))):
            cust = rnd.choice(customers)
            date = dt.date(2026, m, rnd.randint(2, 26))
            amount = round(rnd.uniform(12000, 45000), 2)
            no = f"INV-{inv_no}"
            inv_no += 1
            inv.append({"No": no, "Date": date.isoformat(), "Party": cust["Name"], "Desc": "Catering order",
                        "Account": "Service Income", "Amount": amount, "Tax": vat_of(amount)})
            pay_days = rnd.choice((10, 14, 20, 28, 35, 45, 70))
            pdate = date + dt.timedelta(days=pay_days)
            if pdate <= END:
                if rnd.random() < 0.2:  # partial payment first
                    part = round(amount * 0.5, 2)
                    add(pdate.isoformat(), "Customer Payment", part, "Bank Account", desc="Partial payment",
                        Customer=cust["Name"], Ref=no, Method="Bank Transfer")
                    later = pdate + dt.timedelta(days=20)
                    if later <= END:
                        add(later.isoformat(), "Customer Payment", amount - part, "Bank Account",
                            desc="Balance payment", Customer=cust["Name"], Ref=no, Method="Bank Transfer")
                else:
                    add(pdate.isoformat(), "Customer Payment", amount, "Bank Account", desc="Invoice payment",
                        Customer=cust["Name"], Ref=no, Method="Bank Transfer")

        # Supplier bills (AP) and payments
        for vend, lo, hi, n in ((vendors[0], 30000, 52000, 2), (vendors[1], 12000, 20000, 2),
                                (vendors[2], 6000, 11000, 1)):
            for k in range(n):
                date = dt.date(2026, m, 3 + 13 * k + rnd.randint(0, 3))
                amount = round(rnd.uniform(lo, hi), 2)
                no = f"{vend['Name'][:2].upper()}-{m:02d}{k + 1}"
                account = "Other Direct Costs" if vend is vendors[2] else "Direct Materials"
                bills.append({"No": no, "Date": date.isoformat(), "Party": vend["Name"],
                              "Desc": "Packaging" if vend is vendors[2] else "Ingredients delivery",
                              "Account": account, "Amount": amount, "Tax": vat_of(amount)})
                pdate = date + dt.timedelta(days=vend["Terms"] + rnd.choice((-5, 0, 3, 12)))
                if pdate <= END:
                    add(pdate.isoformat(), "Bill Payment", amount, "Bank Account", desc="Supplier payment",
                        Vendor=vend["Name"], Ref=no, Method="Bank Transfer")

    # One-off events
    add(d(2026, 4, 2), "Owner Investment", 50000, "Bank Account", desc="Owner capital for new oven",
        Method="Bank Transfer")
    add(d(2026, 4, 6), "Money Out", 112000, "Bank Account", "Equipment", "Deck oven (3-tray)", vat_of(112000),
        Method="Bank Transfer", Vendor="PackRight Supplies", Ref="OR-55821")
    add(d(2026, 6, 18), "Money Out", 1680, "Cash on Hand", "Refunds & Discounts", "Refund — damaged cake order",
        vat_of(1680), Method="Cash")
    add(d(2026, 7, 11), "Money Out", 6500, "Cash on Hand", "Repairs & Maintenance", "Mixer motor repair",
        vat_of(6500), Method="Cash")
    add(d(2026, 2, 14), "Money Out", 7800, "Bank Account", "Taxes & Licenses", "Business permit renewal",
        Method="Bank Transfer")
    add(d(2026, 3, 20), "Money Out", 12000, "Bank Account", "Professional Fees", "Bookkeeper — Q1 VAT filing",
        Method="Bank Transfer")
    add(d(2026, 5, 15), "Money Out", 9800, "Bank Account", "Insurance", "Shop insurance (annual)",
        Method="Bank Transfer")

    # Quarterly VAT remittance: output VAT collected minus input VAT paid in the quarter.
    def net_vat(q):
        months = range(3 * q - 2, 3 * q + 1)
        def in_q(date):
            return dt.date.fromisoformat(date).month in months
        out_vat = sum(r.get("Tax", 0) for r in tx if r["Type"] == "Money In" and in_q(r["Date"]))
        out_vat += sum(r["Tax"] for r in inv if in_q(r["Date"]))
        in_vat = sum(r.get("Tax", 0) for r in tx if r["Type"] == "Money Out" and in_q(r["Date"]))
        in_vat += sum(r["Tax"] for r in bills if in_q(r["Date"]))
        return round(out_vat - in_vat, 2)
    for q, date in ((1, d(2026, 4, 24)), (2, d(2026, 7, 24))):
        add(date, "Money Out", net_vat(q), "Bank Account", "Tax Payable", f"VAT return Q{q} 2026",
            Method="Bank Transfer", Ref=f"VAT-Q{q}")

    tx.sort(key=lambda r: r["Date"])
    data = {
        "setup": {
            "BizName": "Sunrise Bakery & Café", "BizType": "Food & Beverage", "OwnerName": "Maria Santos",
            "BizAddress": "123 Mabini St., Quezon City", "BizEmail": "hello@sunrisebakery.example",
            "BizPhone": "0917 555 0100", "StartDate": "2026-01-01", "FYStartMonthName": "January",
            "Currency": "PHP", "OpenCash": 15000, "OpenBank": 285000, "DefaultTerms": 30, "TaxOn": "Yes",
            "TaxLabel": "VAT",
        },
        "report": {"year": 2026, "month": "September"},
        "opening": {"Inventory": 20000, "Equipment": 180000, "Loans Payable": 150000},
        "customers": customers,
        "vendors": vendors,
        "transactions": tx,
        "invoices": inv,
        "bills": bills,
    }
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")
    print(f"{OUT}: {len(tx)} transactions, {len(inv)} invoices, {len(bills)} bills")


if __name__ == "__main__":
    main()
