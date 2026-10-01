"""Independent reference implementation of the accounting rules (the test oracle).

This re-implements docs/accounting-model.md in plain Python, without looking at
the workbook formulas. Tests build a workbook from a data set, let LibreOffice
calculate it, and compare every report figure with this model.
"""

from __future__ import annotations

import calendar
import datetime as dt
from dataclasses import dataclass, field

from simpleledger import model as M

TYPES = {  # type: (direction, other side)
    "Money In": ("In", "Category"), "Money Out": ("Out", "Category"),
    "Customer Payment": ("In", "AR"), "Bill Payment": ("Out", "AP"), "Transfer": ("Out", "Category"),
    "Owner Investment": ("In", "Capital"), "Owner Draw": ("Out", "Drawings"), "Adjustment": ("Out", "Category"),
}
TAX_TYPES = {"Money In", "Money Out"}
SYSTEM = {"AR": "Accounts Receivable", "AP": "Accounts Payable", "Tax": "Tax Payable",
          "Capital": "Owner Capital", "Drawings": "Owner Drawings", "OBE": "Opening Balance Equity"}
TYPE_INFO = {t: (stmt, sign) for t, stmt, sign, _n, _s in M.ACCOUNT_TYPES}
MIN_DATE, MAX_DATE = dt.date(2000, 1, 1), dt.date(2099, 12, 31)


def _date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str) and v:
        try:
            return dt.date.fromisoformat(v)
        except ValueError:
            return "bad"
    return None


def _num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _blank(v):
    return v is None or v == ""


def add_months(d: dt.date, n: int) -> dt.date:
    y, m = divmod(d.month - 1 + n, 12)
    y += d.year
    m += 1
    return dt.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def month_end(d: dt.date) -> dt.date:
    return dt.date(d.year, d.month, calendar.monthrange(d.year, d.month)[1])


@dataclass
class Posting:
    date: dt.date
    source: str
    sid: str
    dr: str
    dr_amt: float
    cr: str
    cr_amt: float
    tax_dr: float = 0.0
    tax_cr: float = 0.0


@dataclass
class Doc:
    no: object
    date: dt.date
    party: str
    amount: float
    eff_due: dt.date
    paid: float = 0.0
    balance: float = 0.0
    status: str = ""
    bucket: str = ""
    days_over: int | None = None


@dataclass
class Ledger:
    data: dict
    accounts: dict = field(default_factory=dict)   # name -> dict(type, normal, cf, open)
    postings: list = field(default_factory=list)
    errors: dict = field(default_factory=dict)      # (source, index) -> True
    tx_posted: list = field(default_factory=list)   # per tx row: Posting or None

    def __post_init__(self):
        d = self.data
        self.setup = d.get("setup", {})
        self.tax_on = self.setup.get("TaxOn", "No") == "Yes"
        self.terms_default = self.setup.get("DefaultTerms", 30)
        self.start = _date(self.setup.get("StartDate"))
        rep = d.get("report", {})
        y, mname = rep["year"], rep["month"]
        mm = M.MONTHS.index(mname) + 1
        fysm = M.MONTHS.index(self.setup.get("FYStartMonthName", "January")) + 1
        self.rpt_start = dt.date(y, mm, 1)
        self.asof = month_end(self.rpt_start)
        self.fy_start = dt.date(y - (1 if mm < fysm else 0), fysm, 1)
        self.fy_end = add_months(self.fy_start, 12) - dt.timedelta(days=1)
        self.rpt_idx = (y * 12 + mm) - (self.fy_start.year * 12 + self.fy_start.month) + 1
        self.prev_start = add_months(self.rpt_start, -1)
        self.prev_end = self.rpt_start - dt.timedelta(days=1)
        self._accounts()
        self._post()

    # --------------------------------------------------------------------------------- accounts
    def _accounts(self):
        seen = set()
        opening = self.data.get("opening", {})
        for code, name, typ, normal, cf, pay, _desc, *_active in list(M.DEFAULT_ACCOUNTS) + [
                tuple(a) for a in self.data.get("extra_accounts", [])]:
            if name in seen:
                continue  # duplicate names: only the first counts
            seen.add(name)
            if name == "Cash on Hand":
                ob = self.setup.get("OpenCash", 0) or 0
            elif name == "Bank Account":
                ob = self.setup.get("OpenBank", 0) or 0
            elif name == SYSTEM["OBE"]:
                ob = 0
            else:
                ob = opening.get(name, 0) or 0
            signed = ob * (-1 if normal == "Credit" else 1)
            self.accounts[name] = dict(code=code, type=typ, normal=normal, cf=cf or "Operating", pay=pay,
                                       open=signed)
        total = sum(a["open"] for a in self.accounts.values())
        self.accounts[SYSTEM["OBE"]]["open"] -= total

    def exists(self, name):
        return isinstance(name, str) and name in self.accounts

    # --------------------------------------------------------------------------------- posting
    def _valid_date(self, v):
        d = _date(v)
        return d if isinstance(d, dt.date) and MIN_DATE <= d <= MAX_DATE else None

    def tx_error(self, r):
        if all(_blank(r.get(k)) for k in ("Date", "Type", "Category", "Amount", "Money")):
            return "blank"
        if _blank(r.get("Date")):
            return "Missing transaction date"
        if self._valid_date(r["Date"]) is None:
            return "Invalid date"
        t = r.get("Type")
        if _blank(t):
            return "Select a type"
        if t not in TYPES:
            return f"Unknown type: {t}"
        other = self.other_side(r)
        if _blank(other):
            return "Select an account (Category)"
        if not self.exists(other):
            return f"Unknown account: {other}"
        money = r.get("Money")
        if _blank(money):
            return "Select Paid From / Deposited To"
        if not self.exists(money):
            return f"Unknown account: {money}"
        a = r.get("Amount")
        if not _num(a) or a == 0:
            return "Enter an amount"
        if abs(self.tax_eff(r)) > abs(a):
            return "Tax is larger than the amount"
        return None

    def other_side(self, r):
        _d, side = TYPES[r["Type"]]
        return r.get("Category") or "" if side == "Category" else SYSTEM[side]

    def tax_eff(self, r):
        t = r.get("Tax")
        if not self.tax_on or not _num(t) or r.get("Type") not in TAX_TYPES:
            return 0.0
        return t

    def _post(self):
        for i, r in enumerate(self.data.get("transactions", [])):
            err = self.tx_error(r)
            if err:
                self.tx_posted.append(None)
                if err != "blank":
                    self.errors[("tx", i)] = err
                continue
            direction, _ = TYPES[r["Type"]]
            other, money, a, t = self.other_side(r), r["Money"], r["Amount"], self.tax_eff(r)
            d = self._valid_date(r["Date"])
            sid = r.get("ID") or f"TX-{i + 1:04d}"
            if direction == "In":
                p = Posting(d, "Transaction", sid, money, a, other, a - t, 0, t)
            else:
                p = Posting(d, "Transaction", sid, other, a - t, money, a, t, 0)
            p.row = r
            self.postings.append(p)
            self.tx_posted.append(p)
        self.invoices = self._docs("ar")
        self.bills = self._docs("ap")

    def doc_error(self, r, kind):
        if all(_blank(r.get(k)) for k in ("No", "Date", "Party", "Amount", "Account")):
            return "blank"
        if _blank(r.get("No")):
            return "missing no"
        if kind == "ap" and _blank(r.get("Party")):
            return "Select a vendor"
        if _blank(r.get("Date")) or self._valid_date(r["Date"]) is None:
            return "date"
        if _blank(r.get("Account")) or not self.exists(r["Account"]):
            return "account"
        if not _num(r.get("Amount")) or r["Amount"] <= 0:
            return "amount"
        if self.doc_tax(r) > r["Amount"]:
            return "tax"
        return None

    def doc_tax(self, r):
        t = r.get("Tax")
        return t if self.tax_on and _num(t) else 0.0

    def _docs(self, kind):
        rows = self.data.get("invoices" if kind == "ar" else "bills", [])
        parties = {p["Name"]: p for p in self.data.get("customers" if kind == "ar" else "vendors", [])}
        out = []
        for i, r in enumerate(rows):
            err = self.doc_error(r, kind)
            if err:
                out.append(None)
                if err != "blank":
                    self.errors[(kind, i)] = err
                continue
            d = self._valid_date(r["Date"])
            a, t = r["Amount"], self.doc_tax(r)
            if kind == "ar":
                p = Posting(d, "Invoice", str(r["No"]), SYSTEM["AR"], a, r["Account"], a - t, 0, t)
            else:
                p = Posting(d, "Bill", str(r["No"]), r["Account"], a - t, SYSTEM["AP"], a, t, 0)
            self.postings.append(p)
            due = _date(r.get("Due"))
            if not isinstance(due, dt.date):
                terms = parties.get(r.get("Party"), {}).get("Terms")
                due = d + dt.timedelta(days=terms if _num(terms) else self.terms_default)
            doc = Doc(r["No"], d, r.get("Party") or "", a, due)
            # payments up to the as-of date (Excel SUMIFS text matching is case-insensitive)
            ptype = "Customer Payment" if kind == "ar" else "Bill Payment"
            for tp in self.tx_posted:
                if tp is None or tp.row["Type"] != ptype or tp.date > self.asof:
                    continue
                if str(tp.row.get("Ref", "")).lower() != str(r["No"]).lower():
                    continue
                if kind == "ap" and str(tp.row.get("Vendor", "")).lower() != str(r.get("Party", "")).lower():
                    continue
                doc.paid += tp.dr_amt if kind == "ar" else tp.cr_amt
            doc.balance = 0.0 if d > self.asof else round(a - doc.paid, 2)
            if d > self.asof:
                doc.status = "Future"
            elif doc.balance < 0:
                doc.status = "Overpaid"
            elif doc.balance == 0:
                doc.status = "Paid"
            elif self.asof > due:
                doc.status = "Overdue"
            else:
                doc.status = "Partial" if doc.paid > 0 else "Unpaid"
            if doc.balance > 0:
                doc.days_over = max(0, (self.asof - due).days)
                o = doc.days_over
                doc.bucket = ("Current" if o == 0 else "1-30" if o <= 30 else "31-60" if o <= 60
                              else "61-90" if o <= 90 else "90+")
            out.append(doc)
        return out

    # --------------------------------------------------------------------------------- queries
    def net(self, acct, lo=None, hi=None):
        """Debit-minus-credit activity of an account for postings with lo <= date <= hi."""
        s = 0.0
        for p in self.postings:
            if (lo and p.date < lo) or (hi and p.date > hi):
                continue
            if p.dr == acct:
                s += p.dr_amt
            if p.cr == acct:
                s -= p.cr_amt
            if acct == SYSTEM["Tax"]:
                s += p.tax_dr - p.tax_cr
        return s

    def balance(self, acct, asof=None):
        return self.accounts[acct]["open"] + self.net(acct, None, asof or self.asof)

    def month_range(self, k):
        start = add_months(self.fy_start, k - 1)
        return start, month_end(start)

    def periods(self):
        return {
            "Cur": (self.rpt_start, self.asof), "Prev": (self.prev_start, self.prev_end),
            "YTD": (self.fy_start, self.asof), "FY": (self.fy_start, self.fy_end),
            **{f"M{k}": self.month_range(k) for k in range(1, 13)},
        }

    def pl_line(self, acct, period):
        lo, hi = self.periods()[period]
        return TYPE_INFO[self.accounts[acct]["type"]][1] * self.net(acct, lo, hi)

    def pl_total(self, typ, period):
        return sum(self.pl_line(a, period) for a, v in self.accounts.items() if v["type"] == typ)

    def net_profit(self, period):
        lo, hi = self.periods()[period]
        return -sum(self.net(a, lo, hi) for a, v in self.accounts.items() if TYPE_INFO[v["type"]][0] == "PL")

    def bs_total(self, typ, asof=None):
        sign = TYPE_INFO[typ][1]
        return sum(sign * self.balance(a, asof) for a, v in self.accounts.items() if v["type"] == typ)

    def is_cash(self, acct):
        return self.accounts.get(acct, {}).get("cf") == "Cash"

    def cash_flows(self, lo, hi):
        """{(group, 'in'|'out'): amount} for postings in [lo, hi] (direct method)."""
        out = {}
        for p in self.postings:
            if p.date < lo or p.date > hi:
                continue
            sign = int(self.is_cash(p.dr)) - int(self.is_cash(p.cr))
            if sign == 0:
                continue
            other = p.cr if sign == 1 else p.dr
            group = self.accounts[other]["cf"]
            main = p.cr_amt if sign == 1 else -p.dr_amt
            tax = p.tax_cr if sign == 1 else -p.tax_dr
            for g, v in ((group, main), ("Operating", tax)):
                if v:
                    key = (g, "in" if v > 0 else "out")
                    out[key] = out.get(key, 0) + v
        return out

    def cash_balance(self, asof=None):
        return sum(self.balance(a, asof) for a in self.accounts if self.is_cash(a))
