"""Each business scenario produces exactly the hand-written double entry."""

import pytest

from conftest import needs_lo
from scenarios import SCENARIOS

pytestmark = needs_lo

FIELDS = ("DrAcct", "DrAmt", "CrAcct", "CrAmt", "TaxDr", "TaxCr")
OFFSET = {"tx": 0, "ar": 2000, "ap": 2500}


def _journal(book):
    cols = {f: book.name(f"j_{f}") for f in FIELDS}
    return lambda i: tuple(cols[f][i] for f in FIELDS)


@pytest.mark.parametrize("sc", SCENARIOS, ids=[f"{s[0]:02d}-{s[1]}" for s in SCENARIOS])
def test_scenario_posts_expected_entry(books, sc):
    sid, title, src, _rec, expected, check = sc
    book = books["scenarios"]
    index = sum(1 for s in SCENARIOS if s[2] == src and s[0] < sid)
    row = _journal(book)(OFFSET[src] + index)
    dr, dra, cr, cra, tdr, tcr = row
    assert (dr, cr) == (expected[0], expected[2]), title
    assert [dra, cra, tdr or 0, tcr or 0] == pytest.approx(list(expected[1:2] + expected[3:]), abs=0.001), title
    sheet = {"tx": "TRANSACTIONS", "ar": "AR", "ap": "AP"}[src]
    assert str(book.column(sheet, "Check")[index]).startswith(check), title


def test_every_entry_balances(books):
    b = books["scenarios"]
    dr = [a + t for a, t in zip(b.name("j_DrAmt"), b.name("j_TaxDr"))]
    cr = [a + t for a, t in zip(b.name("j_CrAmt"), b.name("j_TaxCr"))]
    assert all(abs(x - y) < 1e-9 for x, y in zip(dr, cr))


def test_invoice_and_bill_status_after_payments(books):
    b = books["scenarios"]
    assert b.column("AR", "Status")[:2] == ["Paid", "Overdue"]   # INV-2: 15-day terms, partly paid, Mar 31
    assert b.column("AR", "Balance")[:2] == [0, 1200]
    assert b.column("AP", "Status")[:2] == ["Paid", "Overdue"]
    assert b.column("AP", "Balance")[:2] == [0, 2500]


def test_vat_payable_nets_output_and_input_tax(books):
    calc = books["scenarios"].calc()
    # output VAT: 120 + 360 ; input VAT: 240 + 1200 + 6000 + 60 ; remitted 1000
    expected_credit_balance = (120 + 360) - (240 + 1200 + 6000 + 60) - 1000
    assert -calc["Tax Payable"]["Bal"] == pytest.approx(expected_credit_balance)


def test_tax_off_ignores_tax_columns(books):
    b = books["tax_off"]
    assert sum(b.name("j_TaxDr")) == 0 and sum(b.name("j_TaxCr")) == 0
    assert b.calc()["Tax Payable"]["Bal"] == 1000   # only the explicit VAT remittance (scenario 31)
    checks = b.column("TRANSACTIONS", "Check")
    assert str(checks[1]).startswith("⚠ Tax ignored")   # scenario 2 typed VAT while tracking is off
