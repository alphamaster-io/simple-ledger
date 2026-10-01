"""Error prevention: every bad input gets the right message; bad rows never reach reports."""

import pytest

from conftest import needs_lo
from datasets import VALIDATION_BILL, VALIDATION_INV, VALIDATION_TX

pytestmark = needs_lo


def _cases(rows):
    return [(i, msg) for i, (_r, msg) in enumerate(rows)]


@pytest.mark.parametrize("i,msg", _cases(VALIDATION_TX), ids=[m or "blank" for _r, m in VALIDATION_TX])
def test_transaction_messages(books, i, msg):
    got = books["validation"].column("TRANSACTIONS", "Check")[i]
    if msg:
        assert str(got).startswith(msg), got
    else:
        assert got in (None, ""), got


@pytest.mark.parametrize("i,msg", _cases(VALIDATION_INV), ids=[m for _r, m in VALIDATION_INV])
def test_invoice_messages(books, i, msg):
    got = books["validation"].column("AR", "Check")[i]
    assert str(got).startswith(msg), got


@pytest.mark.parametrize("i,msg", _cases(VALIDATION_BILL), ids=[m for _r, m in VALIDATION_BILL])
def test_bill_messages(books, i, msg):
    got = books["validation"].column("AP", "Check")[i]
    assert str(got).startswith(msg), got


def test_error_rows_are_not_posted(books):
    b = books["validation"]
    posted = b.column("TRANSACTIONS", "Posted")
    checks = b.column("TRANSACTIONS", "Check")
    for p, c in zip(posted, checks):
        if str(c).startswith("⛔"):
            assert p == 0, c
        elif str(c).startswith(("⚠", "✓")):
            assert p == 1, c


def test_dashboard_counts_errors_and_flags_setup(books):
    b = books["validation"]
    assert str(b.name("DB_Errors")).startswith("⛔ ")
    ws = b.ws("DASHBOARD")
    assert str(ws["J19"].value).startswith("⛔"), "duplicate account name must be flagged in Setup health"


def test_duplicate_account_name_flagged_and_not_double_counted(books):
    b = books["validation"]
    checks = [c for n, c in zip(b.column("COA", "Account Name"), b.column("COA", "Check")) if n == "Rent"]
    assert checks == ["⛔ Duplicate account name"] * 2
    assert b.name("TB_Status") == "✓ BALANCED"
