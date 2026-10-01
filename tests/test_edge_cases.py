"""Boundaries: fiscal years, month/year ends, future-dated and backdated entries, openings."""

import datetime as dt

import pytest

from conftest import as_date, needs_lo
from simpleledger.reports import LEDGER_FIRST

pytestmark = needs_lo


def test_prior_year_entries_become_prior_year_profit(books):
    b = books["boundaries"]
    # 700 sale on 31-Dec-2025 + 250 rent opening balance (a P&L account) -> prior profit 450
    assert b.name("BS_PriorProfit") == pytest.approx(700 - 250)


def test_month_end_and_year_end_boundaries(books):
    b = books["boundaries"]
    ms_rev = b.name("ms_Rev")
    assert ms_rev[0] == pytest.approx(1000)            # Jan 1 counts in January
    assert ms_rev[5] == pytest.approx(2000)            # Jun 30 in June
    assert ms_rev[6] == pytest.approx(2100)            # Jul 1 in July
    assert b.name("ms_Opex")[0] == pytest.approx(300)  # Jan 31 in January
    assert b.name("ms_Opex")[11] == pytest.approx(444)  # Dec 31 inside fiscal year
    assert b.name("PL_Opex_FY") == pytest.approx(300 + 310 + 123.45 + 444)  # 1-Jan-2027 excluded


def test_entries_after_report_date_excluded_from_balances(books):
    b = books["boundaries"]
    assert b.name("PL_Rev_YTD") == pytest.approx(1000 + 2000 + 2100 + 987654321.99 + 0.01)
    assert b.name("PL_Rev_FY") == pytest.approx(b.name("PL_Rev_YTD") + 5555)


def test_large_and_tiny_amounts_are_exact(books):
    calc = books["boundaries"].calc()
    assert calc["Product Sales"]["Cur"] == pytest.approx(-(987654321.99 + 0.01), abs=1e-6)


def test_added_account_reported_and_inactive_hidden_from_dropdown(books):
    b = books["boundaries"]
    assert b.row_by_label("PNL", "Packaging Design")
    accounts = b.name("lstAccounts")
    assert "Packaging Design" in accounts and "Old Account" not in accounts


def test_opening_balance_equity_balances_openings(books):
    calc = books["boundaries"].calc()
    total_open = sum(v["Open"] for v in calc.values())
    assert total_open == pytest.approx(0)
    # debit openings 1000+2000+5000+250 ; credit openings 3000+1500  -> OBE credit 3750
    assert calc["Opening Balance Equity"]["Open"] == pytest.approx(-3750)


def test_fiscal_year_starting_in_july(books):
    b = books["fy_july"]
    assert as_date(b.name("FYStart")) == dt.date(2026, 7, 1)
    assert b.name("RptIdx") == 7
    assert b.name("PL_Rev_Prev") == 0                  # December 2026 had no revenue
    assert b.name("PL_Opex_Prev") == pytest.approx(444)
    assert b.name("PL_Opex_Cur") == pytest.approx(999)  # 1-Jan-2027


def test_first_month_of_fiscal_year_previous_month_comes_from_prior_year(books):
    b = books["fy_july_first_month"]
    assert b.name("RptIdx") == 1
    assert b.name("PL_Rev_Prev") == pytest.approx(2000)   # 30-Jun-2026, previous fiscal year
    assert b.name("CF_Begin_Cur") == pytest.approx(b.name("CF_Begin_YTD"))


def test_single_transaction(books):
    b = books["one_transaction"]
    assert b.name("PL_Rev_Cur") == pytest.approx(2500)
    assert b.name("DB_Cash") == pytest.approx(2500)


def test_ledger_lists_entries_in_date_order(books):
    b = books["boundaries"]
    ws = b.ws("LEDGER")
    dates = [as_date(ws[f"C{r}"].value) for r in range(LEDGER_FIRST, LEDGER_FIRST + 30) if ws[f"C{r}"].value]
    assert dates == sorted(dates) and len(dates) >= 5
