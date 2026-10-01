"""Every report figure must equal the independent reference model, for every data set."""

import pytest

from compare import compare
from conftest import needs_lo
import datasets

pytestmark = needs_lo


@pytest.mark.parametrize("name", list(datasets.ALL))
def test_workbook_matches_reference(books, name):
    _ref, mismatches = compare(books[name])
    assert not mismatches, f"{len(mismatches)} mismatches:\n" + "\n".join(mismatches[:40])


@pytest.mark.parametrize("name", list(datasets.ALL))
def test_no_formula_errors(books, name):
    assert books[name].errors() == []


@pytest.mark.parametrize("name", list(datasets.ALL))
def test_accounting_equations_hold(books, name):
    b = books[name]
    assert b.name("TB_Status") == "✓ BALANCED"
    assert b.name("TB_Journal") == 0
    assert round(b.name("BS_TotalAssets") - b.name("BS_TotalLE"), 2) == 0
    assert b.name("CF_End_Cur") == pytest.approx(b.name("CF_EndBS_Cur"), abs=0.01)


def test_blank_workbook_is_all_zero_and_healthy(books):
    b = books["blank"]
    for n in ("PL_Rev_YTD", "PL_Net_YTD", "BS_TotalAssets", "CF_End_Cur", "DB_AR", "DB_AP", "TB_Debit"):
        assert (b.name(n) or 0) == 0, n
    assert b.name("DB_Errors") == "✓ None"
    assert b.name("DB_Warnings") == "✓ None"


def test_sample_business_is_clean(books):
    b = books["sample"]
    assert b.name("DB_Errors") == "✓ None"
    assert b.name("DB_Warnings") == "✓ None"
    assert b.name("BizName") == "Sunrise Bakery & Café"
    assert b.name("PL_Rev_YTD") > 2_000_000
    assert 0 < b.name("DB_NM") < 0.5
