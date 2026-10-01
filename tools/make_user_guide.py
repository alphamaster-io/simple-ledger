"""Generate docs/user-guide.md from the HELP sheet content (single source of truth).

    python tools/make_user_guide.py
"""

from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from simpleledger import __version__  # noqa: E402
from simpleledger.ui import HELP, PRODUCT, PROMISE  # noqa: E402
from simpleledger import model as M  # noqa: E402

lines = [
    f"# {PRODUCT} — User Guide (v{__version__})",
    "",
    f"_{PROMISE}_",
    "",
    "> Generated from the workbook's HELP sheet by `tools/make_user_guide.py` — edit `src/simpleledger/ui.py`, "
    "not this file.",
    "",
    "## The sheets at a glance",
    "",
    "| Sheet | What it's for |",
    "|---|---|",
    "| START_HERE | Welcome page, workflow and buttons to every sheet |",
    "| SETUP | Your business details, start date, currency, opening cash and bank, tax settings |",
    "| DASHBOARD | Pick the report month; KPIs, data health, charts |",
    f"| TRANSACTIONS | Every sale, expense, payment and transfer (up to {M.TX_ROWS:,} rows) |",
    "| AR / AP | Customer invoices you'll be paid for later / supplier bills you'll pay later |",
    "| CUSTOMERS / VENDORS | Your customers and suppliers, with what they owe / you owe |",
    "| COA | Chart of accounts: your categories and opening balances |",
    "| PNL, BALANCE_SHEET, CASH_FLOW, MONTHLY_SUMMARY | Reports (automatic, print-ready) |",
    "| TRIAL_BALANCE, LEDGER | For bookkeepers: balances per account; every posting to one account |",
    "| HELP | This guide |",
    "| LISTS, JOURNAL, CALC | Settings and the calculation engine — look, don't touch |",
    "",
    "## Transaction types",
    "",
    "| Type | Use it for |",
    "|---|---|",
]
lines += [f"| {t} | {h} |" for t, _d, _o, _tax, h in M.TRANSACTION_TYPES]
for head, paras in HELP:
    lines += ["", f"## {head}", ""]
    lines += [f"- {p}" for p in paras]
out = ROOT / "docs" / "user-guide.md"
out.write_text("\n".join(lines) + "\n", encoding="utf-8")
print(out)
