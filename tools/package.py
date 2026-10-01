"""Assemble the customer download bundle from dist/standard into dist/package/.

    python -m simpleledger.build && python tools/package.py
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from simpleledger import __version__  # noqa: E402

NAME = "SimpleLedger-Small-Business-Accounting"
README = f"""SimpleLedger — Small Business Accounting Workbook  v{__version__}
=====================================================================

Thank you for your purchase!

START HERE
1. Open  Excel/{NAME}.xlsx  and read the START_HERE tab (2 minutes).
2. Fill in the yellow cells on the SETUP tab.
3. Record your sales and expenses on the TRANSACTIONS tab.
4. Choose the month at the top of the DASHBOARD.

Want to see a finished example first? Open  Examples/SimpleLedger-Sample-Business-Data.xlsx

REQUIREMENTS
Microsoft Excel 2010 or later (Windows or Mac) or Microsoft 365. No macros.

HELP
The HELP tab inside the workbook answers the most common questions.
Guides/ contains the Quick-Start and User Guide.

License terms: License/License-and-Terms.txt
"""


def main():
    src = ROOT / "dist" / "standard"
    blank = src / f"{NAME}.xlsx"
    sample = src / "SimpleLedger-Sample-Business-Data.xlsx"
    for f in (blank, sample):
        if not f.exists():
            sys.exit(f"missing {f} — run: python -m simpleledger.build")
    out = ROOT / "dist" / "package" / NAME
    if out.exists():
        shutil.rmtree(out)
    for sub in ("Excel", "Examples", "Guides", "License"):
        (out / sub).mkdir(parents=True)
    shutil.copy2(blank, out / "Excel")
    shutil.copy2(sample, out / "Examples")
    shutil.copy2(ROOT / "docs" / "user-guide.md", out / "Guides" / "User-Guide.md")
    shutil.copy2(ROOT / "LICENSE.txt", out / "License" / "License-and-Terms.txt")
    (out / "Read-Me.txt").write_text(README, encoding="utf-8")
    archive = shutil.make_archive(str(out.parent / f"{NAME}-v{__version__}"), "zip", out.parent, NAME)
    print(archive)


if __name__ == "__main__":
    main()
