# RUNBOOK — SimpleLedger (Small Business Accounting Excel Template)

**Controlling execution document.** Read `CLAUDE.md` first. Update this file at the end of every phase.

- **Current version:** 1.0.0 release candidate (not published)
- **Current phase:** **Phase 10: Compatibility** (needs a person with Excel), in parallel with Phase 12 UAT
- **Last updated:** 2026-10-01

## Phase status

| Phase | Deliverable | Status | Evidence / notes |
|---|---|---|---|
| 0 Repository | Structure, CLAUDE/RUNBOOK/README/CHANGELOG, .gitignore, versioning | ✅ Done | This repo; semver in CHANGELOG |
| 1 Market research | `docs/market-research.md`, `docs/product-positioning.md` (customer profile merged in) | 🟡 First pass | Desk research with sources. **Open:** manual marketplace review-level research (list in market-research §7) |
| 2 Product requirements | `docs/product-spec.md` | ✅ Done | Personas, stories, sheets, dashboard, validation, UX, compatibility, scope |
| 3 Accounting model | `docs/accounting-model.md` | ✅ Done | Executable twin: `tests/reference.py` + 35 scenarios |
| 4 Wireframe | Sheet structure, navigation, input zones | ✅ Done | Folded into the generator; START_HERE buttons, colour legend |
| 5 Master data | COA (44 accounts), customers, vendors, lists, transaction-type rules | ✅ Done | Excel Tables `tblAccounts`, `tblCustomers`, `tblVendors` |
| 6 Transaction engine | `tblTransactions`, AR, AP, validation, JOURNAL, CALC | ✅ Done | test_scenarios, test_validation |
| 7 Reports | Ledger, TB, P&L, BS, CF, Monthly, AR, AP | ✅ Done | test_reports (12 data sets vs reference) |
| 8 Dashboard | KPIs, data health, 4 charts, controls | ✅ Done | Rendered and reviewed via LibreOffice PDF export |
| 9 UX polish | Fonts, colours, freeze panes, print areas, protection | ✅ Done (v1) | Re-review after UAT feedback |
| 10 Compatibility | Excel 365 Win/Mac, Excel Online, Google Sheets | ⏳ **Open: blocker for release** | LibreOffice 24.2 verified automatically. Checklist: `docs/qa-plan.md §4` |
| 11 QA | Automated accounting/calculation/edge-case suite | ✅ Done | `python -m pytest`: 142 tests, all passing (see below) |
| 12 UAT | 3 non-developer testers | ⏳ Open | Script: `docs/qa-plan.md §5` |
| 13 Tiers | Free / Standard / Premium definition | 🟡 Defined | `docs/launch-plan.md`. Free tracker not built yet |
| 14 Branding | Name, logo, cover, screenshots | ⏳ Open | Working name "SimpleLedger"; trademark/domain check needed |
| 15 Packaging | Download bundle | 🟡 Script ready | `tools/package.py`; PDF guides pending |
| 16 Listing | Etsy copy | 🟡 Draft | `marketing/listing-copy/etsy.md` |
| 17 Pricing | Hypothesis | 🟡 Draft | Standard $24 (list $39, launch $19) |
| 18–23 Channels, lead magnet, content, social, video, email | Plans & drafts | 🟡 Drafts | `docs/launch-plan.md`, `marketing/` |
| 24–30 Launch → automation | — | ⏳ Not started | Gated on 10, 12, 14, 15 |

## Validation performed (2026-10-01)

- `python -m pytest`: **142 passed** in about 4 min (LibreOffice Calc 24.2.7, 12 data sets, 2 workers).
- Zero formula errors in every data set; TB, BS and CF reconcile in every data set, including 3 randomized books with deliberately broken rows.
- Visual review of START_HERE, DASHBOARD, PNL, BALANCE_SHEET, CASH_FLOW rendered via LibreOffice PDF export.
- Lint: `pyflakes` clean.

## Decisions log

| Decision | Why |
|---|---|
| Generate the workbook from Python instead of hand-building it | Reproducible, diffable, testable; every release is rebuilt and retested |
| Excel-2010 function set, no XLOOKUP/FILTER/LET | Compatibility and reach. The plan suggested XLOOKUP/FILTER "where supported"; the same results are achieved with INDEX/MATCH/SMALL |
| Defined names over A1 ranges instead of structured refs inside formulas | Portable across Excel, LibreOffice and, likely, Google Sheets. Tables are still used for entry UX (filters, banding) |
| Pre-sized tables (2,000 / 500 / 500) + stacked JOURNAL | Lets reports aggregate one journal; capacity is ample for one fiscal year per workbook; Dashboard warns at 90% |
| "Category" = the chart-of-accounts account; the plan's separate "Account" field is shown as the automatic "Account Type" | One dropdown instead of two; fewer errors |
| No "Status" on transactions | Unpaid items live in AR/AP; a transaction is by definition money that moved |
| Data-entry sheets unprotected; reports, engine and settings protected without a password | Excel can't sort, filter or extend tables on protected sheets; the plan says no password without a clear reason |
| Report period chosen on the Dashboard (SETUP shows it) | One control, where the user looks |
| Direct-method cash flow | Easier for non-accountants than the indirect method; reconciles exactly to the balance sheet |
| VAT handled as recoverable input tax | Fits PH/UK/AU/EU; US users leave Tax blank on purchases (documented) |
| JOURNAL and CALC visible (protected, grey tabs), helper columns grouped | Transparency for bookkeepers; the plan forbids hidden dependencies |

## Known issues / limitations

1. **Not yet opened in real Excel.** Possible risks to check first: table XML repair prompts, chart rendering, data-validation lists that reference other sheets via names (Excel 2010+ supports this), `fullCalcOnLoad` speed.
2. Files have **no cached values** (openpyxl). Excel computes on open, but file previewers (Quick Look, email preview, Etsy thumbnails) show blanks. For marketplace screenshots, use the LibreOffice-recalculated copy (`tools/recalc.py`).
3. Rows added **beyond** table capacity are ignored by the JOURNAL. The Dashboard warns at 90% (transactions only; AR/AP capacity is not warned yet).
4. P&L/BS detail slots are fixed (e.g. 22 operating-expense lines). Extra accounts are still in totals, and a ⚠ line says so.
5. Renaming an account orphans existing rows (they show "Unknown account"). Documented in HELP.
6. COUNTIF/SUMIFS matching treats `*` and `?` in invoice/bill numbers as wildcards. Avoid them in document numbers.
7. Ledger shows the first 400 entries of a date range (a message says so).
8. Google Sheets: untested; do not claim support.

## Next steps (in order)

1. **Phase 10:** run `docs/qa-plan.md §4` in Excel 365 (Windows) and Excel for Mac; fix anything found; record results here.
2. **Phase 12:** UAT with 3 testers; turn findings into UX fixes (v1.0.x).
3. Phase 1 manual marketplace research (review-level).
4. Phase 14: name check, logo, cover; screenshots from the sample file.
5. Phase 13: build the Free tracker (generator variant) as the lead magnet.
6. Phase 15: PDF Quick-Start + User Guide; run `tools/package.py`; test the download.
7. Phases 16–25: publish listing, email funnel, soft launch → public launch.
