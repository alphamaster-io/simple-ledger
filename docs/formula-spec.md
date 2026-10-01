# Formula Specification

How the accounting model (`accounting-model.md`) is implemented in Excel. The source of truth is the generator in `src/simpleledger/`. This document explains the design so a reviewer can audit the workbook without reading Python.

## 1. Conventions

| Rule | Why |
|---|---|
| Excel-2010-era functions only (SUMIFS, COUNTIFS, INDEX/MATCH, SMALL/LARGE, SUMPRODUCT, EDATE, EOMONTH, IFERROR) | Compatibility. `test_formulas_avoid_functions_with_poor_compatibility` enforces it |
| No OFFSET / INDIRECT / TODAY in calculations | No volatile recalculation storms. TODAY() appears only as the fallback when the Dashboard month/year are blank |
| Defined names for every column range (`tx_Date`, `j_DrAmt`, `calc_Bal` …) and every scalar setting (`AsOf`, `FYStart`, `acctAR` …) | Readable formulas, e.g. `=SUMIFS(j_DrAmt,j_DrAcct,$C7,j_Per,3)` |
| `INDEX(range, n)` instead of direct cell links from the JOURNAL to source tables | Deleting a source row never produces `#REF!` |
| Every text lookup key is coerced with `&""` | An empty cell referenced through IF/INDEX returns 0, not "" |
| Status strings start with ⛔ / ⚠ / ✓ | One conditional-format rule per colour. COUNTIF with wildcards counts them |
| Amount columns use `#,##0.00;(#,##0.00);"-"` | Accounting style: negatives in parentheses, zero as a dash |

## 2. Register helper columns (TRANSACTIONS)

Visible automatic columns: Net Amount, Account Type, Debit Account, Credit Account, Check. Grouped helper columns (click **+** above the sheet to expand):

| Helper | Formula idea |
|---|---|
| Blank | 1 when Date, Type, Category, Amount and Money account are all empty |
| Dir | `INDEX(tt_Dir, MATCH(Type, tt_Type, 0))`: In / Out from LISTS |
| Other | the other side: Category, or the system account named in `tt_Other` (AR, AP, Capital, Drawings) |
| TaxEff | Tax if tax tracking is on, the cell is numeric and the type allows tax; else 0 |
| Err | first failing rule from accounting-model §4 (nested IF; Excel short-circuits) |
| Warn | first warning, only when Err is empty |
| Posted | 1 if not blank and no error |
| DrAmt / CrAmt / TaxDr / TaxCr | In: Dr = Amount, Cr = Amount − TaxEff, TaxCr = TaxEff. Out: Dr = Amount − TaxEff, TaxDr = TaxEff, Cr = Amount. Zero when not posted |

AR/AP follow the same pattern (Blank, TaxEff, Err, Warn, Posted, DrAcct, CrAcct, DrAmt, CrAmt, TaxAmt), plus visible Due Date Used, Paid Amount, Balance, Status, Days Outstanding, Days Overdue, Aging Bucket.

## 3. JOURNAL

3,000 rows: rows 1–2,000 mirror TRANSACTIONS, 2,001–2,500 mirror AR, 2,501–3,000 mirror AP.

| Column | Formula idea |
|---|---|
| Posted | `IFERROR(N(INDEX(src_Posted, n)), 0)` |
| Date, ID, Description, Party, Dr/Cr account | `IF(Posted=1, INDEX(src_col, n), "")` |
| Amounts | `IF(Posted=1, INDEX(src_amt, n), 0)` |
| Per (fiscal period) | 0 before FYStart, 13 after FYEnd, else `YEAR*12+MONTH − FYYM + 1` |
| DrCash / CrCash | 1 if that account's Cash Flow Section is "Cash" |
| CFGroup | Cash Flow Section of the non-cash side (blank if 0 or 2 cash sides) |
| CFMain / CFTax | cash in: +credit amount / +tax credit; cash out: −debit amount / −tax debit |
| Key | `Date*10000 + n` when the row touches the Ledger's selected account in its date range; `SMALL()` on this column yields entries in date order |

## 4. CALC

150 rows, one per COA row (`INDEX(coa_Name, i)`; later duplicates of a name are blanked). Columns: Type, Normal, CF, Statement, Sign, RawOpen, **Open** (OBE row = RawOpen − ΣRawOpen), **PreFY** (period 0), **M1–M12**, YTD (`SUMPRODUCT((month#<=RptIdx)*M1:M12)`), Cur, Prev, FY, **Bal** (= Open + PreFY + YTD), BalPrev (= Bal − Cur), report keys (`Type|rank`, `Cash|rank`, top-expense key), and **E1–E12** (balance at end of each fiscal month).

Net activity for one account and one criterion:

```text
SUMIFS(j_DrAmt, j_DrAcct, acct, crit) − SUMIFS(j_CrAmt, j_CrAcct, acct, crit)
  + IF(acct = acctTax, SUMIFS(j_TaxDr, j_Posted, 1, crit) − SUMIFS(j_TaxCr, j_Posted, 1, crit), 0)
```

Empty slots short-circuit with `IF(name="", 0, …)`, so the cost is proportional to the accounts actually used.

## 5. Reports

- **Slots.** Detail lines are fixed slots filled by `INDEX(calc_Name, MATCH("Revenue|k", calc_Key, 0))`. **Totals never use slots.** They are `SUMIFS(calc_col, calc_Type, type)`, so a report stays correct if a user adds more accounts than slots, and a ⚠ line says some accounts are not listed individually.
- Sign by statement: Revenue/Other Income/Liabilities/Equity are shown as −(Dr − Cr).
- BS prior-year profit: `−SUMPRODUCT((calc_Stmt="PL")*(calc_Open+calc_PreFY))`. Current-year profit: `−SUMPRODUCT((calc_Stmt="PL")*calc_YTD)`.
- Cash flow: `SUMIFS(j_CFMain, j_CFGroup, group, j_CFMain, ">0", j_Per, RptIdx)` and similar.
- Ledger running balance: previous line + sign(normal) × (debit − credit).

## 6. Named cells used by tests and the Dashboard

`PL_{Rev,COS,GP,Opex,NOI,OthInc,OthExp,Net,GM,NM}_{Cur,Prev,YTD,FY}`, `BS_{TotalAssets,TotalLiab,TotalEquity,TotalLE,Diff,PriorProfit,CYProfit,…}[_Prev]`, `BS_Status`, `CF_{Begin,Op,Inv,Fin,Net,End,EndBS,OpIn,OpOut}_{Cur,YTD}`, `CF_Status`, `TB_{Debit,Credit,Journal,Status}`, `DB_{Revenue,Expenses,Net,Cash,AR,AP,GM,NM,Errors,Warnings,ARAging,APAging}`, `ms_*` (monthly summary columns), `Led_{Open,Close}`.

## 7. Performance

About 115k formulas, 1.9 MB `.xlsx`. A full recalculation in LibreOffice headless takes about 7 s, including load and save. Excel recalculates on open (`fullCalcOnLoad`), because openpyxl stores no cached values.
