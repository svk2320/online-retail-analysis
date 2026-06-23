# Phase 2 — Data Quality Report
### Project: Online Retail Sales Analytics
### Tool: Microsoft Excel 2019
### Dataset: Online Retail II — UCI Machine Learning Repository

---

## What We Did in Phase 2

Built a Data Quality Summary sheet inside `01_raw_combined.xlsx` by running checks directly against the two raw sheet tabs — `Year 2009-2010` and `Year 2010-2011`.

---

## Dataset Overview

| Property | Detail |
|---|---|
| Source | UCI Machine Learning Repository |
| File | online_retail_II.xlsx |
| Sheets | Year 2009-2010, Year 2010-2011 |
| Total Rows | 1,065,371 |
| Columns | 8 |
| Date Range | 01-12-2009 to 09-12-2011 |

---

## Row Counts per Sheet

| Sheet | Rows |
|---|---|
| Year 2009-2010 | 523,461 |
| Year 2010-2011 | 541,910 |
| **Total Combined** | **1,065,371** |

---

## Null Count per Column

| Column | Total Rows | Null Count | Null % | Status |
|---|---|---|---|---|
| Invoice | 1,067,371 | 0 | 0% | ✅ Clean |
| StockCode | 1,067,371 | 0 | 0% | ✅ Clean |
| Description | 1,062,989 | 4,360 | 0.4% | ⚠️ Some missing |
| Quantity | 1,067,371 | 0 | 0% | ✅ Clean |
| InvoiceDate | 1,067,371 | 0 | 0% | ✅ Clean |
| Price | 1,067,371 | 0 | 0% | ✅ Clean |
| Customer ID | 824,364 | 241,958 | 29.4% | ⚠️ Guest purchases |
| Country | 1,067,371 | 0 | 0% | ✅ Clean |

---

## Additional Data Checks

| Check | Count | Notes |
|---|---|---|
| Cancellations (Invoice starts with C) | 19,441 | ~1.8% of transactions — normal for retail |
| Negative Quantities | 22,889 | Returns — need to separate from sales |
| Zero Prices | 6,175 | Free samples or data errors — exclude from analysis |
| Date From | 01-12-2009 | ✅ Expected |
| Date To | 09-12-2011 | ✅ Expected — 2 full years |

---

## Key Findings & Decisions

### 1. Customer ID — 241,958 Nulls (29.4%)
These are anonymous/guest purchases with no customer identity.
**Decision:** Keep these rows — they have valid revenue. Tag them as "Guest" using Customer_Type column in Power Query. Exclude from customer-level analysis only.

### 2. Description — 4,360 Nulls (0.4%)
Small number of rows with no product name.
**Decision:** Exclude from product analysis. Power Query cleaning handles this by filtering blank descriptions.

### 3. Cancellations — 19,441 rows
Invoices starting with "C" are cancellation/return transactions mixed with real sales.
**Decision:** Tag using Transaction_Type column in Power Query. Separate from sales in all analysis.

### 4. Negative Quantities — 22,889 rows
Returned items show as negative quantities.
**Decision:** Tag using Quantity_Type column. Filter out from main sales analysis.

### 5. Zero Prices — 6,175 rows
Free samples, test entries, or data errors with no monetary value.
**Decision:** Remove from dataset — these distort revenue and average price calculations.

---

## Excel Formulas Used

| Formula | Purpose |
|---|---|
| `=COUNTA('Year 2009-2010'!A2:A523462)+COUNTA('Year 2010-2011'!A2:A541911)-2` | Total rows combined |
| `=COUNTBLANK('Year 2009-2010'!G2:G523462)+COUNTBLANK('Year 2010-2011'!G2:G541911)` | Null count per column |
| `=C2/B2` | Null percentage |
| `=COUNTIF('Year 2009-2010'!A2:A523462,"C*")+COUNTIF('Year 2010-2011'!A2:A541911,"C*")` | Cancellation count |
| `=COUNTIF('Year 2009-2010'!D2:D523462,"<0")+COUNTIF('Year 2010-2011'!D2:D541911,"<0")` | Negative quantities |
| `=COUNTIF('Year 2009-2010'!F2:F523462,0)+COUNTIF('Year 2010-2011'!F2:F541911,0)` | Zero prices |
| `=MIN('Year 2009-2010'!E2:E523462)` | Earliest date |
| `=MAX('Year 2010-2011'!E2:E541911)` | Latest date |

---

## New Excel Skills Learned in Phase 2

| Skill | Used For |
|---|---|
| `COUNTBLANK()` | Finding null/empty cells |
| `COUNTA()` | Counting non-empty cells |
| `COUNTIF()` with wildcard `"C*"` | Finding cancellation invoices |
| `COUNTIF()` with condition `"<0"` | Finding negative quantities |
| `MIN()` / `MAX()` on date column | Finding date range |
| Exact range references (A2:A523462) | Avoiding whole-column errors |
| Click tab + click column header | Letting Excel write sheet references |

---

## What Happens Next — Phase 3

All issues found in Phase 2 are already fixed in Power Query (`01_raw_combined.xlsx`):

- ✅ Transaction_Type column added (Sale / Cancellation)
- ✅ Customer_Type column added (Known / Guest)
- ✅ Quantity_Type column added (Sale / Return)
- ✅ Revenue column added (Quantity × Price)
- ✅ Description trimmed of whitespace
- ✅ Zero and negative prices filtered out
- ✅ Zero and negative quantities filtered out
- ✅ Month, Year, Day columns extracted

**Next step:** Load cleaned Power Query to Data Model → Build Pivot Tables (Phase 4)

---

*Project 3 — Online Retail Sales Analytics*
*Phase 2 — Data Quality Report*
*Tool: Excel 2019*
