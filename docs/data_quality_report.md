# Data Quality Report

## Project: Online Retail Sales Analytics
**Dataset:** Online Retail II — UCI Machine Learning Repository
**Tool:** Microsoft Excel 2019 + Power Query
**Prepared:** Phase 2 — Data Quality Assessment

---

## Dataset Summary

| Metric | Value |
|--------|-------|
| Source | UCI Machine Learning Repository |
| File | online_retail_II.xlsx |
| Total Rows | 1,065,371 |
| Columns | 8 |
| Period | 01 Dec 2009 – 09 Dec 2011 |
| Tab 1 (Year 2009-2010) | 523,461 rows |
| Tab 2 (Year 2010-2011) | 541,910 rows |

---

## Column Schema

| Column | Type | Description |
|--------|------|-------------|
| Invoice | Text | Invoice number — prefix `C` indicates cancellation |
| StockCode | Text | Unique product code |
| Description | Text | Product name |
| Quantity | Number | Units sold — negative values indicate returns |
| InvoiceDate | DateTime | Transaction date and time |
| Price | Decimal | Unit price in GBP (£) |
| Customer ID | Number | Unique customer identifier — nullable (guest purchases) |
| Country | Text | Customer country |

---

## Null Analysis

| Column | Non-Null Rows | Null Count | Null % | Status |
|--------|-----------|------------|--------|--------|
| Invoice | 1,065,371 | 0 | 0.00% | ✅ Clean |
| StockCode | 1,065,371 | 0 | 0.00% | ✅ Clean |
| Description | 1,061,011 | 4,360 | 0.41% | ⚠️ Minor — some missing |
| Quantity | 1,065,371 | 0 | 0.00% | ✅ Clean |
| InvoiceDate | 1,065,371 | 0 | 0.00% | ✅ Clean |
| Price | 1,065,371 | 0 | 0.00% | ✅ Clean |
| Customer ID | 823,413 | 241,958 | 22.71% | ⚠️ High — guest purchases |
| Country | 1,065,371 | 0 | 0.00% | ✅ Clean |

---

## Quality Findings

| Issue | Count | % of Dataset | Decision |
|-------|-------|-------------|----------|
| Null Customer IDs | 241,958 | 22.7% | Keep — tag as Guest, exclude from customer-level analysis only |
| Null Descriptions | 4,360 | 0.4% | Exclude from product analysis |
| Cancellations | 19,441 | 1.8% | Tag as Cancellation — analyze separately on Page 4 |
| Negative Quantities | 22,889 | 2.1% | Tag as Return — filter out from main sales analysis |
| Zero Prices | 6,175 | 0.6% | Remove — free samples or data errors; distort revenue metrics |

---

## Issue Detail

### 1. Null Customer IDs — 241,958 rows (22.7%)

Anonymous guest purchases with no customer identity attached.

**Decision:** Retain these rows — they carry valid revenue data. Tag using a derived `Customer_Type` column (`Guest` / `Known`). Exclude from customer-level analysis only. Include in all revenue and product analysis.

**Formula used:**
```excel
=COUNTBLANK('Year 2009-2010'!G2:G523462) + COUNTBLANK('Year 2010-2011'!G2:G541911)
```

---

### 2. Null Descriptions — 4,360 rows (0.4%)

Small set of rows with no product name recorded.

**Decision:** Remove rows with blank descriptions during Power Query cleaning. Product-level analysis requires a valid product name, so these rows are excluded from downstream reporting.

**Formula used:**
```excel
=COUNTBLANK('Year 2009-2010'!C2:C523462) + COUNTBLANK('Year 2010-2011'!C2:C541911)
```

---

### 3. Cancellations — 19,441 rows (1.8%)

Invoices prefixed with `C` (e.g., `C489434`) represent cancellation or return transactions mixed within the sales data.

**Decision:** Tag using a derived `Transaction_Type` column (`Sale` / `Cancellation`). Separate from sales in all main analysis. A dedicated uncleaned table (`Uncleaned_Combined_Data`) is used in Power BI Page 4 for cancellation-specific analysis.

**Formula used:**
```excel
=COUNTIF('Year 2009-2010'!A2:A523462,"C*") + COUNTIF('Year 2010-2011'!A2:A541911,"C*")
```

---

### 4. Negative Quantities — 22,889 rows (2.1%)

Returned items are recorded as negative quantity values. These represent stock being returned to the retailer.

**Decision:** Tag using a derived `Quantity_Type` column (`Sale` / `Return`). Filter out from main sales analysis. Return rate tracked as a separate business metric.

**Formula used:**
```excel
=COUNTIF('Year 2009-2010'!D2:D523462,"<0") + COUNTIF('Year 2010-2011'!D2:D541911,"<0")
```

---

### 5. Zero Prices — 6,175 rows (0.6%)

Rows where unit price is zero — likely free samples, test entries, or data entry errors with no monetary value.

**Decision:** Remove entirely. Zero-price rows distort revenue totals and average price calculations. Not recoverable — no basis for imputation.

**Formula used:**
```excel
=COUNTIF('Year 2009-2010'!F2:F523462,0) + COUNTIF('Year 2010-2011'!F2:F541911,0)
```

---

## Date Range Validation

| Check | Value | Status |
|-------|-------|--------|
| Earliest date | 01 Dec 2009 | ✅ Expected |
| Latest date | 09 Dec 2011 | ✅ Expected — approximately 2 full years |

**Formulas used:**
```excel
=MIN('Year 2009-2010'!E2:E523462)
=MAX('Year 2010-2011'!E2:E541911)
```

---

## Cleaning Decisions Summary

| Column | Issue | Action |
|--------|-------|--------|
| Invoice | Auto-detected as Number; C-prefixed values break conversion | Changed type to Text in Power Query |
| Description | 4,360 nulls; whitespace causing duplicate product entries | Trimmed whitespace and removed blank descriptions from cleaned dataset |
| Quantity | 22,889 negative (returns); zero entries | Tagged as Return; filtered Quantity ≤ 0 from main dataset |
| Price | 6,175 zero prices | Filtered Price ≤ 0 — removed entirely |
| Customer ID | 241,958 nulls (guest purchases) | Kept; tagged as Guest via Customer_Type column |
| Revenue | Not present in raw data | Derived: `Quantity × Price` |

**Rows retained after cleaning:** 1,041,671 (out of 1,065,371)

**Rows removed:** 23,700

*Note: Rows removed are not equal to the sum of individual issue counts because some records contained multiple data quality issues and were counted in more than one category.*

---

## Power Query Transformation Steps

All 12 steps are recorded in Power Query and re-run automatically on Refresh.

| Step | Action | Detail |
|------|--------|--------|
| 1 | Load raw tabs | Loaded Year 2009-2010 and Year 2010-2011 into Power Query via Get Data → From Workbook |
| 2 | Append tabs | Appended as New into `combined_data` query — 1,065,371 rows verified |
| 3 | Fix Invoice type | Changed Invoice from Number → Text to support `C`-prefix detection |
| 4 | Add `Transaction_Type` | `if Text.StartsWith([Invoice], "C") then "Cancellation" else "Sale"` |
| 5 | Add `Customer_Type` | `if [Customer ID] = null then "Guest" else "Known"` |
| 6 | Add `Quantity_Type` | `if [Quantity] < 0 then "Return" else "Sale"` |
| 7 | Add `Revenue` | `[Quantity] * [Price]` — primary metric column |
| 8 | Clean `Description` | Trimmed whitespace and filtered blank descriptions to prevent duplicate or unidentified product entries in analysis |
| 9 | Extract `Month` | Add Column → Date → Month → Month (returns 1–12) |
| 10 | Extract `Year` | Add Column → Date → Year → Year (returns 2009/2010/2011) |
| 11 | Extract `Day` | Add Column → Date → Day → Name of Day (returns Monday–Sunday) |
| 12 | Filter rows | Price > 0 AND Quantity > 0 — removes zero/negative price and quantity rows |

---

## Column Count Before vs After

| Stage | Column Count | Columns Added |
|-------|-------------|---------------|
| Raw data | 8 | — |
| After cleaning | 15 | Transaction_Type, Customer_Type, Quantity_Type, Revenue, Month, Year, Day |

**Raw columns:** Invoice, StockCode, Description, Quantity, InvoiceDate, Price, Customer ID, Country
**Derived columns:** Transaction_Type, Customer_Type, Quantity_Type, Revenue, Month, Year, Day

---

## Two-Table Strategy (Power BI)

Cancellation analysis requires the unfiltered dataset including rows that were removed during cleaning (cancellations, returns). Two separate tables are loaded into Power BI to handle this:

| Table | Rows | Contents | Used For |
|-------|------|----------|----------|
| `Cleaned_Combined_Data` | 1,041,671 | Sales only — filtered, 15 columns | Pages 1, 2, 3, 5 (Revenue, Products, Customers, Geography) |
| `Uncleaned_Combined_Data` | 1,065,371 | All rows including cancellations | Page 4 (Cancellation Analysis) |

This is the correct pattern — clean and raw data are never mixed in the same analysis context.

---

*Online Retail Sales Analytics*
*Phase 2 — Data Quality Report*
*Dataset: Online Retail II — UCI Machine Learning Repository*
