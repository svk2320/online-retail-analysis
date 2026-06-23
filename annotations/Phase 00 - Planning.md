# Project 3 — Online Retail Sales Analytics
### Tools: Microsoft Excel + Power BI (No Python, No SQL)
### Dataset: Online Retail II — UCI Machine Learning Repository

---

## Dataset Overview

| Property | Detail |
|----------|--------|
| Source | UCI Machine Learning Repository |
| File type | Excel (.xlsx) |
| Tabs | Year 2009-2010, Year 2010-2011 |
| Rows | ~500K-1M per tab (~1M+ total) |
| Columns | 8 |
| Type | Real UK online retailer transaction data |

### Columns
| Column | What it means |
|--------|--------------|
| Invoice | Invoice number — starts with C = cancellation |
| StockCode | Product code |
| Description | Product name |
| Quantity | Units sold (negative = return/cancellation) |
| InvoiceDate | Date and time of transaction |
| Price | Unit price in GBP |
| Customer ID | Unique customer identifier (has nulls) |
| Country | Customer country |

---

## Project Structure

online-retail-analysis/
├── data/
│   └── online_retail_II.xlsx              ← raw, never touch
├── excel/
│   └── online_retail_working.xlsx         ← ONE file, everything here
│       ├── Sheet: Data_Quality_Report     ← Phase 2 summary
│       ├── Sheet: Cleaned_Combined_Data   ← Phase 3 clean data
│       ├── Power Query: combined_data     ← cleaning pipeline
│       └── Power Query: uncleaned_data    ← raw + Transaction_Type only
└── powerbi/
    └── online_retail_dashboard.pbix       ← connects to excel file
        ├── Page 1: Revenue Overview
        ├── Page 2: Product Analysis
        ├── Page 3: Customer Analysis
        ├── Page 4: Cancellation Analysis
        └── Page 5: Country & Time Analysis

---

## Phase 1 — Data Understanding & Combining

**What we do:**
- Understand both tabs (same structure, different years)
- Combine both tabs into one sheet using Power Query
- Check row counts before and after combining
- Understand what each column means

**New Excel skills learned:**
- Power Query — Get Data → combine sheets
- Data types — text vs number vs date
- Quick column profiling

**Output:** `online_retail_working.xlsx` — one sheet, all rows

---

## Phase 2 — Data Quality Report

**What we do:**
Build a Data Quality Summary sheet — same concept as your `online_retail_working.py` but in Excel.

### Checks we run:

| Check | How in Excel | Expected |
|-------|-------------|----------|
| Null count per column | COUNTBLANK() | 0 ideally |
| Negative quantities | COUNTIF(D:D,"<0") | Should investigate |
| Negative prices | COUNTIF(F:F,"<0") | Should be 0 |
| Zero prices | COUNTIF(F:F,0) | Should investigate |
| Cancellations (Invoice starts C) | COUNTIF(A:A,"C*") | Need to separate |
| Duplicate invoices | COUNTIF + helper column | Should be 0 |
| Date range | MIN/MAX on InvoiceDate | Should be 2009-2011 |
| Unique customers | COUNTA - COUNTBLANK | Reference number |
| Unique countries | Pivot table | Reference number |

**New Excel skills learned:**
- COUNTIF, COUNTBLANK, COUNTA
- MIN, MAX on dates
- Wildcard in COUNTIF ("C*" finds all starting with C)

**Output:** Data Quality Summary sheet inside `online_retail_working.xlsx`

---

## Phase 3 — Data Cleaning

This is the main phase. Here's every problem and exactly how to fix it in Excel.

### Problem 1 — Null Customer IDs (~135K rows)
**How to find:** COUNTBLANK(G:G)
**What to do:**
- Filter column G → Blanks
- These are guest/anonymous purchases
- Add helper column: =IF(G2="","Guest","Known")
- Do NOT delete — keep for revenue analysis, exclude for customer analysis

**Excel skill:** IF formula, Filter

---

### Problem 2 — Cancellations mixed with Sales
**How to find:** COUNTIF(A:A,"C*")
**What to do:**
- Add helper column: =IF(LEFT(A2,1)="C","Cancellation","Sale")
- This separates returns from real sales
- Use this column to filter in all analysis

**Excel skill:** LEFT(), IF(), text functions

---

### Problem 3 — Negative Quantities
**How to find:** COUNTIF(D:D,"<0")
**What to do:**
- Negative quantity = return or cancellation
- These match with Cancellation invoices (Problem 2)
- Add column: =IF(D2<0,"Return","Sale")
- Cross-check with Invoice column

**Excel skill:** COUNTIF with conditions, nested IF

---

### Problem 4 — Zero or Negative Prices
**How to find:** COUNTIF(F:F,"<=0")
**What to do:**
- Zero price = free sample, test order, or data error
- Add flag column: =IF(F2<=0,"Bad Price","OK")
- Exclude from revenue calculations

**Excel skill:** COUNTIF, IF

---

### Problem 5 — Whitespace in Description
**How to find:** Visual scan — some descriptions have leading spaces
**What to do:**
- Add clean column: =TRIM(C2)
- TRIM removes leading/trailing spaces
- Replace original with cleaned values (Paste Special → Values)

**Excel skill:** TRIM(), Paste Special

---

### Problem 6 — Duplicate Invoices
**How to find:** 
- Add helper: =COUNTIF(A:A,A2) → if >1, it's a duplicate
- Or: Data → Remove Duplicates
**What to do:**
- Check if duplicates are real duplicates or same invoice different products
- Same Invoice + Same StockCode = true duplicate → remove
- Same Invoice + Different StockCode = normal (one order, multiple items) → keep

**Excel skill:** COUNTIF for duplicates, Remove Duplicates

---

### Problem 7 — Revenue Column (derived)
**What to do:**
- No revenue column exists — calculate it
- Add column: =D2*F2 (Quantity × Price)
- Name it: Revenue
- This is the most important derived column

**Excel skill:** Simple formula, derived columns

---

### Problem 8 — Date Parts (derived)
**What to do:**
- Extract month, year, day of week from InvoiceDate
- Month: =MONTH(E2)
- Year: =YEAR(E2)
- Day of week: =TEXT(E2,"dddd")
- Month-Year label: =TEXT(E2,"MMM-YYYY")

**Excel skill:** Date functions — MONTH, YEAR, TEXT

---

### Final Clean Data Rules
Keep a row if ALL of these are true:
- Invoice does NOT start with C (not a cancellation)
- Quantity > 0
- Price > 0
- Description is not blank

**Output:** `online_retail_working.xlsx` — clean, ready for analysis

---

## Phase 4 — Excel Analysis

**Status: Skipped — Industry Decision**

Reason: Dataset contains 1,041,671 rows which exceeds
practical limits for Excel pivot table performance.
Industry standard for datasets above 500K rows is to
move directly from Power Query → Power BI.

Excel pivot tables are retained in this documentation
as a reference for smaller datasets.

### 4A — Summary Statistics Sheet
Build a stats table covering:

| Metric | Formula |
|--------|---------|
| Total rows (raw) | COUNTA(A:A)-1 |
| Total rows (clean) | From cleaned sheet |
| Total revenue | SUM(Revenue column) |
| Avg order value | AVERAGEIF |
| Total customers | COUNTUNIQUE (or pivot) |
| Total countries | Pivot table |
| Date range | MIN/MAX |
| Top country | INDEX/MATCH on pivot |

---

### 4B — Pivot Tables (6 pivots)

**Pivot 1 — Revenue by Month**
```
Rows    : Month-Year
Values  : Revenue → Sum
Sort    : by date ascending
```

**Pivot 2 — Revenue by Country**
```
Rows    : Country
Values  : Revenue → Sum, Orders → Count
Sort    : by Revenue descending
Filter  : Top 10
```

**Pivot 3 — Top 20 Products by Revenue**
```
Rows    : Description
Values  : Revenue → Sum, Quantity → Sum
Sort    : by Revenue descending
Filter  : Top 20
```

**Pivot 4 — Revenue by Day of Week**
```
Rows    : Day of Week
Values  : Revenue → Sum, Orders → Count
Sort    : by Day number
```

**Pivot 5 — Customer Analysis**
```
Rows    : Customer ID
Values  : Revenue → Sum, Orders → Count
Sort    : by Revenue descending
Filter  : Top 20 customers
```

**Pivot 6 — Cancellation Analysis**
```
Rows    : Month-Year
Values  : Cancellations → Count, Sales → Count
Calculated field: Cancel Rate = Cancellations/Sales
```

---

### 4C — Charts (5 charts)

| Chart | Type | Data source |
|-------|------|-------------|
| Monthly Revenue Trend | Line chart | Pivot 1 |
| Revenue by Country | Horizontal bar | Pivot 2 |
| Top 20 Products | Horizontal bar | Pivot 3 |
| Revenue by Day of Week | Column chart | Pivot 4 |
| Cancellation Rate Trend | Line chart | Pivot 6 |

**New Excel skills learned:**
- Pivot Tables — the most important Excel skill
- Pivot Charts
- Slicers — interactive filters on pivot tables
- Calculated fields in pivot tables

---

## Phase 5 — Power BI Dashboard

Connect Power BI to `online_retail_working.xlsx` — Import mode.

### 5 Pages:

**Page 1 — Revenue Overview**
```
KPI Cards  : Total Revenue, Total Orders, Avg Order Value, Total Customers
Line Chart : Monthly Revenue Trend
Bar Chart  : Revenue by Country (Top 10)
Donut      : Sales vs Cancellations split
```

**Page 2 — Product Analysis**
```
Bar Chart  : Top 20 Products by Revenue
Bar Chart  : Top 20 Products by Quantity
Table      : Product details with revenue, qty, avg price
Slicer     : Filter by Country
```

**Page 3 — Customer Analysis**
```
KPI Cards  : Total Customers, Avg Customer Spend, Top Customer Revenue
Bar Chart  : Top 20 Customers by Revenue
Bar Chart  : New vs Repeat customers by month
Table      : Top 20 customer details
```

**Page 4 — Cancellation Analysis**
```
KPI Cards  : Total Cancellations, Cancellation Rate %, Lost Revenue
Line Chart : Monthly Cancellation Trend
Bar Chart  : Cancellation Rate by Country
Bar Chart  : Top Cancelled Products
```

**Page 5 — Country & Time Analysis**
```
Bar Chart  : Revenue by Country
Bar Chart  : Orders by Day of Week
Line Chart : Revenue by Hour of Day
Table      : Country summary — revenue, orders, avg value
```

---

## What You'll Learn by End of Project 3

| Skill | New in Project 3 |
|-------|-----------------|
| Excel COUNTIF, COUNTBLANK | ✅ |
| Excel IF, nested IF | ✅ |
| Excel text functions — TRIM, LEFT, TEXT | ✅ |
| Excel date functions — MONTH, YEAR | ✅ |
| Excel derived columns | ✅ |
| Excel Pivot Tables | ✅ Most important |
| Excel Pivot Charts | ✅ |
| Excel Slicers | ✅ |
| Power Query — combine sheets | ✅ |
| Power BI from Excel source | ✅ |
| Data cleaning without Python | ✅ |
| Business analysis workflow | ✅ |

---

## Key Difference from Previous Projects

| | Project 1 & 2 | Project 3 |
|-|--------------|-----------|
| Cleaning tool | Python + SQL | Excel formulas |
| Analysis tool | SQL + Python | Pivot Tables |
| Charts | Plotly | Excel Charts + Power BI |
| Database | PostgreSQL | No database — flat file |
| Workflow | Code → run → result | Click → drag → result |

This is how a business analyst works day to day.
No code. No terminal. Just Excel and Power BI.

---

*Project 3 — Online Retail Sales Analytics*
*Tools: Excel 2019 + Power BI*
