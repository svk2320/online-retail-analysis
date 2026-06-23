# Phase 5 — Power BI Dashboard Notes
### Project: Online Retail Sales Analytics
### Data Source: Excel 2019 → Power Query → Power BI Import Mode

---

## What We Did in Phase 5

---

## Step 1 — Connecting Power BI to Excel

### Problem faced
Power BI Navigator did not show `combined_data` from the Excel Data Model.
Power BI can only read Excel Tables or Excel sheets — not the Data Model directly.

### Solution
Loaded `combined_data` from Power Query as a Table to an Excel sheet:
- Named it `Cleaned_Combined_Data`
- Row count after cleaning: **1,041,671 rows**
- This is under Excel's row limit of 1,048,576 ✅

### How we connected
```
Power BI → Get Data → Excel Workbook
→ Selected online_retail_working.xlsx
→ Navigator → ticked Cleaned_Combined_Data
→ Transform Data → verified all 15 columns
→ Close & Apply
```

---

## Step 2 — Data Issues Fixed in Power BI Power Query

### Issue 1 — Invoice column stored as Number
**Error:** `DataFormat.Error: We couldn't convert to Number — A563185`
**Cause:** Cancellation invoices like C574950 cannot convert to Number type.
**Fix:**
```
Power Query Editor → Invoice column
→ Click type icon → change to Text
→ Replace Current
```

**Why Invoice must be Text:**
```
Normal invoices      → 574950   (looks like number)
Cancellation invoices→ C574950  (starts with C — not a number)
Keeping as Number crashes on any invoice starting with a letter
```

### Issue 2 — Two error rows found
**Rows:**
```
Row 60572  → Description: "Adjust bad debt" — accounting entry, not a sale
Row 139523 → Description: "Manual"          — manual adjustment, not a sale
```
**Fix:**
```
Selected Invoice column → Home → Remove Errors
```
**Impact:** Lost 2 rows out of 1,041,671 — completely insignificant.

---

## Step 3 — Two Table Strategy

### Why two tables
During cleaning in Excel, cancellations were filtered OUT of `Cleaned_Combined_Data`.
This means `Transaction_Type = "Cancellation"` rows do not exist in the clean table.

### Solution — Industry correct approach
```
Cleaned_Combined_Data    → Sales only (filtered, clean)
                         → Used for Pages 1, 2, 3, 5

Uncleaned_Combined_Data  → All rows (Sales + Cancellations flagged)
                         → Used for Page 4 (Cancellation Analysis) only
```

### How uncleaned table was built
- Combined both year tabs in Power Query
- Added `Transaction_Type` column only (Sale / Cancellation flag)
- No other filtering applied
- Loaded to Power BI as second table

---

## Step 4 — Columns in Each Table

### Cleaned_Combined_Data (15 columns)

| Column | Type | Source | Notes |
|--------|------|--------|-------|
| Invoice | Text | Original | Changed from Number to Text |
| StockCode | Text | Original | |
| Description | Text | Original | Whitespace trimmed |
| Quantity | Number | Original | Filtered > 0 |
| InvoiceDate | DateTime | Original | |
| Price | Decimal | Original | Filtered > 0 |
| Customer ID | Number | Original | Nulls kept, flagged |
| Country | Text | Original | |
| Transaction_Type | Text | Derived | Sale only (cancellations removed) |
| Customer_Type | Text | Derived | Guest or Known |
| Quantity_Type | Text | Derived | Sale or Return |
| Revenue | Decimal | Derived | Quantity × Price |
| Month | Number | Derived | 1-12 |
| Year | Number | Derived | 2009/2010/2011 |
| Day | Text | Derived | Monday-Sunday |

### Uncleaned_Combined_Data (key columns)

| Column | Type | Notes |
|--------|------|-------|
| Invoice | Text | Includes C-prefixed cancellation invoices |
| Transaction_Type | Text | Sale or Cancellation |
| Revenue | Decimal | Fixed type — was Text by default |
| All original columns | Various | No filtering applied |

---

## Step 5 — DAX Measures Created

All measures created by right-clicking table name in Fields panel → New Measure.

```dax
Total Revenue = 
SUM('Cleaned_Combined_Data'[Revenue])

Total Orders = 
DISTINCTCOUNT('Cleaned_Combined_Data'[Invoice])

Avg Order Value = 
DIVIDE([Total Revenue], [Total Orders], 0)

Total Customers = 
DISTINCTCOUNT('Cleaned_Combined_Data'[Customer ID])

Total Cancellations = 
CALCULATE(
    DISTINCTCOUNT('Uncleaned_Combined_Data'[Invoice]),
    'Uncleaned_Combined_Data'[Transaction_Type] = "Cancellation"
)

Cancel Rate = 
DIVIDE(
    CALCULATE(
        DISTINCTCOUNT('Uncleaned_Combined_Data'[Invoice]),
        'Uncleaned_Combined_Data'[Transaction_Type] = "Cancellation"
    ),
    DISTINCTCOUNT('Uncleaned_Combined_Data'[Invoice]),
    0
)

Lost Revenue = 
CALCULATE(
    SUM('Uncleaned_Combined_Data'[Revenue]),
    'Uncleaned_Combined_Data'[Transaction_Type] = "Cancellation"
)
```

---

## Step 6 — Dashboard Pages Built

### Page 1 — Revenue Overview
```
4 KPI Cards    : Total Revenue, Total Orders, Avg Order Value, Total Customers
Line Chart     : Monthly Revenue Trend (Month vs Revenue, Legend = Year)
Donut Chart    : Sales vs Cancellations Split
Bar Chart      : Revenue by Country — Top 10
```
**Table used:** Cleaned_Combined_Data

---

### Page 2 — Product Analysis
```
Bar Chart  : Top 20 Products by Revenue (horizontal)
Bar Chart  : Top 20 Products by Quantity Sold (horizontal)
Table      : Product Details — Revenue, Qty, Avg Price
             Conditional formatting on Revenue (white → blue)
Slicer     : Filter by Country (filters all visuals on page)
```
**Table used:** Cleaned_Combined_Data

---

### Page 3 — Customer Analysis
```
3 KPI Cards : Total Customers, Avg Customer Spend, Top Customer Revenue
Bar Chart   : Top 20 Customers by Revenue
Bar Chart   : Revenue — Known vs Guest Customers
Table       : Top 20 Customer Details
             Conditional formatting on Revenue
```
**Table used:** Cleaned_Combined_Data
**Filter applied:** Customer_Type = "Known" on customer cards and charts

---

### Page 4 — Cancellation Analysis
```
3 KPI Cards : Total Cancellations, Cancel Rate %, Lost Revenue
Line Chart  : Monthly Sales vs Cancellation Trend
Bar Chart   : Cancellation Rate by Country — Top 10
Bar Chart   : Top 10 Most Cancelled Products
```
**Table used:** Uncleaned_Combined_Data
**Why:** Cleaned table has no cancellation rows — uncleaned table has both Sale and Cancellation

---

### Page 5 — Country & Time Analysis
```
Bar Chart  : Total Revenue by Country — Top 15
Bar Chart  : Orders by Day of Week
Line Chart : Revenue by Month — Year on Year (Legend = Year)
Table      : Country Summary — Revenue, Orders, Qty, Avg Price
             Conditional formatting on Revenue
```
**Table used:** Cleaned_Combined_Data

---

## Step 7 — Formatting Applied

### Conditional Formatting on Tables
```
Click Table visual
→ Format pane (paint roller icon)
→ Visual tab
→ Cell elements
→ Revenue → Background color → On
→ fx button → Gradient
→ Minimum: white (#FFFFFF)
→ Maximum: blue (#1F4E79)
→ OK
```

### Page Navigation
```
Insert → Buttons → Navigator → Page Navigator
Placed at top of every page
```

---

## Problems Faced & Solutions

| Problem | Cause | Solution |
|---------|-------|----------|
| combined_data not visible in Power BI Navigator | Excel Data Model not readable by Power BI | Loaded as Excel Table to sheet |
| DataFormat.Error on Invoice | Invoice loaded as Number, C-prefixed values can't convert | Changed Invoice type to Text |
| 2 error rows (Adjust bad debt, Manual) | Non-transaction accounting entries | Removed via Home → Remove Errors |
| Transaction_Type only showed Sale | Cancellations filtered out during Excel cleaning | Created separate Uncleaned table with both Sale and Cancellation |
| Lost Revenue measure error — SUM on String | Revenue column in uncleaned table loaded as Text | Changed Revenue type to Decimal Number in Power Query |

---

## Key Learnings from Phase 5

| Learning | Detail |
|----------|--------|
| Power BI cannot read Excel Data Model | Must load as Excel Table or Sheet |
| Always set data types manually | Power BI auto-detection is unreliable for mixed columns |
| Two table strategy | Clean table for analysis, raw table for exception/cancellation analysis |
| Count Distinct vs Count | Use Count Distinct for invoices — same invoice has multiple product rows |
| DAX measures vs drag and drop | Measures give more control — use for Cancel Rate, Lost Revenue |
| Uncleaned table for Page 4 | Cancellation analysis needs raw data — never mix with clean data |

---

## Industry Workflow Used

```
Raw Data (online_retail_II.xlsx)
    ↓
Excel Power Query
    → Combined both year tabs
    → Added derived columns (Revenue, Transaction_Type, Customer_Type etc.)
    → Cleaned data (filtered zero prices, zero quantities)
    → Loaded to Excel sheet (Cleaned_Combined_Data)
    ↓
Power BI
    → Imported Cleaned_Combined_Data (sales analysis)
    → Imported Uncleaned_Combined_Data (cancellation analysis)
    → Built 5 page dashboard
    → No SQL, No Python — Excel + Power BI only
```

---

*Project 3 — Online Retail Sales Analytics*
*Phase 5 — Power BI Dashboard*
*Tools: Excel 2019 + Power BI Desktop*