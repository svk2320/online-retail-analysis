# Phase 1 — Data Understanding & Combining
### Project: Online Retail Sales Analytics
### Tool: Microsoft Excel 2019 — Power Query Editor
### Dataset: Online Retail II — UCI Machine Learning Repository

---

## What We Did in Phase 1

Loaded the raw dataset into Excel, understood the structure of both tabs, and combined them into a single query using Power Query — ready for data quality checks in Phase 2.

---

## Dataset Overview

| Property | Detail |
|---|---|
| Source | UCI Machine Learning Repository |
| File | online_retail_II.xlsx |
| Sheets | Year 2009-2010, Year 2010-2011 |
| Total Rows | ~1,065,371 (before cleaning) |
| Columns | 8 |
| Date Range | 01-12-2009 to 09-12-2011 |

---

## Step 1 — Open the Raw File and Understand the Structure

Opened `online_retail_II.xlsx` directly in Excel first — before touching Power Query.

Both tabs have the exact same 8 columns:

| Column | Data Type | What it means |
|--------|-----------|---------------|
| Invoice | Text / Number | Invoice number — if it starts with `C`, it is a cancellation |
| StockCode | Text | Unique product code |
| Description | Text | Product name |
| Quantity | Number | Units sold — negative values mean returns |
| InvoiceDate | DateTime | Date and time of the transaction |
| Price | Decimal | Unit price in GBP (£) |
| Customer ID | Number | Unique customer identifier — has blank cells (guest purchases) |
| Country | Text | Country of the customer |

**Key observation:** Both tabs have identical structure — same columns, same order.
This means they can be safely combined without any column mapping or transformation.
Tab 1 covers December 2009 to December 2010.
Tab 2 covers December 2010 to December 2011.
Together they form two complete years of transaction data.

---

## Step 2 — Load Raw Data into Power Query

```
Data tab → Get Data → From File → From Workbook
→ Navigated to online_retail_II.xlsx
→ Navigator window opened showing both tabs
→ Checked "Select multiple items" checkbox at the top
→ Ticked both: Year 2009-2010 and Year 2010-2011
→ Clicked Transform Data (not Load — we want to see it first)
```

**Why Transform Data and not Load?**
Clicking Load sends data straight to Excel as-is.
Clicking Transform Data opens Power Query Editor first.
Always use Transform Data — it lets you check the data before committing.

**What happened:** Power Query Editor opened with two separate queries in the left panel — one for each tab.

---

## Step 3 — Check Row Counts per Tab

Before combining, verified row counts for each tab individually in Power Query.

**How to check row count in Power Query:**
```
Bottom status bar shows: "X rows | 8 columns"
Visible as soon as you click a query in the left panel
```

| Tab | Row Count |
|-----|-----------|
| Year 2009-2010 | 523,461 |
| Year 2010-2011 | 541,910 |
| **Expected Combined** | **1,065,371** |

This check is important.
After combining, if the total doesn't match the sum of both tabs, rows were lost or duplicated during the append — a sign something went wrong.

---

## Step 4 — Combine Both Tabs into One Query

```
In Power Query Editor:
Home tab → Append Queries → Append Queries as New
→ Selected "Two tables"
→ First table:  Year 2009-2010
→ Second table: Year 2010-2011
→ Clicked OK
→ New query created automatically
→ Renamed the new query: combined_data
```

**Why Append and not Merge?**
Append = stack rows on top of each other (same columns, more rows) — this is what we want.
Merge = join tables side by side on a matching column (like SQL JOIN) — not what we want here.
Both tabs have the same structure, so Append is the correct operation.

**Why "as New"?**
Append as New creates a third query that combines the two.
The original two tab queries are preserved untouched.
This is safer — original queries remain available if needed.

**Row count check after combining:**
```
Bottom status bar on combined_data query: 1,065,371 rows | 8 columns
```
✅ Matches the sum of both tabs — no rows lost or duplicated.

---

## Step 5 — Check Column Data Types

After combining, checked data types assigned by Power Query for each column.

```
Click the icon to the left of each column header to see/change the type
ABC = Text
123 = Whole Number
1.2 = Decimal Number
📅 = Date/Time
```

| Column | Auto-detected Type | Correct? |
|--------|-------------------|----------|
| Invoice | Whole Number | ❌ Should be Text — C-prefixed cancellations will error |
| StockCode | Text | ✅ |
| Description | Text | ✅ |
| Quantity | Whole Number | ✅ |
| InvoiceDate | Date/Time | ✅ |
| Price | Decimal Number | ✅ |
| Customer ID | Decimal Number | ✅ (nulls allowed) |
| Country | Text | ✅ |

**Invoice column issue noted here — fixing it is Phase 3's job.**
In Phase 1 we just observe and document.
The Invoice column being auto-detected as Number is a known problem — any invoice starting with `C` (cancellations) will throw a conversion error when analysis runs.

---

## Step 6 — Quick Column Profiling

Enabled column profiling to get a fast overview of data quality before writing any formulas.

```
View tab in Power Query Editor
→ Check: Column Quality
→ Check: Column Distribution
→ Check: Column Profile
```

**Column Quality** shows three bars per column:
- Valid — rows with a proper value
- Error — rows Power Query couldn't parse
- Empty — blank/null rows

**Column Distribution** shows a mini histogram — useful for spotting outliers at a glance.

**Column Profile** (bottom panel) shows: distinct count, min, max, null count, value distribution.

**Important note:**
By default Power Query profiles only the first 1,000 rows.
For a 1M+ row dataset this is not representative.

```
To profile the full dataset:
Bottom left of Power Query → click "Column profiling based on top 1000 rows"
→ Change to "Column profiling based on entire data set"
```

This makes null counts and distributions accurate for the full 1,065,371 rows.

---

## Step 7 — Close Without Loading

Phase 1 ends here. The combined_data query exists in Power Query but we do not load it to Excel yet.

```
Home → Close & Load → Close & Load To...
→ Only create connection (not to sheet)
→ OK
```

**Why only connection?**
Loading 1M+ rows to an Excel sheet at this stage is unnecessary and slow.
Phase 2 runs quality checks directly against the raw tabs using Excel formulas — not the combined query.
The combined query will be loaded to a sheet only after cleaning is complete in Phase 3.

---

## Summary — What Was Accomplished in Phase 1

| Task | Status |
|------|--------|
| Understood both tab structures | ✅ |
| Verified both tabs have identical columns | ✅ |
| Loaded both tabs into Power Query | ✅ |
| Verified row counts per tab | ✅ |
| Combined both tabs using Append as New | ✅ |
| Verified combined row count (1,065,371) | ✅ |
| Checked auto-detected data types | ✅ |
| Noted Invoice column type issue for Phase 3 | ✅ |
| Ran full column profiling | ✅ |
| Saved as connection only | ✅ |

---

## New Excel / Power Query Skills Learned in Phase 1

| Skill | What it does |
|-------|-------------|
| Get Data → From File → From Workbook | Loads an Excel file into Power Query |
| Select multiple items in Navigator | Loads multiple tabs at once instead of one at a time |
| Transform Data vs Load | Opens Power Query Editor instead of dumping data straight to sheet |
| Append Queries as New | Stacks rows from two tables into one new query |
| Append vs Merge | Append = more rows (union). Merge = more columns (join) |
| Column Quality / Distribution / Profile | Fast data overview without writing any formulas |
| Profile based on entire dataset | Changes profiling from first 1,000 rows to all rows |
| Close & Load → Only Create Connection | Keeps query available without loading rows to sheet |
| Data type icons (ABC / 123 / 1.2 / 📅) | Read and change column data types in Power Query |

---

## What Happens Next — Phase 2

With the combined query ready, Phase 2 moves to the Excel sheet to build a Data Quality Report.

Using Excel formulas (`COUNTBLANK`, `COUNTIF`, `MIN`, `MAX`) directly against the two raw tabs to quantify:
- Null counts per column
- Cancellation count
- Negative quantity count
- Zero price count
- Date range

These findings drive every cleaning decision made in Phase 3.

---

*Project 3 — Online Retail Sales Analytics*
*Phase 1 — Data Understanding & Combining*
*Tool: Excel 2019 + Power Query*
