# Power Query Cleaning Notes — Online Retail Project
### Tool: Microsoft Excel 2019 — Power Query Editor
### Dataset: Online Retail II (UCI Machine Learning Repository)

---

## What is Power Query?

Power Query is Excel's built-in data transformation tool.
It lets you clean, reshape, and combine data before it loads into Excel.

**Key advantage over manual cleaning:**
- Every step is recorded automatically
- If source data changes, click Refresh — all steps re-run automatically
- No formulas needed — point and click
- Can handle millions of rows without Excel's row limit
- Industry standard for Excel-based data pipelines

**Where it lives:**
```
Excel → Data tab → Get Data
```

---

## What We Did — Step by Step

### Step 1 — Load Raw Data into Power Query
```
Data → Get Data → From File → From Workbook
→ Selected online_retail_II.xlsx
→ Navigator opened showing both tabs
→ Checked "Select multiple items"
→ Checked both: Year 2009-2010 and Year 2010-2011
→ Clicked Transform Data
```

**What happened:** Power Query opened with both tabs as separate queries.

---

### Step 2 — Combine Both Tabs into One
```
Home → Append Queries → Append Queries as New
→ Two tables selected
→ First table: Year 2009-2010
→ Second table: Year 2010-2011
→ Clicked OK
→ New query created: renamed to combined_data
```

**What happened:** Both years merged into one query with 1,067,371 rows.

**Why:** Analyzing both years together gives complete picture. 
Analyzing separately means running every analysis twice.

---

### Step 3 — Fix Data Type: Invoice Column
```
Clicked ABC/123 icon on Invoice column header →
Selected: Text
→ Clicked Replace Current
```

**Why:** Invoice column was loaded as a number type.
Text.StartsWith() function only works on text columns.
Without this fix, the Transaction_Type column showed "Error" for all rows.

**What we learned:** Always check data types before adding calculated columns.
Wrong data type = formula errors downstream.

---

### Step 4 — Add Transaction_Type Column (Flag Cancellations)
```
Add Column → Custom Column
New column name: Transaction_Type
Formula: if Text.StartsWith([Invoice], "C") then "Cancellation" else "Sale"
```

**What it does:** Invoices starting with "C" are cancellations/returns.
Regular sales have numeric invoice numbers like 489434.
Cancellations look like C489434.

**Why:** Cancellations mixed with sales will inflate revenue figures.
We need to separate them before any analysis.

**Result:** Every row tagged as "Sale" or "Cancellation"

---

### Step 5 — Add Customer_Type Column (Flag Guest Purchases)
```
Add Column → Custom Column
New column name: Customer_Type
Formula: if [Customer ID] = null then "Guest" else "Known"
```

**What it does:** ~135K rows have no Customer ID — anonymous/guest purchases.
Tags them as "Guest" so we can include or exclude them in analysis.

**Why:** We don't delete null Customer IDs — they still have valid revenue.
We just can't do customer-level analysis on them.
For revenue analysis: include all rows.
For customer analysis: filter to Known only.

**Result:** Every row tagged as "Guest" or "Known"

---

### Step 6 — Add Quantity_Type Column (Flag Returns)
```
Add Column → Custom Column
New column name: Quantity_Type
Formula: if [Quantity] < 0 then "Return" else "Sale"
```

**What it does:** Negative quantities = returned items.
Tags them so we can separate returns from sales in analysis.

**Why:** Negative quantities reduce revenue.
Need to track return rate separately as a business metric.

**Result:** Every row tagged as "Return" or "Sale"

---

### Step 7 — Add Revenue Column (Derived)
```
Add Column → Custom Column
New column name: Revenue
Formula: [Quantity] * [Price]
```

**What it does:** Calculates revenue per line item.
Price is unit price. Quantity is units sold.
Revenue = how much money this row generated.

**Why:** This is the most important column in the dataset.
The raw data only has unit price — not total revenue.
Every revenue analysis depends on this column.

**Result:** Revenue column showing £ value per transaction line

---

### Step 8 — Trim Whitespace from Description
```
Selected Description column →
Transform tab → Format → Trim
```

**What it does:** Removes leading and trailing spaces from product names.
" WHITE CHERRY LIGHTS " becomes "WHITE CHERRY LIGHTS"

**Why:** Whitespace causes duplicate entries in pivot tables.
"WHITE CHERRY LIGHTS" and " WHITE CHERRY LIGHTS" would appear as two different products.
This inflates product counts and splits revenue incorrectly.

**Result:** All product descriptions cleaned of extra spaces

---

### Step 9 — Add Date Part Columns
```
Selected InvoiceDate column each time before adding:

Month    : Add Column → Date → Month → Month
Year     : Add Column → Date → Year → Year  
Day Name : Add Column → Date → Day → Name of Day
```

**What it does:** Extracts month number, year number, and day name from the date.

**Why:** Can't group by month or year directly from a datetime column in pivot tables.
Need separate columns for time-based analysis.
Month/Year enables monthly trend charts.
Day Name enables day-of-week analysis.

**Result:** Three new columns — Month (1-12), Year (2009/2010/2011), Day (Monday etc)

---

### Step 10 — Remove Negative and Zero Prices
```
Clicked Price column dropdown →
Number Filters → Greater Than → 0 → OK
```

**What it does:** Removes rows where Price ≤ 0.

**Why:** Zero price = free samples, test entries, or data errors.
Negative price = data entry mistake.
These rows have no valid revenue and distort avg price calculations.

**Result:** All rows with Price ≤ 0 removed

---

### Step 11 — Remove Zero Quantities
```
Clicked Quantity column dropdown →
Number Filters → Greater Than → 0 → OK
```

**What it does:** Removes rows where Quantity ≤ 0.

**Why:** After flagging returns in Quantity_Type, we remove them from main dataset.
Zero quantity rows are test/error entries.
Analysis should only include actual completed sales.

**Result:** All rows with Quantity ≤ 0 removed

---

### Step 12 — Load to Excel
```
Home → Close & Load → Close & Load To...
→ Table
→ New Worksheet
→ OK
```

**What happened:** Power Query ran all 11 steps on all 1M+ rows
and loaded the clean result into a new Excel sheet.

---

## Final Column List After Cleaning

| Column | Type | Source | Notes |
|--------|------|--------|-------|
| Invoice | Text | Original | Changed from number to text |
| StockCode | Text | Original | |
| Description | Text | Original | Trimmed whitespace |
| Quantity | Number | Original | Filtered > 0 |
| InvoiceDate | DateTime | Original | |
| Price | Decimal | Original | Filtered > 0 |
| Customer ID | Number | Original | Nulls kept, flagged |
| Country | Text | Original | |
| Transaction_Type | Text | **Derived** | Sale or Cancellation |
| Customer_Type | Text | **Derived** | Guest or Known |
| Quantity_Type | Text | **Derived** | Sale or Return |
| Revenue | Decimal | **Derived** | Quantity × Price |
| Month | Number | **Derived** | 1-12 |
| Year | Number | **Derived** | 2009/2010/2011 |
| Day | Text | **Derived** | Monday-Sunday |

**Original columns: 8 → Final columns: 15**
**Added 7 derived columns**

---

## Problems Found and How We Fixed Them

| Problem | How Found | Fix Applied |
|---------|-----------|-------------|
| Invoice stored as number | Error in Transaction_Type | Changed type to Text |
| Cancellations mixed with sales | Invoice starts with C | Added Transaction_Type flag |
| 135K null Customer IDs | Visual inspection | Added Customer_Type flag |
| Negative quantities (returns) | Knowledge of domain | Added Quantity_Type flag |
| No revenue column | Missing from raw data | Calculated Quantity × Price |
| Whitespace in descriptions | Knowledge of domain | Trim in Transform |
| Zero/negative prices | Knowledge of domain | Filtered out |
| Zero quantities | Knowledge of domain | Filtered out |

---

## Power Query vs Python Cleaning — Comparison

| Task | Python (clean.py) | Power Query |
|------|------------------|-------------|
| Read data | pd.read_csv() | Get Data → From File |
| Filter rows | df[df['col'] > 0] | Number Filters → Greater Than |
| Add column | df['col'] = formula | Add Column → Custom Column |
| Trim text | df['col'].str.strip() | Transform → Format → Trim |
| Change type | df['col'].astype() | Click type icon |
| Record steps | Code in .py file | Applied Steps panel |
| Re-run | python clean.py | Click Refresh |
| Row limit | None | None (Data Model) |

**Same result, different tool.**
Power Query is point-and-click. Python is code.
In industry, analysts use Power Query. Engineers use Python.

---

## Key Power Query Concepts Learned

**Applied Steps:**
Every action is recorded as a step on the right panel.
You can click any step to see the data at that point.
You can delete a step to undo it.
This is like version control for your data cleaning.

**M Language:**
Power Query uses a language called M behind the scenes.
You can see it in the formula bar at the top.
Custom Column formulas use M syntax.
Example: `if Text.StartsWith([Invoice], "C") then "Cancellation" else "Sale"`

**Data Model:**
When data exceeds Excel's 1M row limit, load to Data Model instead.
Data Model has no row limit.
Only accessible through Pivot Tables — not regular Excel formulas.

**Refresh:**
If source data changes, click Data → Refresh All.
All Power Query steps re-run automatically.
Clean data updates without redoing anything manually.

---

## What's Next

```
Phase 4 — Excel Analysis
→ Build pivot tables from cleaned data
→ Revenue by month, country, product
→ Customer analysis
→ Cancellation analysis
→ Charts

Phase 5 — Power BI Dashboard
→ Connect to cleaned Excel file
→ 5 page dashboard
```

---

*Project 3 — Online Retail Sales Analytics*
*Phase 3 — Power Query Cleaning Notes*
