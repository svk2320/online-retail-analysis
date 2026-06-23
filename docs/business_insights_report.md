# Business Questions & Answers

## Project: Online Retail Sales Analytics
**Dataset:** Online Retail II — UCI Machine Learning Repository  
**Tool:** Microsoft Excel 2019 + Power Query + Power BI  
**Period:** 01 December 2009 — 09 December 2011  
**Prepared:** Phase 5 — Business Insights

---

## Executive Insight Summary

This analysis reveals 5 key business patterns:

- Strong seasonality with Q4 (Oct–Nov) contributing peak revenue
- Highly skewed customer base with B2B dominance (Pareto distribution)
- Significant cancellation issues driven by supply-demand mismatch
- UK market over-reliance (~85% revenue concentration risk)
- Clear opportunity for B2B expansion in high AOV European markets

---

## Methodology

- Data cleaned using Power Query (12 transformation steps)
- Cancellations identified using Invoice prefix 'C'
- Revenue calculated as Quantity × Price
- Guest customers identified where Customer ID is NULL
- Analysis performed in Power BI with DAX measures and filters

---

## Data Limitations

- Customer IDs are missing for ~22.7% of transactions
- Cancellation reasons are not explicitly provided in dataset
- Some non-product entries (e.g., DOTCOM POSTAGE) distort revenue metrics
- No marketing or acquisition channel data available
- Inventory and supply chain constraints are not directly recorded

---

## Overview

This document records 15 business questions derived from the five-page Power BI dashboard, along with data-backed answers. Questions are organized by business theme.

---

## 1. Revenue Analysis

---

### Q1. Which months drove the highest revenue, and what caused the peak?

**Answer:**  
Months 10 and 11 (October and November) recorded the highest revenue at £2.98M and £2.73M respectively. This aligns with pre-Christmas purchasing behaviour typical of a UK-based gifting and homewares retailer. The product mix — decorative items, party supplies, and novelty gifts — strongly supports seasonal demand spikes in Q4. Inventory and logistics planning should account for this 6–8 weeks in advance.

---

### Q2. What is the overall revenue growth trend across the dataset period?

**Answer:**  
The monthly revenue trend shows a clear upward trajectory, rising from approximately £1.08M in the early months to £2.98M at peak. The dataset spans December 2009 to December 2011, covering two full Q4 cycles. Both years show the same seasonal pattern, but the second year's peak is higher, indicating underlying business growth beyond seasonal effects.

---

### Q3. Why is the average order value only £20.13 despite over 1 million orders?

**Answer:**  
The average unit price across all products is approximately £4.08, and most products are low-cost decorative and gift items. Customers typically purchase in small quantities per order, keeping the per-invoice total low. This represents a pricing and bundling opportunity — introducing minimum order thresholds, product bundles, or upsell prompts for complementary items could meaningfully increase average order value without requiring new customer acquisition.

---

## 2. Product Performance

---

### Q4. DOTCOM POSTAGE appears in the top products by revenue at £322K but has only 1,436 quantity entries. What is causing this anomaly?

**Answer:**  
DOTCOM POSTAGE is not a physical product — it is a shipping or postage charge that has been recorded as a product line item in the transaction data. Its high revenue-per-unit (effectively £322,657 ÷ 1,436 = ~£225 per entry) reflects aggregated shipping fees rather than actual product sales. This is a data governance issue. DOTCOM POSTAGE should be excluded from all product analysis to prevent it from distorting top-product rankings and revenue calculations.

---

### Q5. WORLD WAR 2 GLIDERS ranks first in quantity sold (110K units) but does not appear in the top revenue list. Are these products underpriced?

**Answer:**  
With 110,000 units sold and no appearance in the top revenue chart, the unit price is likely below £0.50. This is consistent with a high-volume, low-margin product — possibly a pocket-money or impulse-buy item. Whether this is intentional depends on the business model. If the goal is volume and basket-starter behaviour, the current price may be correct. A small price increase (e.g., from £0.30 to £0.50) across 110K units would generate approximately £22K in additional revenue with minimal demand impact.

---

### Q6. WHITE HANGING HEART T-LIGHT HOLDER ranks near the top in both revenue and cancellations. What is driving this?

**Answer:**  
This product ranks second in quantity sold (97K units) and is a top revenue contributor, yet it also leads all cancellations at 5,600 cancelled invoices. Possible causes include stockout-driven cancellations (demand consistently exceeds available inventory), quality complaints leading to returns, or bulk B2B orders placed and subsequently cancelled. The combination of high sales volume and high cancellation volume is a supply chain signal — this product likely needs dedicated inventory buffer and a returns/cancellation root cause analysis.

---

## 3. Customer Analysis

---

### Q7. Guest customers contribute £3.2M in revenue but cannot be tracked or retained. What is the opportunity here?

**Answer:**  
Guest purchases represent 22.7% of all transactions (241,958 rows) and generate £3.2M in revenue. Since these customers have no Customer ID, they cannot be targeted for re-engagement, loyalty programs, or personalized offers. Converting even 20% of guest revenue to known customers (approximately £640K) would unlock long-term retention value significantly beyond the initial purchase, given that known customers average higher repeat purchase rates. A registration incentive at checkout (e.g., discount on next order) is a low-cost mechanism to drive this conversion.

---

### Q8. Customer 18102 alone generates £610K. How risky is this revenue concentration?

**Answer:**  
A single customer contributing £610K represents approximately 3% of total revenue (£20.97M). When combined with the next four highest-value customers (14646, 14156, 14911, 17450), the top five likely account for 8–10% of total revenue. This level of B2B wholesale concentration is a risk — losing one of these accounts would create an immediate, material revenue gap. These accounts should have dedicated relationship management, retention contracts, and early warning monitoring on order frequency changes.

---

### Q9. How wide is the gap between top customers and the average customer — and what does it tell us about customer segmentation?

**Answer:**  
The top customer spends £168.47K while the average customer spend is £20.13 — a gap of over 8,000x. This extreme skew is consistent with a Pareto distribution, where a small number of high-value B2B wholesale buyers account for a disproportionate share of revenue. This distribution strongly supports a formal customer segmentation strategy: VIP/wholesale accounts (top 5–10%), regular B2B buyers (mid-tier), and one-time or occasional retail customers (long tail). Each segment requires a different engagement and pricing approach.

---

## 4. Cancellation Analysis

---

### Q10. Czech Republic has a 60% cancellation rate — what is likely causing this?

**Answer:**  
A 60% cancellation rate at the country level almost certainly points to one of three causes: a fulfillment or logistics failure making reliable delivery to the Czech Republic impossible; potential payment fraud where orders are placed and immediately cancelled to probe payment systems; or a single wholesale reseller account that bulk-orders and cancels regularly. Given the low total invoice count for Czech Republic, this is likely a single-account issue rather than a systemic country-level problem. The account should be individually reviewed and flagged before further orders are accepted.

---

### Q11. Which single intervention would recover the most lost revenue from the £1.53M in cancellations?

**Answer:**  
WHITE HANGING HEART T-LIGHT HOLDER alone accounts for 5,600 cancelled invoices — the highest of any product. Since it is also a top revenue-generating product, each cancelled invoice represents a meaningful revenue loss. Addressing its supply chain reliability (maintaining safety stock, forecasting demand spikes, securing supplier lead times ahead of Q4) would directly reduce the largest single source of cancellation volume. This one product fix is likely to recover more lost revenue than addressing any individual country-level cancellation issue.

---

### Q12. Do cancellation spikes correlate with peak sales periods?

**Answer:**  
Yes. The monthly cancellation trend shows spikes in months 9 through 11, which is the same window as peak sales. This is a demand-supply mismatch pattern — the business generates its highest order volume during the holiday run-up, but inventory and fulfillment capacity cannot keep pace, leading to cancellations. The fix is not demand reduction but supply-side preparation: building stock earlier in Q3, pre-committing carrier capacity, and flagging high-risk SKUs (like WHITE HANGING HEART) for buffer stock before the peak season begins.

---

## 5. Geography & Time Patterns

---

### Q13. The UK generates 85%+ of total revenue. Is the business dangerously over-reliant on a single market?

**Answer:**  
Yes — the UK contributing £17.9M out of £20.97M total revenue is a significant concentration risk. A UK-specific disruption (economic downturn, regulatory change, or a major logistics failure) would have an outsized impact on the business. However, the data shows that international demand does exist: EIRE (£664K), Netherlands (£554K), Germany (£431K), and France (£357K) all have meaningful revenue bases. A targeted Western European expansion strategy, particularly in markets with high average order values like Netherlands (£2,430 per invoice vs UK's £489), would meaningfully diversify this risk.

---

### Q14. Thursday has the highest invoice count at 8,300. Should the business align its campaigns and operations around mid-week?

**Answer:**  
Thursday and Wednesday together account for the busiest order days, while Sunday records near-zero activity. This pattern is consistent with a B2B customer base making purchasing decisions during the business week — likely procurement buyers placing orders on behalf of their organizations. Marketing campaigns, flash sales, and promotional emails launched on Tuesday or Wednesday morning would likely achieve higher engagement and conversion than weekend pushes. Operationally, order processing and warehouse staffing should peak mid-week to match this demand pattern.

---

### Q15. Netherlands averages £2,430 per invoice against the UK's £489. Is there a B2B expansion opportunity in high-AOV international markets?

**Answer:**  
Yes. Netherlands, EIRE (£1,061 per invoice), and Sweden (£875 per invoice) all show significantly higher average invoice values than the UK. This indicates wholesale or bulk purchasing behaviour in these markets — buyers are placing large consolidated orders rather than individual retail purchases. There is a clear opportunity to formalize a B2B channel for these high-AOV markets with dedicated account management, volume-based pricing tiers, and minimum order quantities. This would increase revenue per customer relationship and improve margin efficiency versus the same effort spent on low-AOV retail customers.

---

## Business Impact & Recommendations

If this analysis were used in production decision-making:

- Introduce customer registration incentives to reduce guest transactions (~22.7% of revenue is untrackable)
- Build inventory buffers for top cancellation-driving SKUs ahead of Q4
- Formalize B2B strategy for Netherlands, EIRE, and Sweden markets based on high average invoice value
- Exclude non-product revenue lines (e.g., DOTCOM POSTAGE) from KPI dashboards
- Shift marketing campaigns to mid-week (Tue–Thu) alignment to match B2B purchasing patterns

---

## Top 5 Critical Insights

1. Q4 seasonality drives up to 30% of annual revenue — inventory and logistics must be prepared 6–8 weeks in advance
2. 22.7% of customers are unidentifiable (guest transactions) — significant lost retention opportunity worth £3.2M
3. Top 5 customers likely contribute ~8–10% of revenue — B2B concentration risk requiring dedicated account management
4. Cancellation rate spikes during peak demand periods — a supply-demand capacity issue, not a pricing issue
5. UK dependency (~85% of revenue) creates geographic concentration risk — Netherlands B2B channel is the highest-priority diversification path

---

## Summary Table

| # | Theme | Question | Key Finding |
|---|-------|----------|-------------|
| 1 | Revenue | Peak months? | Oct–Nov peak; Q4 holiday buying pattern |
| 2 | Revenue | Growth trend? | Clear upward trajectory over 2 years |
| 3 | Revenue | Why low AOV? | Low unit prices + small basket sizes |
| 4 | Product | DOTCOM POSTAGE anomaly? | Shipping fee logged as product — exclude from analysis |
| 5 | Product | GLIDERS underpriced? | High-volume, low-margin; small price lift = meaningful gain |
| 6 | Product | WHITE HEART cancellations? | Demand outpacing supply; needs inventory buffer |
| 7 | Customer | Guest revenue opportunity? | £3.2M untrackable; registration incentive recommended |
| 8 | Customer | Customer 18102 risk? | Single account = 3% of revenue; needs dedicated management |
| 9 | Customer | Spend distribution? | 8,000x gap top-to-average; B2B segmentation required |
| 10 | Cancellation | Czech Republic 60%? | Likely single-account fraud or logistics failure |
| 11 | Cancellation | Biggest recovery lever? | Fix WHITE HEART supply chain — highest cancelled volume |
| 12 | Cancellation | Seasonal cancellations? | Spikes in Sep–Nov; demand exceeds supply at peak |
| 13 | Geography | UK concentration risk? | 85% dependency; Netherlands B2B is diversification path |
| 14 | Time | Mid-week campaign timing? | Thursday peak (8.3K); B2B buying pattern confirmed |
| 15 | Geography | Netherlands B2B opportunity? | £2,430 avg invoice vs UK £489; formalize wholesale channel |

---

*Online Retail Sales Analytics*  
*Phase 5 — Business Insights*  
*Dataset: Online Retail II — UCI Machine Learning Repository*
