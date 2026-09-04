# Emerald Springs Bottling Ltd — Sales & Order Analysis

Monthly sales and order analysis for an Irish beverage distributor, built in Excel. This project cleans and connects raw operational data, calculates commercial metrics, and turns them into a dashboard and findings that support management decisions.

> Portfolio case study completed through the 10Alytics data analytics training programme. Dataset covers June 2026.

![Dashboard](emerald-dashboard.png)

---

## Project Overview

Emerald Springs Bottling Ltd is an Irish beverage company that produces and distributes bottled water, juices, soft drinks, teas and energy drinks. It sells to restaurants, hotels, retailers, wholesalers, supermarkets and cafés across Ireland, through five sales channels: phone, online, sales reps, distributors and email.

Each month the business records its orders, products and customers, but the raw data arrives uncleaned and split across separate tables. This project takes one month of that activity and turns it into a single connected dataset, a set of commercial metrics, and a dashboard management can act on.

---

## Business Problem

Management could see revenue was healthy but could not see what was driving it. The raw data had three issues:

- Customer, product and order records sat in three separate tables with no link between them
- Text fields were inconsistent (spacing and capitalisation)
- Order dates were stored as serial numbers, not readable dates

Without cleaning and connecting the data, the business could not reliably answer basic questions like which products sell, which customers buy most, or which channels and regions drive revenue.

---

## Project Objectives

- Clean inconsistent text fields and convert serial-number dates to standard dates
- Connect the customers, products and orders tables using ID-based lookups
- Calculate order, sales, cost, discount and profit metrics
- Summarise findings in a dashboard and translate them into business recommendations

---

## Business Questions

1. What do the core metrics reveal about overall business health?
2. What is the daily revenue trend, and which day peaked?
3. Which county drives the highest revenue?
4. Which sales channel generates the most revenue?
5. Which product category contributes the most to revenue?
6. What is the cost and profit per product?
7. How does revenue break down by customer type?

---

## Dataset & Data Description

Three source tables covering June 2026:

| Table | Records | Key fields |
|---|---|---|
| Customers | 100 | Customer_ID, County, Customer_Type, Preferred_Channel |
| Products | 20 | Product_ID, Category, Bottle_Size, Unit_Price, Cost_Price |
| Orders | 250 | Order_ID, Order_Date, Customer_ID, Product_ID, Quantity, Discount, Order_Status |

The three tables connect through `Customer_ID` and `Product_ID`.

---

## Tools & Technologies

- **Microsoft Excel** — data cleaning, analysis, dashboard
- Functions: `TRIM`, `CLEAN`, `PROPER`, date conversion, `VLOOKUP`, `INDEX`/`MATCH`, `SUMIFS`, `COUNTIFS`
- PivotTables and dashboard visuals with slicers

---

## Data Cleaning & Preparation

- Standardised inconsistent text fields with `TRIM`, `CLEAN` and `PROPER`
- Converted order dates from serial-number format to standard short dates
- Connected the three tables with ID-based lookups, so every order line carries the customer's county, type and preferred channel, plus the product's category, bottle size, unit price and cost price
- Built calculated fields at the line level: gross sales, total cost, discount and profit
- Cross-checked every summary figure against a control total before accepting it

---

## KPIs & Calculations

| Metric | Definition |
|---|---|
| Gross Sales | Revenue before cost and discount |
| Total Cost | Cost price × quantity across all orders |
| Total Discount | Value given through discounting |
| Net Sales / Profit | Gross sales − cost − discount |
| Profit Margin | Profit as a percentage of gross sales |

Every metric was segmented across five dimensions: day, county, channel, category and customer type.

---

## Core Metrics (June 2026)

| Metric | Value |
|---|---|
| Total Orders | 250 |
| Total Quantity | 6,954 units |
| Gross Sales | €12,960.50 |
| Total Discount | €250.99 |
| Net Sales / Profit | €5,759.11 |
| Profit Margin | 44.4% |

---

## Key Findings

### Revenue rank is not profit rank
Energy drinks are the smallest category by revenue at 11.3% of sales, yet they contain the single most profitable product in the range. **Energy Boost earns more profit per unit sold than any of the other 19 products**, including products in categories several times its size. The business is organised around revenue, but the profit picture points elsewhere.

### Phone dominates the channel mix
| Channel | Revenue | Share |
|---|---|---|
| Phone | €4,773.30 | 36.8% |
| Online | €2,386.00 | 18.4% |
| Distributor | €2,050.40 | 15.8% |
| Email | €1,961.00 | 15.1% |
| Sales Rep | €1,789.80 | 13.8% |

Phone alone brings in more than Online and Sales Rep combined (€4,773.30 vs €4,175.80).

### Category performance
| Category | Revenue | Share |
|---|---|---|
| Water | €4,496.50 | 34.7% |
| Juice | €3,052.40 | 23.6% |
| Soft Drink | €2,336.20 | 18.0% |
| Tea | €1,610.40 | 12.4% |
| Energy Drink | €1,465.00 | 11.3% |

Water leads on revenue but sits near the bottom on profit per unit. Volume and margin tell different stories.

### Customer type
| Customer Type | Revenue | Share |
|---|---|---|
| Restaurant | €3,121.00 | 24.1% |
| Wholesaler | €2,413.90 | 18.6% |
| Hotel | €2,348.00 | 18.1% |
| Retailer | €2,202.90 | 17.0% |
| Café | €2,113.90 | 16.3% |
| Supermarket | €760.80 | 5.9% |

Restaurants lead. Supermarkets sit far below every other segment.

### Regional and other findings
- **Top county:** Waterford (€1,530), followed by Wexford (€1,474). The top five counties account for roughly half of total revenue.
- **Peak day:** 12 June (€745.80 in gross sales).
- **Bottle size:** 500ml dominates at €5,362.50; 1.5L barely sells (€582). Demand is single-serve.
- **Order status:** 195 completed, 49 pending, 6 cancelled. Nearly 20% of orders sit pending, an operational flag worth investigating separately from revenue.

---

## Business Recommendations

**Invest**
- Prioritise Energy Boost for shelf space and marketing on the strength of its own margin, independent of its category's size. Category-level reporting is hiding the most profitable product.

**Maintain**
- Protect the Phone channel before shifting investment further into digital. It is the single largest revenue channel and demand already sits there.
- Protect the top five counties and the Water category. Reliable revenue that needs holding, not fixing.

**Reconsider**
- Decide deliberately on the Supermarket segment: invest to grow it or accept it as minor. At 5.9% it should not stay a passive gap.
- Investigate the nearly 20% of pending orders as a fulfilment question, separate from the revenue story.

---

## Limitations

- The dataset covers a single month (June 2026), so findings show one month's pattern, not a trend. Repeating the review monthly would confirm whether these patterns hold.
- A small number of order records had missing status or category values and were treated as-is.
- This is a training case study using a supplied dataset, not live commercial data.

---

## Conclusion

Emerald Springs converted 44.4% of gross sales into profit in June, with tight discount control. The headline numbers look healthy, but the value of the analysis is underneath them: revenue rank and profit rank disagree, one channel carries most of the demand, and one customer segment sits far below the rest. Reporting on totals alone would have pointed this business in the wrong direction. Breaking the same data down by margin, channel, region and customer type is what turns a monthly report into a set of decisions.

---

## Files in this repository

| File | Description |
|---|---|
| `emerald-project.xlsx` | Full workbook: cleaned data, analysis and dashboard |
| `emerald-raw-data.xlsx` | Original raw data before cleaning |
| `emerald-dashboard.png` | Dashboard screenshot |
| `emerald-presentation.pptx` | Presentation deck |
| `emerald-presentation.pdf` | Presentation (PDF, viewable in browser) |
