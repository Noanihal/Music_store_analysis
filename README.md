# E-Commerce Sales Analysis & Interactive Dashboard (Excel)

An end-to-end sales analysis of ~10,000 e-commerce order lines (2011–2014), built entirely in **Microsoft Excel** using PivotTables, PivotCharts, slicers and a KPI-style dashboard. The project answers a practical business question: **where does the business make money, where does it lose it, and how is it trending year over year?**

![Dashboard preview](images/dashboard-preview.png)

---

## Table of Contents

- [Project Overview](#project-overview)
- [Business Questions Answered](#business-questions-answered)
- [Repository Contents](#repository-contents)
- [Dataset](#dataset)
- [Dashboard Features](#dashboard-features)
- [Headline KPIs](#headline-kpis)
- [Key Insights](#key-insights)
- [Methodology](#methodology)
- [How to Use](#how-to-use)
- [Tools & Skills Demonstrated](#tools--skills-demonstrated)
- [Possible Next Steps](#possible-next-steps)
- [Author](#author)

---

## Project Overview

Raw transactional data is hard to act on. This project turns a 9,994-row sales table into a single-page, interactive dashboard that a manager can filter by **Year, Segment and Region** to see revenue, profit, volume and margin at a glance, along with year-over-year growth for every headline metric.

The work is split into two files:

1. **Analysis workbook**: the cleaned data table plus one PivotTable/PivotChart per analysis, so every number is traceable.
2. **Dashboard workbook**: the finished, presentation-ready dashboard with KPI cards, charts and slicers.

## Business Questions Answered

- How much revenue and profit is the business generating, and at what margin?
- How have Sales, Profit, Quantity, Orders and Profit Margin changed year over year?
- Which months are the strongest, and is there a seasonal pattern?
- Which product **categories** drive sales versus profit?
- Which **sub-categories** sell the most?
- Which **states** generate the most revenue?
- How do results differ across **customer segments** and **regions**?

## Repository Contents

```
├── Ecommerce_Sales_Analysis.xlsx      # Data + PivotTables + PivotCharts (analysis layer)
├── E-Commerce_Sales_Dashboard.xlsx    # Final interactive dashboard (presentation layer)
├── images/
│   └── dashboard-preview.png          # Screenshot of the dashboard
└── README.md
```

| File | What's inside |
|---|---|
| `Ecommerce_Sales_Analysis.xlsx` | `Data` sheet (Excel Table, 9,994 rows × 22 columns) and analysis sheets: Combo chart, Waterfall chart, Pie chart, Map Chart, Top 5, KPI, KPI YOY GROWTH |
| `E-Commerce_Sales_Dashboard.xlsx` | A single visible `Dashboard` sheet with KPI cards, 9 visuals and Year / Segment / Region slicers. The supporting analysis sheets are hidden (right-click any sheet tab → **Unhide** to inspect them) |

## Dataset

A retail sales dataset of **9,994 order lines** covering **January 2011 – December 2014**, across **793 customers**, **1,862 products** and **49 US states**.

| Field group | Columns |
|---|---|
| Order | Row ID, Order ID, Year, Order Date, Ship Date, Ship Mode |
| Customer | Customer ID, Customer Name, Segment (Consumer / Corporate / Home Office) |
| Geography | Country, City, State, Postal Code, Region (East / West / Central / South) |
| Product | Product ID, Category, Sub-Category, Product Name |
| Measures | Sales, Quantity, Discount, Profit |

There are no missing values in any column.

## Dashboard Features

| Visual | Purpose |
|---|---|
| **KPI cards** (Sales, Profit, Quantity, Orders, Profit Margin) | Headline totals, each with a trend line and a year-over-year growth indicator |
| **Combo chart**: Sales & Profit by month | Shows seasonality and how profit tracks revenue |
| **Waterfall chart**: Profit by category | Shows how each category contributes to total profit |
| **Doughnut chart**: Sales share by category | Shows the sales mix |
| **Filled map**: Sales by state | Shows geographic concentration |
| **Bar chart**: Top 5 sub-categories by sales | Shows the biggest revenue drivers |
| **Slicers**: Year, Segment, Region | One click to re-cut every visual on the page |

## Headline KPIs

| Metric | Value (2011–2014) |
|---|---|
| Total Sales | $2,297,201 |
| Total Profit | $286,397 |
| Profit Margin | 12.47% |
| Units Sold | 37,873 |
| Order Lines | 9,994 (5,009 unique orders) |

**Year-over-year performance**

| Year | Sales | Profit | Profit Margin | Quantity | Order Lines |
|---|---:|---:|---:|---:|---:|
| 2011 | $484,247 | $49,544 | 10.2% | 7,581 | 1,993 |
| 2012 | $470,533 | $61,619 | 13.1% | 7,979 | 2,102 |
| 2013 | $608,474 | $81,727 | 13.4% | 9,810 | 2,580 |
| 2014 | $733,947 | $93,508 | 12.7% | 12,503 | 3,319 |

2014 vs 2013: **Sales +20.6%**, **Profit +14.4%**, **Quantity +27.5%**, **Order lines +28.6%**.

## Key Insights

1. **Steady growth, with margin pressure.** Sales grew 20.6% in 2014, but profit grew only 14.4%. Profit margin slipped from 13.4% to 12.7%, so growth is being bought at a slightly lower profitability.
2. **Strong seasonality.** September through December generate about **52% of annual sales**, with November the peak month. February is the weakest month.
3. **Sales share is not profit share.** Furniture is 32% of sales but only ~6% of profit (about a 2.5% margin). Technology is 36% of sales and delivers **~51% of profit**.
4. **A top-selling sub-category loses money.** Tables is a top-5 sub-category by sales (about $207K) yet posts a net **loss of about $17.7K**. Bookcases and Supplies are also loss-making.
5. **Deep discounts destroy margin.** Order lines discounted 30% or more have a combined margin of roughly **−37%**, and about 19% of all order lines are unprofitable.
6. **Geographic concentration.** California alone accounts for about 20% of sales, followed by New York and Texas. By region, the **West** is the most profitable (~14.9% margin) and **Central** the least (~7.9%).

> Insights 3–6 were derived from the underlying `Data` sheet with additional grouping; insights 1–2 and the sales-mix and state views are shown directly on the dashboard.

## Methodology

1. **Prepared the data.** Loaded the raw sales table into an Excel Table (`Table1`) so every PivotTable refreshes from a single structured source.
2. **Built PivotTables.** One per analysis: monthly trend, category profit, category sales share, sales by state, top sub-categories, KPI series and year-over-year summary.
3. **Created a calculated field** for **Profit Margin** (Profit ÷ Sales) so that margin is recomputed correctly at every level of aggregation, rather than averaging row-level percentages.
4. **Calculated YoY growth** for each KPI from the year-level PivotTable.
5. **Built PivotCharts** (combo, waterfall, doughnut, bar, filled map, KPI trend lines) linked to those pivots.
6. **Designed the dashboard.** Arranged KPI cards and charts on one canvas and connected **Year, Segment and Region slicers** to the pivots so the page filters interactively.

## How to Use

1. Download `E-Commerce_Sales_Dashboard.xlsx`.
2. Open it in **Microsoft Excel (desktop)**. Excel 2019 / Microsoft 365 is recommended, since the Waterfall and Filled Map charts are not supported in older versions. The filled map also needs an internet connection.
3. Use the **Year, Segment and Region slicers** to filter the dashboard.
4. To see how a number was produced, open `Ecommerce_Sales_Analysis.xlsx` and go to the matching pivot sheet, or unhide the analysis sheets inside the dashboard file.

## Tools & Skills Demonstrated

- **Microsoft Excel**: PivotTables, PivotCharts, calculated fields, Excel Tables, slicers
- **Advanced chart types**: combo, waterfall, doughnut, filled map
- **Dashboard design**: KPI cards, layout, consistent visual hierarchy, interactivity
- **Data analysis**: trend and seasonality analysis, year-over-year growth, profitability and mix analysis
- **Business storytelling**: turning raw data into findings a stakeholder can act on

## Possible Next Steps

- Add a distinct-count measure (Data Model) so "Orders" reports unique orders (5,009) as well as order lines (9,994).
- Add a **Discount vs. Profit** analysis to quantify the margin impact of discounting by sub-category.
- Add customer-level analysis (repeat purchase rate, top customers, RFM segmentation).
- Rebuild the dashboard in **Power BI** or **Tableau** for web sharing.

## Author

**[Your Name]**
[LinkedIn](https://www.linkedin.com/in/your-profile) · [Email](mailto:you@example.com)

---

*If you found this project useful, feel free to star the repository.*
