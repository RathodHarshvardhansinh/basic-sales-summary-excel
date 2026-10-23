# Basic Sales Summary (Task 14 – Data Analytics Track)

**Intern:** Rathod Harshvardhansinh J
**Internship:** Veda Technology – Data Analytics
**Level:** 1 | **Day:** 14
**Tool used:** Microsoft Excel

## Objective
Build a basic business summary by calculating **total sales**, **average sales**, and **transaction count** from a retail sales dataset.

## Deliverables
| Deliverable | Where |
|---|---|
| Summary sheet | `Basic_Sales_Summary.xlsx` → **Summary** tab |
| 3 KPIs | Top of the Summary tab (cells C6:C8) |
| Source data | `Basic_Sales_Summary.xlsx` → **Sales Data** tab |

## KPIs

| # | KPI | Formula | Result |
|---|---|---|---|
| 1 | Total Sales | `=SUM('Sales Data'!I2:I201)` | 93,328.00 |
| 2 | Average Sales per Transaction | `=AVERAGE('Sales Data'!I2:I201)` | 466.64 |
| 3 | Transaction Count | `=COUNTA('Sales Data'!A2:A201)` | 200 |

> **Definition:** A KPI (Key Performance Indicator) is a measurable value showing how well a business is achieving an objective.
> **Average vs total:** Total adds up all sales; average divides that total by the number of transactions, so it shows the typical size of one sale.

## Approach
1. Loaded the sales data into the **Sales Data** sheet (Order ID, Date, Customer, Region, Category, Product, Quantity, Unit Price, Sales).
2. Calculated `Sales = Quantity × Unit Price` with a formula in every row.
3. Created the **Summary** sheet with the 3 KPIs using Excel formulas (`SUM`, `AVERAGE`, `COUNTA`).
4. Added breakdown tables by **Region** and **Category** using `SUMIFS`, `AVERAGEIFS`, and `COUNTIFS`.
5. **Verified manually:** Total ÷ Count = Average (check cell shows `OK`), and region totals add up to the grand total.

## Outcome
The workbook gives a quick business snapshot: about **93.3K in total sales across 200 transactions, averaging 466.64 per transaction**. North is the top region by total sales, and East has the highest average sale.

## Dataset Note
The dataset in this repo is a **sample retail dataset** created for practice (suggested datasets: Superstore, Retail Sales). To use the real Superstore data, paste it into the **Sales Data** sheet with Sales in column I; all summary formulas will update automatically.

## Files
```
├── Basic_Sales_Summary.xlsx
├── README.md
└── (optional) screenshots/
```

## Author
Rathod Harshvardhansinh J – Data Analytics Intern, Veda Technology
