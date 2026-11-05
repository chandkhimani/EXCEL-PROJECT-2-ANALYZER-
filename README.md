::: {align="center"}
# 📊 PR. 2 ANALYZER --- SALES & CUSTOMER PROFITABILITY ANALYSIS

### ✦ Excel Data Analysis · What-If Modelling · Statistical Analysis · Executive Dashboard ✦

**An Academic Project in Microsoft Excel**

`<br>`{=html}

![Project](https://img.shields.io/badge/PROJECT-PR.%202%20ANALYZER-1d4ed8?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Status](https://img.shields.io/badge/STATUS-ANALYSIS%20WORKBOOK-16a34a?style=for-the-badge)
![Focus](https://img.shields.io/badge/FOCUS-BUSINESS%20ANALYTICS-7c3aed?style=for-the-badge)

`<br>`{=html}

**🎓 BCA --- Final Year**   •   **📚 Data Analysis Project**\
**👨‍🏫 Faculty Guide: Prof. Girish Gondaliya**

`<br>`{=html}

------------------------------------------------------------------------

**📘 Documentation** · **🧹 Data Preparation** · **🧮 Excel Formulas** ·
**📉 Regression** · **📊 Dashboard Reporting**
:::

> \[!IMPORTANT\] **Evaluator quick access:** Open
> `PR2_Analyzer_Complete_Workbook.xlsx` and start with the
> **Interactive_Dashboard** worksheet. Review the source records in
> **Data_Master**, then inspect **Top_Customers**, **What_If_Analysis**,
> and **Statistical_Analysis**. Some requirements that depend on native
> Excel tools (PivotTables, Data Tables, Goal Seek and the Analysis
> ToolPak) must still be completed or verified in Microsoft Excel
> Desktop; see the completion checklist below.

## 🧭 Table of Contents

-   [✨ Project Overview](#-project-overview)
-   [🎯 Project Objectives](#-project-objectives)
-   [🗂️ Workbook Architecture](#️-workbook-architecture)
-   [🔍 Worksheet-by-Worksheet
    Breakdown](#-worksheet-by-worksheet-breakdown)
-   [🧠 Excel Formula Toolkit](#-excel-formula-toolkit)
-   [📊 Statistical Analysis](#-statistical-analysis)
-   [📈 Dashboard Visualizations](#-dashboard-visualizations)
-   [🧪 Practical Demonstrations](#-practical-demonstrations)
-   [🚀 How to Review](#-how-to-review)
-   [✅ Submission Checklist](#-submission-checklist)
-   [🏆 Learning Outcomes](#-learning-outcomes)
-   [👩‍🎓 Academic Details](#-academic-details)

------------------------------------------------------------------------

## ✨ Project Overview

**PR. 2 Analyzer** is a Microsoft Excel business analytics project built
around order-level sales data. It organizes customer, region, product
category, sales, quantity, discount, order date and profit information
into a multi-sheet workbook.

The project demonstrates customer-level sales aggregation, conditional
formatting, discount sensitivity analysis, descriptive statistics,
linear regression, formula-driven monthly sales growth and an executive
dashboard.

### 📌 Project Snapshot

  ---------------------------------------------------------------------------
  📍 Metric                           📊 Project Scope
  ----------------------------------- ---------------------------------------
  🗂️ Total worksheets                 **5**

  🧾 Source order records             **200**

  🧩 Original data fields             **9**

  🧮 Formula/helper fields            **3**

  👥 Customer ID/name combinations    **171**

  📊 Main analysis areas              Customer, discount, statistics,
                                      regional and category performance

  📁 Source dataset                   `pr 2 data set.xlsx`

  📁 Analysis workbook                `PR2_Analyzer_Complete_Workbook.xlsx`
  ---------------------------------------------------------------------------

### 🔄 Analysis Workflow

``` text
      SOURCE DATA             ANALYSIS & MODELLING            REPORTING
           │                           │                         │
           ▼                           ▼                         ▼
     200 Order Records ──► Data_Master + Formula Columns ──► Customer Analysis
                                        │                  ├─ What-If Scenarios
                                        │                  ├─ Regression & Statistics
                                        ▼                  └─ Dashboard Charts
                                  Business Insights
```

------------------------------------------------------------------------

## 🎯 Project Objectives

-   🗃️ **Structure the source data** in a master worksheet while
    retaining the original nine fields.
-   🧮 **Create formula-driven columns** for order month, monthly sales
    and monthly sales growth.
-   🏆 **Identify high-value customers** using customer aggregation and
    Top 10 conditional formatting.
-   🎛️ **Model discount scenarios** and compare estimated profit at
    different discount rates.
-   📉 **Study the relationship between Sales and Profit** using simple
    linear regression.
-   📊 **Summarize numerical variables** using descriptive statistics.
-   📈 **Communicate business performance** with KPIs, charts and
    concise data-storytelling observations.

------------------------------------------------------------------------

## 🗂️ Workbook Architecture

  ------------------------------------------------------------------------------------
               \#              📑 Worksheet                🎯 Purpose
  ---------------------------- --------------------------- ---------------------------
               01              **Data_Master**             Original order records plus
                                                           formula-driven month and
                                                           growth columns; Excel table
                                                           `SalesData`

               02              **Top_Customers**           Customer ID/name
                                                           aggregation, sales and
                                                           profit totals, Top 10
                                                           highlighting, INDEX + MATCH
                                                           lookup

               03              **What_If_Analysis**        Discount scenario table
                                                           from 0% to 30% in 5%
                                                           increments, projected
                                                           profit and native Excel
                                                           setup guidance

               04              **Statistical_Analysis**    Descriptive statistics for
                                                           Sales, Quantity, Discount
                                                           and Profit;
                                                           Profit-versus-Sales
                                                           regression output

               05              **Interactive_Dashboard**   KPI cells,
                                                           regional/monthly/category
                                                           summaries, charts,
                                                           pivot-style cross-tab
                                                           summaries and storytelling
                                                           callout
  ------------------------------------------------------------------------------------

------------------------------------------------------------------------

## 🔍 Worksheet-by-Worksheet Breakdown

### 01 · 🗃️ Data_Master --- Source Records & Formula Columns

Contains the 200 source order records and preserves these original
fields:

`Customer_ID`, `Customer_Name`, `Region`, `Product_Category`, `Sales`,
`Quantity`, `Discount`, `Order_Date`, and `Profit`.

Three additional formula-driven columns are included:

  -----------------------------------------------------------------------
  🧩 Added Column                     💡 Purpose
  ----------------------------------- -----------------------------------
  `Order_Month`                       Extracts the first day of the
                                      order's month for monthly grouping

  `Monthly_Total_Sales`               Calculates total sales for the
                                      corresponding calendar month

  `Monthly_Sales_Growth_%`            Compares the current month's sales
                                      with the previous calendar month's
                                      sales
  -----------------------------------------------------------------------

The source records are formatted as an Excel table named
**`SalesData`**.

**Custom number format for monthly growth:**

``` excel
[Color10]"▲ "0.0%;[Red]"▼ "0.0%;[Black]"▬ "0.0%
```

The first month without a prior-month comparison displays 0% through the
formula's error-handling logic.

### 02 · 🏆 Top_Customers --- Customer Value Analysis

Groups transactions by **Customer_ID + Customer_Name** and calculates:

-   Total Sales
-   Total Profit
-   Number of Orders
-   Total Quantity

The Total Sales column uses the **Top 10 Items** conditional-formatting
rule with a soft green fill and dark green text.

**INDEX + MATCH lookup examples:**

``` excel
=INDEX(B5:B175,MATCH(MAX(C5:C175),C5:C175,0))
=INDEX(A5:A175,MATCH(MAX(C5:C175),C5:C175,0))
=MAX(C5:C175)
```

These patterns return the top aggregated customer name, customer ID and
maximum total sales. Adjust the ending row if the customer aggregation
changes.

> \[!NOTE\] The source dataset contains Customer IDs that appear with
> more than one Customer_Name. For this reason, the aggregation keeps
> the ID/name combination visible rather than silently assuming that
> every ID maps to one unique name.

### 03 · 🎛️ What_If_Analysis --- Discount Sensitivity

Compares projected profit at discount rates from **0% to 30% in 5%
increments**.

  📌 Scenario     🎯 Discount Rate
  ------------- ------------------
  01                            0%
  02                            5%
  03                           10%
  04                           15%
  05                           20%
  06                           25%
  07                           30%

The worksheet shows current total profit, projected profit and change
versus current profit.

**Scenario model assumption:** the model estimates the effect of
changing discount by adjusting reported profit by the difference between
the current discount amount and the scenario discount amount, applied to
Sales. This is a simplified sensitivity assumption; it does not model
changes in customer demand, costs, product mix or order volume.

**Native Excel Data Table:** The worksheet includes instructions for
setting up a one-variable Data Table through **Data → What-If Analysis →
Data Table**.

**Goal Seek:** The worksheet also documents how to use Goal Seek to find
a discount input that reaches a chosen profit target.

> \[!IMPORTANT\] The scenario comparison is formula-driven. The native
> Excel Data Table and Goal Seek procedures must be run in Excel Desktop
> if the evaluator specifically requires those Excel features to be
> demonstrated.

### 04 · 📉 Statistical_Analysis --- Descriptive Statistics & Regression

Descriptive statistics are provided for four numerical fields:

-   `Sales`
-   `Quantity`
-   `Discount`
-   `Profit`

The summary includes **Mean, Standard Error, Median, Mode, Standard
Deviation, Sample Variance, Kurtosis, Skewness, Range, Minimum, Maximum,
Sum and Count**.

#### Linear Regression --- Profit vs Sales

-   **Dependent variable (Y):** Profit
-   **Independent variable (X):** Sales

The workbook reports Multiple R, R Square, Adjusted R Square, Intercept,
Sales coefficient, observations and the Sales slope p-value.

**Regression equation:**

``` text
Estimated Profit = Intercept + (Sales Coefficient × Sales)
```

**Dataset result:** R² is approximately **0.5982** and the Sales
coefficient is approximately **0.3509**. In this simple model, Sales
explains about 59.8% of the observed variation in Profit. The
coefficient describes an association in this dataset; it does not
establish causation.

**Native Analysis ToolPak procedure:** Enable the Analysis ToolPak in
Excel, then select **Data → Data Analysis → Regression**. Set Profit as
the Y input range and Sales as the X input range. The workbook also
provides the input-range guidance.

> \[!NOTE\] The workbook includes calculated regression and
> descriptive-statistics outputs. If the assignment requires screenshots
> or a live demonstration of the Excel Data Analysis ToolPak dialog, run
> the corresponding ToolPak procedures in Excel Desktop.

### 05 · 📊 Interactive_Dashboard --- Executive Summary

The dashboard worksheet contains:

**KPI headings** - Total Sales - Total Profit - Number of Orders -
Average Discount - Profit Margin

**Summary sections** - Total Sales and Total Profit by Region - Monthly
Sales Trend - Sales Distribution by Product Category - Region × Product
Category sales summary - Region × Product Category profit summary

**Charts** - Bar/Column Chart --- Total Sales and Profit by Region -
Line Chart --- Monthly Sales Trend - Doughnut Chart --- Sales
Distribution by Product Category

**Dynamic timestamp:**

``` excel
=NOW()
```

The timestamp is formatted as:

``` text
yyyy-mm-dd hh:mm:ss
```

It updates when Excel recalculates the workbook; it is not a permanent
audit timestamp.

**Data storytelling callout:** The dashboard highlights the leading
sales region, the category with the highest profit margin and the
observed Discount-versus-Profit correlation.

> \[!NOTE\] The region/category matrices in this workbook are
> pivot-style summary tables. They are not native Excel PivotTables and
> do not automatically provide PivotTable slicers or refresh behavior.
> Create native PivotTables in Excel Desktop if those are required for
> assessment.

------------------------------------------------------------------------

## 🧠 Excel Formula Toolkit

  -----------------------------------------------------------------------
  🧩 Function / Feature               ⚙️ Application
  ----------------------------------- -----------------------------------
  `SUM`                               Calculate total sales and profit

  `SUMIF` / `SUMIFS`                  Aggregate values by region,
                                      category or date range

  `SUMPRODUCT`                        Estimate profit impact across sales
                                      and discount values

  `IF` / `IFERROR`                    Apply conditions and handle missing
                                      prior-month comparisons

  `DATE` / `YEAR` / `MONTH`           Derive the calendar month from
                                      Order_Date

  `EDATE`                             Set month boundaries for monthly
                                      calculations

  `INDEX` + `MATCH`                   Retrieve top customer information

  `MAX`                               Identify the highest aggregated
                                      customer sales

  `NOW`                               Display the analysis recalculation
                                      timestamp

  Conditional Formatting              Highlight the Top 10 customer sales
                                      values

  What-If Analysis                    Explore discount sensitivity and
                                      target-profit scenarios

  Data Analysis ToolPak               Run regression and descriptive
                                      statistics in Excel
  -----------------------------------------------------------------------

### 💻 Selected Formula Patterns

``` excel
=SUM(SalesData[Sales])

=SUMIFS(SalesData[Sales],SalesData[Order_Month],J4)

=INDEX(Customer_Name_Range,MATCH(MAX(Total_Sales_Range),Total_Sales_Range,0))

=NOW()
```

*These are representative formula patterns. The workbook's actual
formulas should be inspected in the Excel Formula Bar.*

------------------------------------------------------------------------

## 📊 Statistical Analysis

  -----------------------------------------------------------------------
  📌 Measure                          📊 Purpose
  ----------------------------------- -----------------------------------
  Mean / Median / Mode                Describe central tendency

  Standard Error / Standard Deviation Describe sampling precision and
                                      spread

  Sample Variance                     Measure dispersion

  Kurtosis / Skewness                 Describe distribution shape

  Range / Minimum / Maximum           Show the spread and endpoints

  Sum / Count                         Show total value and number of
                                      observations

  R Square / Adjusted R Square        Describe regression fit

  Intercept / Sales Coefficient       Define the fitted linear
                                      relationship
  -----------------------------------------------------------------------

------------------------------------------------------------------------

## 📈 Dashboard Visualizations

  -----------------------------------------------------------------------
  📊 Visualization                    📌 Business Question
  ----------------------------------- -----------------------------------
  **01 · Regional Sales & Profit**    Which region contributes the most
                                      sales and profit?

  **02 · Monthly Sales Trend**        How do sales change across the
                                      available order months?

  **03 · Product Category Sales Mix** What percentage of sales comes from
                                      each product category?

  **04 · Region × Category Summary**  How are sales and profit
                                      distributed across regions and
                                      product categories?
  -----------------------------------------------------------------------

------------------------------------------------------------------------

## 🧪 Practical Demonstrations

  -------------------------------------------------------------------------
  🎯 Demo                 🖱️ Action                 👀 Expected Observation
  ----------------------- ------------------------- -----------------------
  Top customer analysis   Open `Top_Customers`      Review aggregated sales
                                                    and highlighted Top 10
                                                    values

  Customer lookup         Inspect the INDEX + MATCH Top customer name and
                          cells                     ID are returned from
                                                    the aggregation

  Discount sensitivity    Open `What_If_Analysis`   Compare projected
                                                    profit across the seven
                                                    discount rates

  Regression              Open                      Review R², adjusted R²,
                          `Statistical_Analysis`    intercept and Sales
                                                    coefficient

  Monthly trend           Open                      Compare sales totals
                          `Interactive_Dashboard`   across calendar months

  Timestamp               Inspect cell B2 on the    See the latest workbook
                          dashboard                 recalculation time

  Data refresh            Use **Data → Refresh      Refresh the PivotTables
                          All** after native        connected to the source
                          PivotTables are created   table
  -------------------------------------------------------------------------

> \[!TIP\] For a viva demonstration, begin with the dashboard, explain
> one regional result, show the Top 10 customer rule, compare two
> discount scenarios, and finish by explaining what the regression R²
> means.

------------------------------------------------------------------------

## 🚀 How to Review

1.  📂 Open `PR2_Analyzer_Complete_Workbook.xlsx` in **Microsoft Excel
    Desktop**.
2.  🏠 Start with **Interactive_Dashboard** to review the KPIs, charts
    and data-story callout.
3.  🗃️ Open **Data_Master** and inspect the original fields and three
    calculated columns.
4.  🏆 Open **Top_Customers** and review the customer aggregation and
    Top 10 formatting.
5.  🎛️ Open **What_If_Analysis** and compare the discount scenarios.
6.  📉 Open **Statistical_Analysis** and review the descriptive
    statistics and regression output.
7.  🧪 Complete and verify the native Excel features listed in the
    submission checklist.
8.  💾 Save the final workbook after verifying formulas, charts and any
    native Excel analysis features.

**💻 Compatibility:** Microsoft Excel Desktop is recommended for the
native Data Table, Goal Seek, PivotTables and Analysis ToolPak
demonstrations. Availability of the `GROUPBY` function depends on the
Excel version.

------------------------------------------------------------------------

## ✅ Submission Checklist

-   [x] Five required worksheets created
-   [x] Original 200 order records included in Data_Master
-   [x] Three formula-driven helper columns added
-   [x] Customer aggregation and Top 10 conditional-formatting rule
    prepared
-   [x] INDEX + MATCH customer lookup formulas included
-   [x] Discount scenario comparison from 0% to 30% prepared
-   [x] Descriptive statistics for four numerical variables included
-   [x] Profit-versus-Sales regression output included
-   [x] Monthly sales growth arrow number format included
-   [x] Dashboard timestamp formula included
-   [x] Regional, monthly and category charts included
-   [x] Pivot-style region/category summary tables included
-   [x] Three data-storytelling observations included
-   [ ] Run and verify a native Excel What-If Data Table
-   [ ] Run and verify Goal Seek in Excel Desktop
-   [ ] Create and verify native PivotTables with Grand Totals and
    required subtotals
-   [ ] Run the Analysis ToolPak Regression and Descriptive Statistics
    procedures in Excel Desktop if required by the evaluator
-   [ ] Test formulas and charts after opening the workbook in Excel
    Desktop
-   [ ] Confirm student details and final filenames before submission
-   [ ] Submit the Excel workbook and `README.md` together

------------------------------------------------------------------------

## 🏆 Learning Outcomes

Through this project, I practiced organizing order-level business data,
building formula-driven analysis columns, aggregating customer
performance, applying conditional formatting and modelling discount
scenarios.

I also practiced descriptive statistics, simple linear regression,
monthly trend analysis, dashboard chart design and communicating
business observations in a structured Excel report.

------------------------------------------------------------------------

## 👩‍🎓 Academic Details

  -----------------------------------------------------------------------
  🏷️ Field                            📋 Information
  ----------------------------------- -----------------------------------
  **Student Name**                    Chand Khimani

  **Academic Program**                Bachelor of Computer Applications
                                      (BCA) --- Final Year

  **Course Focus**                    BCA Data Analysis

  **Project Title**                   PR. 2 Analyzer --- Sales & Customer
                                      Profitability Analysis

  **Faculty Guide**                   Prof. Girish Gondaliya

  **Tools**                           Microsoft Excel · Excel Formulas ·
                                      Markdown

  **Deliverables**                    Excel Workbook · README
                                      Documentation
  -----------------------------------------------------------------------

------------------------------------------------------------------------

::: {align="center"}
### 🌟 Thank You for Reviewing My Project!

**PR. 2 Analyzer --- Sales & Customer Profitability Analysis**\
*Microsoft Excel · Business Analytics · Statistical Reporting*

[⬆️ Back to
Top](#-pr-2-analyzer--sales--customer-profitability-analysis)
:::
