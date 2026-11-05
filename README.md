 <div align="center">

# 📊 PR. 2 ANALYZER
## Sales & Customer Profitability Analysis

**Microsoft Excel | Business Analytics | What-If Analysis | Statistical Analysis**

![Project](https://img.shields.io/badge/PROJECT-PR.2%20ANALYZER-1d4ed8?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Status](https://img.shields.io/badge/STATUS-COMPLETED%20ANALYSIS-16a34a?style=for-the-badge)
![Focus](https://img.shields.io/badge/FOCUS-BUSINESS%20ANALYTICS-7c3aed?style=for-the-badge)

**BCA – Final Year**  
**Faculty Guide: Prof. Girish Gondaliya**

</div>

---

## 📌 1. Project Overview

**PR. 2 Analyzer – Sales & Customer Profitability Analysis** is a Microsoft Excel-based business data analysis project designed to analyze customer orders, sales performance, discounts, and profitability.

The project uses a dataset containing **200 order records** and applies Excel formulas, Conditional Formatting, What-If Analysis, Descriptive Statistics, Linear Regression, and Dashboard Visualization.

The main purpose of this project is to transform raw sales data into meaningful business information and present the results through structured worksheets, calculations, statistical analysis, and interactive-style reporting.

## 🎯 2. Project Objectives

- Organize and analyze customer order records.
- Calculate total sales and profit.
- Identify top-performing customers.
- Analyze regional and product category performance.
- Apply Conditional Formatting to highlight important values.
- Perform What-If Analysis using different discount scenarios.
- Calculate Descriptive Statistics for numerical columns.
- Perform Linear Regression between Sales and Profit.
- Create charts and a professional Excel Dashboard.
- Generate meaningful business insights from the dataset.

## 🗂️ 3. Dataset Information

| Dataset Details | Description |
|---|---|
| Dataset Name | `pr 2 data set.xlsx` |
| Total Records | 200 Orders |
| Original Columns | 9 |
| Additional Columns | 3 |
| File Format | Excel (.xlsx) |
| Analysis Tool | Microsoft Excel |

### Original Dataset Columns

| Column Name | Description |
|---|---|
| Customer_ID | Unique customer identification field |
| Customer_Name | Customer name |
| Region | Sales region |
| Product_Category | Product category |
| Sales | Total sales amount |
| Quantity | Number of products sold |
| Discount | Discount applied |
| Order_Date | Date of order |
| Profit | Profit earned from the order |

### Additional Calculated Columns

| Column Name | Purpose |
|---|---|
| Order_Month | Extracts the month from Order_Date |
| Monthly_Total_Sales | Calculates total sales for each month |
| Monthly_Sales_Growth_% | Calculates monthly sales growth percentage |

## 🏗️ 4. Workbook Architecture

The complete Excel workbook is organized into **5 worksheets**.

| No. | Worksheet Name | Description |
|:---:|---|---|
| 01 | Data_Master | Original dataset and calculated columns |
| 02 | Top_Customers | Customer-wise sales and profit analysis |
| 03 | What_If_Analysis | Discount scenario and profit analysis |
| 04 | Statistical_Analysis | Descriptive statistics and regression |
| 05 | Interactive_Dashboard | KPIs, charts, summaries, and business insights |

## 🔍 5. Worksheet-by-Worksheet Analysis

### 01. Data_Master – Data Management

This worksheet contains all 200 original order records along with three additional calculated columns.

**Operations Performed:**

- Organized the original dataset.
- Added Order_Month using date functions.
- Calculated Monthly_Total_Sales.
- Calculated Monthly_Sales_Growth_%.
- Converted the dataset into an Excel Table named `SalesData`.

**Excel Features Used:**

- DATE()
- YEAR()
- MONTH()
- EDATE()
- SUMIFS()
- IF()
- IFERROR()

**Custom Number Format:**

```excel
[Color10]"▲ "0.0%;[Red]"▼ "0.0%;[Black]"▬ "0.0%
```

This format displays upward, downward, and neutral monthly growth indicators.

### 02. Top_Customers – Customer Profitability Analysis

This worksheet analyzes customer performance by grouping records using Customer_ID and Customer_Name.

**Analysis Performed:**

- Total Sales by Customer
- Total Profit by Customer
- Number of Orders
- Total Quantity
- Identification of Top 10 Customers

**Excel Features Used:**

- SUMIF()
- SUMIFS()
- INDEX()
- MATCH()
- MAX()
- Conditional Formatting

**INDEX + MATCH Formula Examples:**

```excel
=INDEX(B5:B175,MATCH(MAX(C5:C175),C5:C175,0))
```

```excel
=INDEX(A5:A175,MATCH(MAX(C5:C175),C5:C175,0))
```

These formulas retrieve the customer name and customer ID associated with the highest aggregated sales value.

**Conditional Formatting:**

The Top 10 Items rule highlights customers with the highest total sales using green formatting.

### 03. What_If_Analysis – Discount Scenario Analysis

This worksheet evaluates how different discount percentages may affect estimated profit.

**Discount Scenarios:**

| Scenario | Discount Rate |
|:---:|:---:|
| 01 | 0% |
| 02 | 5% |
| 03 | 10% |
| 04 | 15% |
| 05 | 20% |
| 06 | 25% |
| 07 | 30% |

**Analysis Includes:**

- Current Total Profit
- Projected Profit
- Profit Difference
- Discount Comparison

**Excel Features:**

- What-If Analysis
- Data Table
- Goal Seek
- Scenario Calculations

**Business Purpose:**

To understand how changes in discount percentages can influence estimated business profitability.

**Model Assumption:** The projected profit calculation adjusts reported profit according to the difference between the existing discount and the selected scenario discount, applied to Sales. It is a simplified estimation and does not account for demand, cost, or product-mix changes.

### 04. Statistical_Analysis – Statistical & Regression Analysis

This worksheet performs statistical analysis on four numerical variables:

- Sales
- Quantity
- Discount
- Profit

**Descriptive Statistics:**

| Statistical Measure | Purpose |
|---|---|
| Mean | Calculates average value |
| Standard Error | Measures estimated sampling precision |
| Median | Identifies the middle value |
| Mode | Identifies the most frequent value |
| Standard Deviation | Measures data variation |
| Sample Variance | Measures data dispersion |
| Kurtosis | Describes distribution shape |
| Skewness | Measures distribution asymmetry |
| Range | Difference between maximum and minimum |
| Minimum | Lowest value |
| Maximum | Highest value |
| Sum | Total of values |
| Count | Number of observations |

#### Linear Regression Analysis

**Analysis:** Relationship between Sales and Profit.

| Variable | Description |
|---|---|
| Independent Variable (X) | Sales |
| Dependent Variable (Y) | Profit |

**Regression Equation:**

```text
Estimated Profit = Intercept + (Sales Coefficient × Sales)
```

**Key Results:**

| Regression Metric | Result |
|---|---:|
| R Square | Approximately 0.5982 |
| Sales Coefficient | Approximately 0.3509 |

**Interpretation:**

The R Square value of approximately 0.5982 indicates that Sales explains around 59.8% of the observed variation in Profit in this linear regression model.

The Sales coefficient represents the estimated relationship between Sales and Profit in the dataset. This relationship does not establish causation.

**Excel Feature Used:**

Data → Data Analysis → Regression

Data → Data Analysis → Descriptive Statistics

### 05. Interactive_Dashboard – Business Performance Dashboard

The Interactive_Dashboard worksheet presents the important results of the complete analysis in one place.

**Key Performance Indicators (KPIs):**

| KPI | Description |
|---|---|
| Total Sales | Overall sales amount |
| Total Profit | Overall profit earned |
| Number of Orders | Total order records |
| Average Discount | Average discount percentage |
| Profit Margin | Profit as a percentage of sales |

**Dashboard Summary Tables:**

- Regional Sales and Profit
- Monthly Sales Trend
- Product Category Sales
- Region vs Product Category Sales
- Region vs Product Category Profit

**Dashboard Charts:**

| Chart | Purpose |
|---|---|
| Bar/Column Chart | Regional Sales and Profit Comparison |
| Line Chart | Monthly Sales Trend |
| Doughnut Chart | Product Category Sales Distribution |

**Dynamic Timestamp Formula:**

```excel
=NOW()
```

Displays the current workbook recalculation date and time.

## 🧮 6. Excel Functions & Tools Used

| Function / Feature | Application |
|---|---|
| SUM() | Calculates total values |
| SUMIF() | Conditional summation |
| SUMIFS() | Multiple-condition summation |
| SUMPRODUCT() | Discount scenario calculations |
| IF() | Logical conditions |
| IFERROR() | Error handling |
| DATE() | Date calculations |
| YEAR() | Extracts year |
| MONTH() | Extracts month |
| EDATE() | Calculates monthly date intervals |
| INDEX() | Retrieves a value |
| MATCH() | Finds a matching position |
| MAX() | Finds the highest value |
| NOW() | Displays date and time |
| Conditional Formatting | Highlights important values |
| What-If Analysis | Compares alternative scenarios |
| Goal Seek | Finds an input for a target result |
| Data Analysis ToolPak | Statistical analysis and regression |

## 📊 7. Business Insights & Analysis

The project focuses on the following business questions:

| No. | Business Question | Analysis |
|:---:|---|---|
| 01 | Which customers generate the highest sales? | Top Customer Analysis |
| 02 | Which region performs better? | Regional Sales Analysis |
| 03 | How does sales performance change monthly? | Monthly Sales Analysis |
| 04 | Which product category contributes more sales? | Category Analysis |
| 05 | How do discount changes affect estimated profit? | What-If Analysis |
| 06 | What is the relationship between Sales and Profit? | Linear Regression |

## 🚀 8. How to Review the Project

Follow these steps to explore the workbook:

1. Open `PR2_Analyzer_Complete_Workbook.xlsx` in Microsoft Excel Desktop.
2. Start with the `Interactive_Dashboard` worksheet.
3. Review the original dataset in `Data_Master`.
4. Explore customer performance in `Top_Customers`.
5. Compare discount scenarios in `What_If_Analysis`.
6. Review statistical calculations in `Statistical_Analysis`.
7. Verify formulas, charts, and analysis outputs.
8. Run native Excel features where required by the assignment.
9. Save the final workbook after verification.

## 🛠️ 9. Software & Technologies

| Tool | Purpose |
|---|---|
| Microsoft Excel | Data analysis and calculations |
| Excel Formulas | Data processing and aggregation |
| Conditional Formatting | Data highlighting |
| What-If Analysis | Scenario modelling |
| Data Analysis ToolPak | Statistical calculations |
| Excel Charts | Data visualization |
| Markdown | Project documentation |
| GitHub | Project storage and presentation |

## ✅ 10. Project Completion Checklist

- [x] Original dataset imported.
- [x] 200 order records organized.
- [x] Original nine columns retained.
- [x] Three calculated columns added.
- [x] Excel Table created.
- [x] Customer-wise analysis prepared.
- [x] Top 10 Conditional Formatting prepared.
- [x] INDEX and MATCH formulas included.
- [x] Discount scenarios from 0% to 30% prepared.
- [x] Descriptive Statistics prepared.
- [x] Linear Regression output included.
- [x] Monthly Sales Growth number format applied.
- [x] Dashboard KPIs prepared.
- [x] Dashboard charts and summaries prepared.
- [ ] Native Excel Data Table verified, if required.
- [ ] Goal Seek demonstrated, if required.
- [ ] Native PivotTables verified, if required.
- [ ] Analysis ToolPak procedures verified in Excel Desktop, if required.
- [ ] Final workbook formulas and charts checked.

> **Note:** Formula-based summaries and native Excel features are not identical. If the instructor requires actual Data Tables, Goal Seek, PivotTables, or Analysis ToolPak demonstrations, those operations should be performed and verified in Microsoft Excel Desktop.

## 🏆 11. Learning Outcomes

Through this project, I gained practical experience in Microsoft Excel and business data analysis.

I learned how to organize raw data, apply formulas, analyze customer profitability, use Conditional Formatting, perform discount scenario analysis, calculate descriptive statistics, and understand linear regression.

I also practiced creating dashboards, visualizing business performance, and presenting analytical results in a structured and professional format.

## 👨‍🎓 12. Academic Details

| Field | Information |
|---|---|
| Student Name | Chand Khimani |
| Academic Program | Bachelor of Computer Applications (BCA) – Final Year |
| Course | BCA Data Analysis |
| Project Number | PR. 2 |
| Project Title | PR. 2 Analyzer – Sales & Customer Profitability Analysis |
| Faculty Guide | Prof. Girish Gondaliya |
| Software | Microsoft Excel |
| Documentation | README.md |

---

<div align="center">

## 🌟 THANK YOU FOR REVIEWING MY PROJECT!

**PR. 2 ANALYZER**

*Sales & Customer Profitability Analysis*

**Microsoft Excel | Business Analytics | Statistical Reporting**

</div>
