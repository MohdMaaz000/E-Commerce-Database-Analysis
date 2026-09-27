# Pivot Tables & Slicers Setup Guide

## Overview
The `Pivot_Analysis` sheet contains pre-calculated summary tables that mirror
what Pivot Tables would produce. To create **actual interactive Pivot Tables
with Slicers and Timeline**, follow these steps in Excel.

---

## Step 1: Create Pivot Tables from Power_Query_Output

1. Go to the **Power_Query_Output** sheet
2. Select any cell in the data range
3. Go to **Insert → PivotTable**
4. Choose "New Worksheet" or "Existing Worksheet"
5. Click OK

### Recommended Pivot Tables:

#### PT1: Total Sales by Region
- Rows: Region
- Values: Sum of Sales_Amount, Sum of Profit

#### PT2: Sales by Product Category
- Rows: Category
- Values: Sum of Sales_Amount, Sum of Profit, Count of Order_ID

#### PT3: Monthly Sales Trend
- Rows: Month (or Order_Date grouped by Month)
- Values: Sum of Sales_Amount, Sum of Profit

#### PT4: Top 10 Products
- Rows: Product_Name
- Values: Sum of Sales_Amount
- Sort: Descending by Sales_Amount
- Filter: Top 10

#### PT5: Sales by Customer Segment
- Rows: Customer_Segment
- Values: Sum of Sales_Amount, Sum of Profit

#### PT6: Sales Channel Performance
- Rows: Sales_Channel
- Values: Sum of Sales_Amount, Sum of Profit, Count of Order_ID

#### PT7: Monthly Growth
- Rows: Month_Num
- Values: Sum of Sales_Amount
- Add Calculated Field: Growth = (Current - Previous) / Previous

---

## Step 2: Add Slicers

1. Click on any Pivot Table
2. Go to **PivotTable Analyze → Insert Slicer**
3. Select these fields:
   - ☑ Region
   - ☑ Category
   - ☑ Customer_Segment
   - ☑ Sales_Channel
   - ☑ Order_Status
4. Click OK
5. Position slicers on the Dashboard sheet

### Connect Slicers to Multiple Pivot Tables:
1. Right-click on a slicer
2. Select **Report Connections...**
3. Check ALL Pivot Tables that should respond to this slicer
4. Click OK

---

## Step 3: Add Timeline Slicer

1. Click on any Pivot Table
2. Go to **PivotTable Analyze → Insert Timeline**
3. Select **Order_Date**
4. Click OK
5. Set the timeline to show **Months**
6. Position on the Dashboard sheet

### Connect Timeline to Multiple Pivot Tables:
1. Right-click on the Timeline
2. Select **Report Connections...**
3. Check ALL relevant Pivot Tables
4. Click OK

---

## Step 4: Create Pivot Charts

For each Pivot Table:
1. Click on the Pivot Table
2. Go to **PivotTable Analyze → PivotChart**
3. Choose the appropriate chart type:
   - Monthly Trend → Line Chart
   - Sales by Region → Clustered Bar
   - Sales by Category → Horizontal Bar
   - Top 10 Products → Horizontal Bar
   - Profit by Region → Pie Chart
   - Customer Segment → Pie Chart
   - Sales Channel → Clustered Column

4. Move charts to the Dashboard sheet
5. Resize and align professionally

---

## Step 5: Dashboard Layout

Arrange on the Dashboard sheet:
```
Row 1:      Title Bar
Row 2:      Subtitle / Date Range
Row 4-5:    KPI Cards (already created)
Row 7-22:   Monthly Trend (left) | Sales by Region (right)
Row 23-38:  Sales by Category (left) | Top 10 Products (right)
Row 39-54:  Profit by Region (left) | Customer Segment (right)
Row 55-70:  Sales Channel Performance (full width)
```

Place Slicers in a sidebar or above the charts.

---

## Tips

- Use **Slicer Styles** (Slicer tab) to match the dashboard color scheme
- Set Slicer columns to 1 or 2 for compact layout
- Use **number formatting** on Pivot Table values (₹ format)
- Remove gridlines on the Dashboard sheet (View → uncheck Gridlines)
- Set a consistent font (Calibri recommended)
