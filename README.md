# 📊 Sales Performance Dashboard

An end-to-end **Excel-based Sales Analytics Dashboard** built from raw CSV data using **Power Query**, **Pivot Tables**, **XLOOKUP**, and **interactive visualizations**. This project simulates a real-world business reporting workflow — from messy raw data to a polished executive dashboard.

---

## 📌 Business Problem

**Manual monthly reporting** is a common pain point for small and mid-size businesses:

- Sales data arrives in **separate monthly CSV files** from multiple channels
- Analysts spend hours **copying, cleaning, and consolidating** data in Excel
- **Inconsistent formatting** (region names, dates, missing values) causes errors
- Reports are **static** — stakeholders can't filter or drill down
- Every month, the same manual steps are repeated from scratch

**Result:** Slow turnaround, error-prone reports, and limited analytical depth.

---

## ✅ Solution

This project automates the entire reporting pipeline:

```text
📁 Monthly CSV Files (Raw Data)
        ↓
🔄 Power Query (Import + Combine)
        ↓
🧹 Data Cleaning & Transformation
        ↓
📋 Consolidated Dataset (10,000+ rows)
        ↓
🔗 XLOOKUP + Power Query Merges (Enrichment)
        ↓
📊 Pivot Tables (Analysis)
        ↓
📈 Dashboard (Charts + KPIs + Slicers)
```

**Monthly refresh process:**
1. Drop a new CSV file into the `data/` folder
2. Open the workbook → **Data → Refresh All**
3. Everything updates automatically

---

## 📁 Dataset

### Source Files

| File | Description | Rows |
|------|-------------|------|
| `sales_jan.csv` – `sales_dec.csv` | Monthly transaction data | 10,544 total |
| `customers.csv` | Customer lookup table | 500 |
| `products.csv` | Product lookup table | 51 |

### Sales Data Columns

| Column | Type | Description |
|--------|------|-------------|
| `Order_ID` | Text | Unique order identifier (e.g., ORD-100001) |
| `Order_Date` | Date | Transaction date |
| `Customer_ID` | Text | Customer reference (e.g., C-0001) |
| `Product_ID` | Text | Product reference (e.g., P-1001) |
| `Region` | Text | North / South / East / West |
| `State` | Text | Indian state |
| `Sales_Channel` | Text | Online / Retail Store / Marketplace / Wholesale |
| `Quantity` | Integer | Units ordered |
| `Unit_Price` | Decimal | Price per unit (₹) |
| `Discount` | Decimal | Discount percentage (0–30%) |
| `Sales_Amount` | Decimal | Revenue after discount (₹) |
| `Cost` | Decimal | Cost of goods (₹) |
| `Profit` | Decimal | Sales Amount − Cost (₹) |
| `Payment_Method` | Text | Credit Card / UPI / COD / etc. |
| `Order_Status` | Text | Delivered / Shipped / Cancelled / etc. |

### Data Quality Issues (Intentional)

The raw CSV files contain realistic data quality problems:

- **Inconsistent region names** — `"North"`, `"north"`, `"NORTH"`, `" North"`, `"North "`
- **Missing values** — Blank discounts (~6%), missing Customer_IDs (~2%), missing payment methods (~1.5%)
- **Duplicate rows** — ~50 duplicated transactions across all months
- **Date format inconsistencies** — `YYYY-MM-DD`, `DD-MM-YYYY`, `MM/DD/YYYY` mixed in ~5% of rows
- **Inconsistent category names** — `"Electronics"` vs `"electronics"` vs `"Electrnics"`

All issues are resolved through the Power Query transformation process.

---

## 🔄 Data Pipeline

### Power Query Transformation Steps

1. **Import** — Load all `sales_*.csv` files from the `data/` folder
2. **Combine** — Append all monthly files into one table
3. **Promote Headers** — Set first row as column names
4. **Remove Metadata Columns** — Drop file system columns (path, extension, etc.)
5. **Clean Nulls** — Replace blank discounts with 0, missing IDs with "Unknown"
6. **Remove Duplicates** — Deduplicate on `Order_ID`
7. **Standardize Text** — Trim whitespace + Title Case for Region/Category
8. **Parse Dates** — Convert multiple date formats to proper `Date` type
9. **Set Data Types** — Cast Quantity (Int), Price/Sales/Cost/Profit (Decimal)
10. **Add Calculated Columns** — Month, Month Number, Quarter
11. **Recalculate Revenue** — `Quantity × Unit_Price × (1 − Discount)`
12. **Calculate Profit Margin** — `(Profit ÷ Sales_Amount) × 100`
13. **Output** — Clean, consolidated table ready for analysis

The complete **M code** is provided in [`docs/power_query_m_code.pq`](docs/power_query_m_code.pq).

### XLOOKUP vs Power Query Merge

| Approach | Best For | Used In This Project |
|----------|----------|---------------------|
| **XLOOKUP** | Ad-hoc lookups, small datasets, formula-based workflows | `XLOOKUP_Demo` sheet — demonstrates Product & Customer lookups |
| **Power Query Merge** | Large datasets (10K+ rows), automated refresh pipelines | Main `Power_Query_Output` — used for production data enrichment |

Both approaches are demonstrated to show understanding of when to use each.

---

## 📊 Excel Features Used

| Feature | Purpose |
|---------|---------|
| **Power Query** | Import, clean, transform, and combine 12 CSV files |
| **XLOOKUP** | Look up Product Name, Category, Brand, Customer Segment from reference tables |
| **Pivot Tables** | Summarize sales by Region, Category, Month, Segment, Channel |
| **Pivot Charts** | Visualize Pivot Table summaries |
| **Slicers** | Interactive filters for Region, Category, Segment, Channel, Status |
| **Timeline** | Date-based filtering for Order_Date |
| **Excel Tables** | Structured data references |
| **Conditional Formatting** | Highlight positive/negative month-over-month growth |
| **KPI Calculations** | Total Sales, Profit, Margin, Orders, Units, AOV |
| **Number Formatting** | ₹ currency, percentages, comma-separated numbers |

---

## 📈 Dashboard KPIs

| KPI | Value | Description |
|-----|-------|-------------|
| **Total Sales** | ₹234,651,065 | Sum of all Sales_Amount (after discounts) |
| **Total Profit** | ₹46,241,394 | Sum of all Profit (Sales − Cost) |
| **Profit Margin** | 19.7% | Total Profit ÷ Total Sales × 100 |
| **Total Orders** | 10,494 | Count of unique Order_IDs |
| **Total Units Sold** | 30,530 | Sum of all Quantity |
| **Avg Order Value** | ₹22,360 | Total Sales ÷ Total Orders |

---

## ❓ Business Questions Answered

The dashboard enables stakeholders to answer:

| Question | Answered By |
|----------|-------------|
| What are overall sales doing? | KPI cards + Monthly Trend chart |
| Which region is performing best? | Sales by Region bar chart + Profit by Region pie |
| Which category generates the most revenue? | Sales by Category horizontal bar |
| Which products generate the most profit? | Top 10 Products chart |
| How are sales trending over time? | Monthly Sales & Profit Trend line chart |
| Which customer segment contributes most? | Customer Segment pie chart |
| Which sales channel performs best? | Sales Channel Performance bar chart |

All charts respond to **Slicers** and **Timeline**, enabling drill-down analysis.

---

## 🔁 Refresh Process

### Adding New Monthly Data

When a new month's data is available:

1. **Export** the new month's transactions as a CSV file  
   (e.g., `sales_jan_2026.csv`)
2. **Save** the file into the `data/` folder
3. **Open** `excel/Sales_Performance_Dashboard.xlsx`
4. Go to **Data → Refresh All** (or press `Ctrl + Alt + F5`)
5. Power Query automatically:
   - Detects the new CSV file
   - Imports and appends the data
   - Applies all cleaning transformations
6. **Pivot Tables** update with the new data
7. **Dashboard charts** reflect the latest numbers
8. **KPI cards** recalculate automatically

> **Note:** The CSV file must follow the same column structure as existing files. The filename must start with `sales_` and end with `.csv`.

### Setting Up Power Query (First Time)

If Power Query connections need to be re-established:

1. Open Excel → **Data → Get Data → From File → From Folder**
2. Navigate to the `data/` folder
3. Filter to `sales_*.csv` files only
4. Click **Combine & Transform Data**
5. In Power Query Editor, reference the M code in [`docs/power_query_m_code.pq`](docs/power_query_m_code.pq)
6. Click **Close & Load To...** → select the `Power_Query_Output` sheet

For Pivot Table and Slicer setup, see [`docs/pivot_slicer_setup_guide.md`](docs/pivot_slicer_setup_guide.md).

---

## 📂 Project Structure

```text
Sales-Performance-Dashboard/
│
├── data/                              # Raw CSV data files
│   ├── sales_jan.csv                  #   January transactions (641 rows)
│   ├── sales_feb.csv                  #   February transactions (597 rows)
│   ├── sales_mar.csv                  #   March transactions (686 rows)
│   ├── sales_apr.csv                  #   April transactions (730 rows)
│   ├── sales_may.csv                  #   May transactions (775 rows)
│   ├── sales_jun.csv                  #   June transactions (822 rows)
│   ├── sales_jul.csv                  #   July transactions (866 rows)
│   ├── sales_aug.csv                  #   August transactions (913 rows)
│   ├── sales_sep.csv                  #   September transactions (956 rows)
│   ├── sales_oct.csv                  #   October transactions (1,003 rows)
│   ├── sales_nov.csv                  #   November transactions (1,184 rows)
│   ├── sales_dec.csv                  #   December transactions (1,371 rows)
│   ├── customers.csv                  #   Customer lookup (500 records)
│   └── products.csv                   #   Product lookup (51 products)
│
├── excel/                             # Excel workbook
│   └── Sales_Performance_Dashboard.xlsx
│
├── docs/                              # Documentation
│   ├── power_query_m_code.pq          #   Complete Power Query M code
│   └── pivot_slicer_setup_guide.md    #   Step-by-step Pivot/Slicer guide
│
├── screenshots/                       # Dashboard screenshots (for README)
│
├── generate_project.py                # Data generation script
├── README.md                          # This file
└── .gitignore                         # Git ignore rules
```

---

## 🔍 Key Insights

*Based on the generated FY 2025 dataset (10,494 transactions):*

- **Revenue concentration:** Electronics and Home & Kitchen categories drive the majority of revenue due to higher unit prices, despite Clothing having comparable order volumes
- **Seasonal pattern:** Sales show a clear upward trend through the year, with Q4 (Oct–Dec) significantly outperforming Q1 (Jan–Mar) — consistent with e-commerce seasonality
- **Regional distribution:** All four regions (North, South, East, West) contribute to sales, with distribution influenced by customer population density
- **Channel mix:** Online and Retail Store channels handle the highest transaction volumes; Wholesale has fewer but potentially higher-value orders
- **Profit margins:** Overall profit margin of ~19.7% varies significantly by product category — high-margin categories (Books & Stationery) contrast with lower-margin electronics
- **Customer segments:** Consumer and Corporate segments account for the largest share of revenue across all regions

> **Note:** These insights are derived from the actual generated dataset. Specific values will vary if the dataset is regenerated with a different seed.

---

## ⚡ Performance & Automation

### Before (Manual Workflow)
| Step | Manual Effort |
|------|--------------|
| Open 12 CSV files | Copy-paste into one sheet |
| Clean inconsistencies | Find & Replace, manual review |
| Create summary tables | Manual formulas, prone to errors |
| Build charts | Recreate from scratch monthly |
| **Total** | **Significant time per monthly cycle** |

### After (Automated Workflow)
| Step | Automated Effort |
|------|-----------------|
| Drop CSV in folder | Seconds |
| Refresh All | One click, ~10 seconds |
| Review dashboard | Ready immediately |
| **Total** | **Minutes per monthly cycle** |

> **Note:** Actual time savings depend on the organization's current process, data volume, and analyst experience. The value proposition is eliminating repetitive manual work and reducing error risk — not a specific "X hours saved" claim.

---

## 🚀 Future Improvements

| Enhancement | Description |
|-------------|-------------|
| **Power BI Migration** | Convert the dashboard to Power BI for web publishing and sharing |
| **Automated Email Reports** | Use Power Automate or VBA to email the dashboard on a schedule |
| **Scheduled Refresh** | Set up Windows Task Scheduler to refresh the workbook automatically |
| **Customer Segmentation** | Add RFM analysis (Recency, Frequency, Monetary) for customer scoring |
| **Sales Forecasting** | Implement FORECAST.ETS or ARIMA models for trend prediction |
| **Anomaly Detection** | Flag unusual sales spikes/drops with conditional formatting rules |
| **DAX Measures** | Migrate to Power Pivot for more sophisticated calculated measures |
| **API Integration** | Replace CSV files with live API connections to sales platforms |

---

## 🛠 Technical Requirements

- **Microsoft Excel 365** or **Excel 2021** (for XLOOKUP, Power Query, Timeline support)
- **Python 3.8+** (only needed if regenerating the dataset)
- **openpyxl** Python package (for workbook generation)

---

## 📜 License

This project uses **fictional data only**. No real personal information is included.

Built for portfolio and learning purposes.

---

## 👤 Author

**[Your Name]**  
Data Analyst | Business Analyst  
[LinkedIn](#) | [GitHub](#)
