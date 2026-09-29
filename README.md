# Swiggy Sales & Restaurant Analytics Dashboard — Excel + VBA

An interactive, automation-enabled **Microsoft Excel analytics dashboard** built on a Swiggy food-order dataset. The project combines advanced Excel analytics, PivotTables/PivotCharts, data-quality validation, interactive dashboard controls, and VBA automation into a single `.xlsm` workbook.

## Project Overview

The dashboard is designed to turn a large restaurant/order dataset into an interactive reporting solution for exploring:

- Restaurant record volume
- Monthly order/record trends
- Top restaurants
- Category-level performance
- State-level record distribution
- Restaurant and category lookup/analysis
- Data-quality status and validation results
- Interactive dashboard filtering

## Dataset

The workbook contains **197,430 data records** in the raw Swiggy dataset.

The raw dataset includes fields such as:

- State
- City
- Order Date
- Day
- Quarter
- Week
- Restaurant Name
- Location
- Category
- Dish Name
- Food Type
- Price (INR)
- Rating
- Rating Count

The workbook also contains a cleaned data table with additional analytical fields used by the dashboard.

## Key Features

### 1. Interactive Dashboard

The `Dashboard` worksheet provides the main reporting interface with KPI/analysis areas and visualizations connected to PivotTables and supporting calculation sheets.

### 2. PivotTable Analytics

The workbook contains five named PivotTables:

| PivotTable | Purpose |
|---|---|
| `ptMapStateVolume` | State-level record volume / map source |
| `ptTopRestaurants` | Top restaurants by record volume |
| `ptRestaurantVolume` | Restaurant-level volume analysis |
| `ptMonthlyTrend` | Monthly record-volume trend |
| `ptCategoryVolume` | Category-level record volume |

### 3. Dashboard Charts

The workbook includes charts for:

- Monthly Record Volume Trend
- Top 10 Restaurants by Record Volume
- Top 10 Categories by Record Volume
- Record Volume by State

### 4. Control Panel

The `Control_Panel` worksheet provides dashboard controls used to drive filtering and synchronization across the analytical views.

### 5. Data Quality Validation

The `Data_Quality` worksheet provides automated checks for:

- Total records
- Blank state values
- Blank rows
- Missing values
- Invalid prices
- Invalid ratings
- Duplicate rows
- Invalid rating counts
- Overall validation status

The current workbook records **28 duplicate rows** while the other listed validation checks are passing.

### 6. VBA Automation

The workbook is macro-enabled and contains VBA modules for dashboard automation and maintenance, including:

- `modDataCleaning`
  - `CleanAndValidateSwiggyData`
- `modDashboard`
  - `TestDashboardVBA`
  - `ApplyDashboardFilters`
  - `RefreshDashboard`
  - `ResetDashboard`
- `modExport`
  - `ExportDashboardPDF`
- Pivot/filter helper procedures such as:
  - `ApplyPivotFilter`
  - `ClearAllFilters`
  - `ApplyControlPanelFilters`
  - `ResetDashboardFilters`

### 7. PDF Export

A VBA automation routine is included to export the dashboard to PDF using Excel's PDF export functionality.

## Workbook Structure

| Worksheet | Purpose |
|---|---|
| `Dashboard` | Main interactive dashboard |
| `Analysis` | Analytical calculations/views |
| `Restaurant_Lookup` | Restaurant lookup/analysis |
| `Data_Quality` | Data validation and quality summary |
| `Pivot_Data` | Pivot and chart source data |
| `Calculations` | Supporting calculations |
| `Control_Panel` | Interactive dashboard controls |
| `Lists` | Supporting lists / validation values |
| `Clean_Data` | Cleaned analytical dataset |
| `Swiggy Data` | Raw dataset |

## Excel Tables

The workbook contains these structured Excel tables:

- `tblSwiggyRaw`
- `tblSwiggyClean`
- `tblMapSource`

## Tools & Technologies

- Microsoft Excel
- Excel Tables
- PivotTables
- PivotCharts / Excel Charts
- Advanced Excel formulas
- Data Validation
- Power Query / data-preparation workflow
- VBA / Macros
- Dashboard design
- Data Quality Checks

## How to Use

1. Download `Swiggy_Sales_Dashboard.xlsm`.
2. Open it in **Microsoft Excel desktop**.
3. Enable macros when prompted.
4. Start from the `Dashboard` worksheet.
5. Use the `Control_Panel` to interact with dashboard filters.
6. Use `Data_Quality` to review validation results.
7. Run the VBA automation from Excel's macro interface when required.
8. Use the PDF export macro to generate a dashboard report.

> **Important:** Excel for the web may not run VBA macros. Use the desktop version of Microsoft Excel for the full interactive and automation functionality.

## Repository Structure

```text
swiggy-sales-dashboard/
│
├── README.md
├── .gitignore
├── Swiggy_Sales_Dashboard.xlsm
│
└── documentation/
    ├── project-overview.md
    ├── dashboard-features.md
    └── vba-macros.md
```

## Portfolio Highlights

This project demonstrates practical experience with:

- Large-row-count Excel reporting
- Data cleaning and validation
- Interactive business dashboards
- Pivot-based analytics
- KPI/reporting design
- Excel automation using VBA
- Reusable filtering and refresh logic
- Automated PDF reporting
- Structured analytical workbook design

## Future Improvements

- Add a dedicated executive KPI layer
- Add more time-series metrics such as revenue and average rating
- Add automated refresh workflows
- Add additional geographic visualizations
- Add a Power BI version using the same analytical model

## Author

**Kiran Shankar Ghosh**

Data Analytics | Excel | SQL | Power BI | Python | VBA
