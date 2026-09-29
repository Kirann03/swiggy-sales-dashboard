# VBA & Macro Automation

The workbook is stored as `.xlsm` because it contains an embedded VBA project.

## Modules

### modDataCleaning

Responsible for data cleaning and validation automation.

Primary procedure:

- `CleanAndValidateSwiggyData`

### modDashboard

Responsible for dashboard testing, filtering, refreshing, and reset operations.

Primary procedures include:

- `TestDashboardVBA`
- `ApplyDashboardFilters`
- `RefreshDashboard`
- `ResetDashboard`
- `ApplyPivotFilter`
- `ClearAllFilters`
- `ApplyControlPanelFilters`
- `ResetDashboardFilters`

### modExport

Responsible for exporting the dashboard to PDF.

Primary procedure:

- `ExportDashboardPDF`

## VBA Workflow

```text
Control Panel Change
        ↓
Apply Dashboard Filters
        ↓
Apply Pivot Filters
        ↓
Refresh Dashboard
        ↓
User Reviews Dashboard
        ↓
Export Dashboard PDF
```
