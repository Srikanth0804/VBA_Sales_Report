# Excel VBA Sales Report Automation

**Project:** Automated Sales Report Generator using Excel VBA  
**Files included:** `Sales_Data.xlsx` (sample data), `vba_module.bas` (VBA code), `VBA_CheatSheet.txt` (quick reference)

## What it does
- Reads sales data from `RawData` sheet (columns: Date, Region, Product, Quantity, UnitPrice).
- Aggregates TotalQuantity and TotalSales by Product.
- Creates a 'Report' sheet with summary and a chart.
- Exports the report as `Sales_Report.pdf` in the same folder (if the workbook is saved).

## How to run locally (quick)
1. Download `Sales_Data.xlsx` and `vba_module.bas` from this repository.
2. Open `Sales_Data.xlsx` in Microsoft Excel.
3. Press **Alt + F11** to open the VBA editor.
4. Insert → Module. Open `vba_module.bas` in a text editor, copy all contents, and paste into the new module.
5. Save the workbook as **Excel Macro-Enabled Workbook (*.xlsm)** (e.g., `Excel-VBA-Sales-Report-Automation.xlsm`).
6. Close and re-open the saved `.xlsm` file (to enable macros if prompted).
7. Press **Alt + F8**, select `GenerateSalesReport`, and click **Run**.
8. Check the `Report` sheet and the generated `Sales_Report.pdf` next to your workbook (if saved).

## How to upload to GitHub (quick)
1. Create a new repository (e.g., `Excel-VBA-Sales-Report-Automation`).
2. Upload `Sales_Data.xlsx`, `vba_module.bas`, and this `README.md` via the GitHub web UI.
3. Commit and copy the repo link — add it to your job application.

## Notes / Tips
- This is a small demo project to demonstrate VBA automation skills.
- If asked in an interview, be prepared to explain:
  - How the macro loops through rows and aggregates values.
  - How to paste/import modules in Excel.
  - Basic VBA objects: Workbook, Worksheet, Range, ChartObject.
