# Excel Merger for NetSuite Income Statements

Go tool to merge multiple NetSuite Income Statement exports into a single Excel workbook.

## Features

- Merges three Excel files into one workbook with named sheets
- Handles Excel 2003 XML format (.xls files from NetSuite)
- Handles modern .xlsx format
- Creates properly named sheets: "IS - YTD", "IS - Quarterly", "IS - Monthly"
- **Professional formatting** (matches Q42025-Income_Statement.xlsx):
  - **Title section** (rows 1-4): Bold size 12, centered
    - Company name
    - Report name
    - Date range
  - **Column headers** (row 7): Bold size 7, gray background (#D0D0D0)
    - Left-aligned for first column
    - Right-aligned for amount column
  - **Data rows** (row 9+): Bold size 8
    - Left-aligned for account names
    - Right-aligned for amounts with currency formatting
    - **Number format**: `$#,##0.00` with thousands separators
    - Negative numbers shown in parentheses: `($1,234.56)`
    - Applies to ALL amount columns (single column for YTD, multiple for Quarterly/Monthly)
  - **Column widths**: A=46.25, B=15.25 (optimized for financial statements)
  - **Frozen panes** at row 8 for easy scrolling
  - Board presentation-ready output

## Quick Start

1. Download your Income Statement reports from NetSuite as Excel files

2. Run the merger:
```bash
go run merge_excel.go \
  -ytd "IncomeStatement649.xls" \
  -quarterly "IncomeStatement459.xls" \
  -monthly "IncomeStatement-850.xls" \
  -output "Merged_Income_Statement.xlsx"
```

## Build Executable

```bash
go build -o merge_excel
./merge_excel -ytd file1.xls -quarterly file2.xls -monthly file3.xls
```

## Output

Creates a single `.xlsx` file with three sheets:
- **IS - YTD**: Year-to-date income statement
- **IS - Quarterly**: Quarterly income statement
- **IS - Monthly**: Monthly income statement

## Requirements

- Go 1.21 or later
- github.com/xuri/excelize/v2 (installed via `go mod download`)
