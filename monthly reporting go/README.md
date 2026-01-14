# Excel Merger for NetSuite Income Statements

Go tool to merge multiple NetSuite Income Statement exports into a single Excel workbook.

## Features

- Merges three Excel files into one workbook with named sheets
- Handles Excel 2003 XML format (.xls files from NetSuite)
- Handles modern .xlsx format
- Creates properly named sheets: "IS - YTD", "IS - Quarterly", "IS - Monthly"

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
