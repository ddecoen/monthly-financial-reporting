package main

import (
	"encoding/xml"
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Excel XML structures for parsing Excel 2003 XML format
type Workbook struct {
	XMLName    xml.Name    `xml:"Workbook"`
	Worksheets []Worksheet `xml:"Worksheet"`
}

type Worksheet struct {
	Name  string `xml:"Name,attr"`
	Table Table  `xml:"Table"`
}

type Table struct {
	Rows []Row `xml:"Row"`
}

type Row struct {
	Cells []Cell `xml:"Cell"`
	Index string `xml:"Index,attr"`
}

type Cell struct {
	Data      Data   `xml:"Data"`
	Index     string `xml:"Index,attr"`
	MergeDown string `xml:"MergeDown,attr"`
}

type Data struct {
	Type  string `xml:"Type,attr"`
	Value string `xml:",chardata"`
}

func main() {
	// Command line flags
	ytdFile := flag.String("ytd", "", "Path to YTD Excel file")
	quarterlyFile := flag.String("quarterly", "", "Path to Quarterly Excel file")
	monthlyFile := flag.String("monthly", "", "Path to Monthly Excel file")
	outputFile := flag.String("output", "Merged_Income_Statement.xlsx", "Path to output Excel file")
	flag.Parse()

	// Validate inputs
	if *ytdFile == "" || *quarterlyFile == "" || *monthlyFile == "" {
		fmt.Println("Usage: merge_excel -ytd <file> -quarterly <file> -monthly <file> [-output <file>]")
		flag.PrintDefaults()
		os.Exit(1)
	}

	// Create new workbook
	merged := excelize.NewFile()
	defer merged.Close()

	// Delete the default Sheet1
	merged.DeleteSheet("Sheet1")

	// Process each file
	files := []struct {
		path      string
		sheetName string
	}{
		{*ytdFile, "IS - YTD"},
		{*quarterlyFile, "IS - Quarterly"},
		{*monthlyFile, "IS - Monthly"},
	}

	for _, file := range files {
		if err := processFile(file.path, merged, file.sheetName); err != nil {
			log.Fatalf("Error processing %s: %v", file.path, err)
		}
		fmt.Printf("✓ Added %s\n", file.sheetName)
	}

	// Save the merged workbook
	if err := merged.SaveAs(*outputFile); err != nil {
		log.Fatalf("Error saving output file: %v", err)
	}

	fmt.Printf("\n✓ Successfully merged files to: %s\n", *outputFile)
}

// processFile handles both .xlsx and XML-based .xls files
func processFile(sourcePath string, dest *excelize.File, newSheetName string) error {
	// Try to detect file type
	f, err := os.Open(sourcePath)
	if err != nil {
		return fmt.Errorf("failed to open file: %w", err)
	}
	defer f.Close()

	// Read first few bytes to detect format
	buf := make([]byte, 512)
	n, _ := f.Read(buf)
	content := string(buf[:n])

	// Check if it's XML-based Excel
	if strings.Contains(content, "<?xml") && strings.Contains(content, "Workbook") {
		return processXMLExcel(sourcePath, dest, newSheetName)
	}

	// Otherwise try as modern Excel format
	return copySheet(sourcePath, dest, newSheetName)
}

// processXMLExcel handles Excel 2003 XML format (.xls files that are actually XML)
func processXMLExcel(sourcePath string, dest *excelize.File, newSheetName string) error {
	// Read and parse XML file
	xmlFile, err := os.Open(sourcePath)
	if err != nil {
		return fmt.Errorf("failed to open XML file: %w", err)
	}
	defer xmlFile.Close()

	byteValue, err := io.ReadAll(xmlFile)
	if err != nil {
		return fmt.Errorf("failed to read XML file: %w", err)
	}

	var workbook Workbook
	if err := xml.Unmarshal(byteValue, &workbook); err != nil {
		return fmt.Errorf("failed to parse XML: %w", err)
	}

	// Create new sheet in destination
	idx, err := dest.NewSheet(newSheetName)
	if err != nil {
		return fmt.Errorf("failed to create sheet: %w", err)
	}

	// Get first worksheet (usually the only one in NetSuite exports)
	if len(workbook.Worksheets) == 0 {
		return fmt.Errorf("no worksheets found in XML file")
	}

	worksheet := workbook.Worksheets[0]

	// Create styles for formatting
	headerStyle, err := dest.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
			Size: 11,
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#D3D3D3"},
			Pattern: 1,
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		return fmt.Errorf("failed to create header style: %w", err)
	}

	normalStyle, err := dest.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		return fmt.Errorf("failed to create normal style: %w", err)
	}

	// Track max column for width adjustment
	maxCol := 0

	// Write data to new sheet
	currentRow := 1
	for _, row := range worksheet.Table.Rows {
		currentCol := 1
		for _, cell := range row.Cells {
			// Handle cell index if specified (for sparse data)
			if cell.Index != "" {
				// Cell index starts at 1 in XML
				fmt.Sscanf(cell.Index, "%d", &currentCol)
			}

			cellName, err := excelize.CoordinatesToCellName(currentCol, currentRow)
			if err != nil {
				return fmt.Errorf("failed to get cell name: %w", err)
			}

			// Set cell value
			value := strings.TrimSpace(cell.Data.Value)
			if value != "" {
				if err := dest.SetCellValue(newSheetName, cellName, value); err != nil {
					return fmt.Errorf("failed to set cell value: %w", err)
				}

				// Apply style
				if currentRow == 1 {
					dest.SetCellStyle(newSheetName, cellName, cellName, headerStyle)
				} else {
					dest.SetCellStyle(newSheetName, cellName, cellName, normalStyle)
				}
			}

			if currentCol > maxCol {
				maxCol = currentCol
			}

			currentCol++
		}
		currentRow++
	}

	// Auto-adjust column widths
	for col := 1; col <= maxCol; col++ {
		colName, _ := excelize.ColumnNumberToName(col)
		dest.SetColWidth(newSheetName, colName, colName, 15)
	}

	// Freeze the top row
	dest.SetPanes(newSheetName, &excelize.Panes{
		Freeze:      true,
		Split:       false,
		XSplit:      0,
		YSplit:      1,
		TopLeftCell: "A2",
		ActivePane:  "bottomLeft",
	})

	// Set the first sheet as active
	if idx == 1 {
		dest.SetActiveSheet(idx)
	}

	return nil
}

// copySheet copies from modern .xlsx format
func copySheet(sourcePath string, dest *excelize.File, newSheetName string) error {
	// Open source file
	source, err := excelize.OpenFile(sourcePath)
	if err != nil {
		return fmt.Errorf("failed to open source file: %w", err)
	}
	defer source.Close()

	// Get the first sheet name from source
	sourceSheetName := source.GetSheetName(0)
	if sourceSheetName == "" {
		return fmt.Errorf("no sheets found in source file")
	}

	// Create new sheet in destination
	idx, err := dest.NewSheet(newSheetName)
	if err != nil {
		return fmt.Errorf("failed to create new sheet: %w", err)
	}

	// Get all rows from source
	rows, err := source.GetRows(sourceSheetName)
	if err != nil {
		return fmt.Errorf("failed to read rows: %w", err)
	}

	// Create styles for formatting
	headerStyle, err := dest.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
			Size: 11,
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#D3D3D3"},
			Pattern: 1,
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		return fmt.Errorf("failed to create header style: %w", err)
	}

	normalStyle, err := dest.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		return fmt.Errorf("failed to create normal style: %w", err)
	}

	// Track max column for width adjustment
	maxCol := 0

	// Copy data row by row
	for rowIdx, row := range rows {
		for colIdx, cellValue := range row {
			cell, err := excelize.CoordinatesToCellName(colIdx+1, rowIdx+1)
			if err != nil {
				return fmt.Errorf("failed to get cell name: %w", err)
			}
			if err := dest.SetCellValue(newSheetName, cell, cellValue); err != nil {
				return fmt.Errorf("failed to set cell value: %w", err)
			}

			// Apply style
			if rowIdx == 0 {
				dest.SetCellStyle(newSheetName, cell, cell, headerStyle)
			} else {
				dest.SetCellStyle(newSheetName, cell, cell, normalStyle)
			}

			if colIdx+1 > maxCol {
				maxCol = colIdx + 1
			}
		}
	}

	// Auto-adjust column widths
	for col := 1; col <= maxCol; col++ {
		colName, _ := excelize.ColumnNumberToName(col)
		dest.SetColWidth(newSheetName, colName, colName, 15)
	}

	// Freeze the top row
	dest.SetPanes(newSheetName, &excelize.Panes{
		Freeze:      true,
		Split:       false,
		XSplit:      0,
		YSplit:      1,
		TopLeftCell: "A2",
		ActivePane:  "bottomLeft",
	})

	// Set the first sheet as active
	if idx == 1 {
		dest.SetActiveSheet(idx)
	}

	return nil
}
