package main

import (
	"fmt"
	"os"
	"strconv"

	"github.com/xuri/excelize/v2"
)

// v1 performs a series of operations on an Excel file.
// It takes a slice of Operation structs, the input file path, and the output file path as parameters.
// The function opens the input Excel file, performs the specified operations on it, and saves the modified file.
// The supported operations include updating cells, removing columns, inserting columns, hiding columns, hiding sheets, showing columns, and showing sheets.
// The function returns an error if there is any issue with opening or saving the Excel file.
func v1(operations []Operation, inputFile string, outputFile string) {
	fmt.Fprintln(os.Stderr, "using V1...")
	var f *excelize.File

	// Load the input excel file
	f, err := excelize.OpenFile(inputFile)
	if err != nil {
		fmt.Fprintln(os.Stderr, "Error opening Excel file: ", err)
		os.Exit(1)
	} else {
		fmt.Println("Excel file read from: ", inputFile)
	}

	// Perform the operations on the Excel file
	for _, op := range operations {
		switch op.Type {

		// - "updateCells": Updates the values of specific cells in a sheet.
		case "updateCells":
			fmt.Printf("updateCells: %s  ->  ", op.Sheet)
			for key, value := range op.Mappings {
				fmt.Printf("%s : %s  |  ", key, value)
				f.SetCellDefault(op.Sheet, key, value)
			}

		// - "removeColumn": Removes a column from a sheet.
		case "removeColumn":
			fmt.Printf("removeColumn: %s  ->  ", op.Sheet)
			for i := 1; i <= op.Count; i++ {
				fmt.Printf("%d : %s | ", i+1, op.Column)
				f.RemoveCol(op.Sheet, op.Column)
			}

		// - "insertColumn": Inserts a column into a sheet.
		case "insertColumn":
			fmt.Printf("insertColumn: %s  ->  ", op.Sheet)
			colStyle, _ := f.GetColStyle(op.Sheet, op.Column)
			for i := 0; i < op.Count; i++ {
				fmt.Printf("%d : %s | ", i+1, op.Column)
				f.InsertCols(op.Sheet, op.Column, 1)
				f.SetColStyle(op.Sheet, op.Column, colStyle)

				// Copy formatting from old column to new column
				colIndex, _ := excelize.ColumnNameToNumber(op.Column)
				cols, _ := f.GetCols(op.Sheet)
				numCells := len(cols[colIndex+1])
				for i := 1; i <= numCells; i++ {
					nextColName, _ := excelize.ColumnNumberToName(colIndex + 1)
					cellStyle, _ := f.GetCellStyle(op.Sheet, nextColName+strconv.Itoa(i))
					f.SetCellStyle(op.Sheet, op.Column+strconv.Itoa(i), op.Column+strconv.Itoa(i), cellStyle)
				}
			}

		// - "hideColumn": Hides a column in a sheet.
		case "hideColumn":
			fmt.Printf("hideColumn: %s  ->  ", op.Sheet)
			for i := 0; i < op.Count; i++ {
				colIndex, _ := excelize.ColumnNameToNumber(op.Column)
				colName, _ := excelize.ColumnNumberToName(colIndex + i)
				fmt.Printf("%s | ", colName)
				f.SetColVisible(op.Sheet, colName, false)
			}

		// - "showColumn": Shows a hidden column in a sheet.
		case "showColumn":
			fmt.Printf("showColumn: %s -> ", op.Sheet)
			for i := 0; i < op.Count; i++ {
				colIndex, _ := excelize.ColumnNameToNumber(op.Column)
				colName, _ := excelize.ColumnNumberToName(colIndex + i)
				fmt.Printf("%s | ", colName)
				f.SetColVisible(op.Sheet, colName, true)
			}

		// - "hideSheet": Hides a sheet in the workbook.
		case "hideSheet":
			fmt.Printf("hideSheet: %s", op.Sheet)
			f.SetSheetVisible(op.Sheet, false)

		// - "showSheet": Unhides a sheet in the workbook.
		case "showSheet":
			fmt.Printf("showSheet: %s", op.Sheet)
			f.SetSheetVisible(op.Sheet, true)

		// display error message if unknown operation
		default:
			fmt.Printf("Unknown operation type: %s", op.Type)
		}

		fmt.Println()
	}

	// Save the output excel file
	if err := f.SaveAs(outputFile); err != nil {
		panic(err)
	} else {
		fmt.Println("Excel file written to:", outputFile)
	}
}
