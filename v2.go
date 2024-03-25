package main

import (
	"fmt"
	"os"
	"strconv"

	"github.com/tealeg/xlsx/v3"
)

// v1 performs a series of operations on an Excel file.
// It takes a slice of Operation structs, the input file path, and the output file path as parameters.
// The function opens the input Excel file, performs the specified operations on it, and saves the modified file.
// The supported operations include updating cells, removing columns, inserting columns, hiding columns, hiding sheets, showing columns, and showing sheets.
// The function returns an error if there is any issue with opening or saving the Excel file.
func v2(operations []Operation, inputFile string, outputFile string) {
	fmt.Fprintln(os.Stderr, "using V2...")
	var f *xlsx.File

	// Load the input excel file
	f, err := xlsx.OpenFile(inputFile)
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
				x, y, _ := xlsx.GetCoordsFromCellIDString(key)
				cell, _ := f.Sheet[op.Sheet].Cell(y, x)
				if _, err := strconv.ParseInt(value, 10, 64); err == nil {
					cell.SetNumeric(value)
				} else {
					cell.SetValue(value)
				}
			}

		// - "removeColumn": Removes a column from a sheet.
		case "removeColumn":
			fmt.Printf("removeColumn: %s  ->  ", op.Sheet)
			for i := 0; i < op.Count; i++ {
				fmt.Printf("%d : %s | ", i+1, op.Column)
				colIndex := xlsx.ColLettersToIndex(op.Column)
				hidden := true
				f.Sheet[op.Sheet].Col(colIndex).Hidden = &hidden
			}

		// - "insertColumn": Inserts a column into a sheet.
		case "insertColumn":
			fmt.Printf("insertColumn: %s  ->  ", op.Sheet)
			for i := 0; i < op.Count; i++ {
				fmt.Printf("%d : %s | ", i+1, op.Column)
				newCol := f.Sheet[op.Sheet].Col(xlsx.ColLettersToIndex(op.Column) + i)
				f.Sheet[op.Sheet].Cols.Add(newCol)
				// f.Sheet[op.Sheet].Cols = append(f.Sheet[op.Sheet].Cols, xlsx.Col{})
				// f.InsertCols(op.Sheet, op.Column, 1)
				// f.SetColStyle(op.Sheet, op.Column, colStyle)
			}

		// - "hideColumn": Hides a column in a sheet.
		case "hideColumn":
			fmt.Printf("hideColumn: %s  ->  ", op.Sheet)
			for i := 0; i < op.Count; i++ {
				colIndex := xlsx.ColLettersToIndex(op.Column) + i
				hidden := true
				fmt.Printf("%s | ", xlsx.ColIndexToLetters(colIndex))
				f.Sheet[op.Sheet].Col(colIndex).Hidden = &hidden
			}

		// - "showColumn": Shows a hidden column in a sheet.
		case "showColumn":
			fmt.Printf("showColumn: %s  ->  ", op.Sheet)
			for i := 0; i < op.Count; i++ {
				colIndex := xlsx.ColLettersToIndex(op.Column) + i
				hidden := false
				fmt.Printf("%s | ", xlsx.ColIndexToLetters(colIndex))
				f.Sheet[op.Sheet].Col(colIndex).Hidden = &hidden
			}

		// - "hideSheet": Hides a sheet in the workbook.
		case "hideSheet":
			fmt.Printf("hideSheet: %s", op.Sheet)
			f.Sheet[op.Sheet].Hidden = true

		// - "showSheet": Unhides a sheet in the workboo
		case "showSheet":
			fmt.Printf("showSheet: %s", op.Sheet)
			f.Sheet[op.Sheet].Hidden = false

		// display error message if unknown operation
		default:
			fmt.Printf("Unknown operation type: %s", op.Type)
		}

		fmt.Println()
	}

	// Save the output excel file
	if err := f.Save(outputFile); err != nil {
		panic(err)
	} else {
		fmt.Println("Excel file written to: ", outputFile)
	}
}
