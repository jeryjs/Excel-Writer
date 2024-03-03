package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

type Operation struct {
	Type     string            `json:"type"`
	Sheet    string            `json:"sheet"`
	Column   string            `json:"column,omitempty"`
	Count    int               `json:"count,omitempty"`
	Mappings map[string]string `json:"mappings,omitempty"`
}

func main() {
	var f *excelize.File
	args := os.Args

	// If operations was not provided, use some default values
	if len(args) < 2 || args[1] == "" {
		operations := []string{
			`{"type": "updateCells", "sheet": "START", "mappings": {"C06": "Test", "C07": "Test Position", "C08": "GoLang", "C09": "22GO01", "C10": "CSE/AI", "C11": "4", "C12": "2024"}}`,
			`{"type": "insertColumn", "sheet": "IA", "column": "O", "count": 3}`,
			`{"type": "removeColumn", "sheet": "IA", "column": "R"}`,
			`{"type": "updateCells", "sheet": "IA", "mappings": {"E08": "CO1", "F08": "CO2", "G08": "CO3", "H08": "CO4", "I08": "CO5", "J08": "CO6", "K08": "CO1", "L08": "CO2", "M08": "CO3", "N08": "CO4", "O08": "CO5"}}`,
			`{"type": "updateCells", "sheet": "IA", "mappings": {"E09": "5", "F09": "5", "G09": "5", "H09": "5", "I09": "5", "J09": "5", "K09": "5", "L09": "5", "M09": "5", "N09": "5", "O09": "5"}}`,
			`{"type": "updateCells", "sheet": "IA", "mappings": {"E10": "3", "F10": "3", "G10": "3", "H10": "3", "I10": "3", "J10": "3", "K10": "3", "L10": "3", "M10": "3", "N10": "3", "O10": "3"}}`,
			`{"type": "updateCells", "sheet": "IA", "mappings": {"E11": "4", "F11": "4", "G11": "4", "H11": "4", "I11": "4", "J11": "4", "K11": "4", "L11": "4", "M11": "4", "N11": "4", "O11": "4"}}`,
			`{"type": "updateCells", "sheet": "IA", "mappings": {"E12": "5", "F12": "5", "G12": "5", "H12": "5", "I12": "5", "J12": "5", "K12": "5", "L12": "5", "M12": "5", "N12": "5", "O12": "5"}}`,
		}
		args = append(os.Args, "["+strings.Join(operations, ",")+"]")
		if len(args) < 3 {
			args = append(args, "input.xlsx")
		}
	}

	// Unmarshal the JSON input
	var operations []Operation
	err := json.Unmarshal([]byte(args[1]), &operations)
	if err != nil {
		fmt.Fprintln(os.Stderr, "Error unmarshalling JSON:", err)
		os.Exit(1)
	}

	// If input file path was provided use it, else read from standard input
	if len(args) >= 3 && args[2] != "" {
		var err error
		f, err = excelize.OpenFile(args[2])
		if err != nil {
			fmt.Fprintln(os.Stderr, "Error opening Excel file:", err)
			os.Exit(1)
		}
	} else {
		// Read the Excel file from standard input
		inputBytes, err := io.ReadAll(os.Stdin)
		if err != nil {
			panic(err)
		}
		f, err = excelize.OpenReader(bytes.NewReader(inputBytes))
		if err != nil {
			panic(err)
		}
	}

	// Perform the operations on the Excel file
	for _, op := range operations {
		switch op.Type {
		case "updateCells":
			fmt.Fprintf(os.Stderr, "\nSheet: %s  ->  ", op.Sheet)
			for key, value := range op.Mappings {
				fmt.Fprintf(os.Stderr, "%s : %s  |  ", key, value)
				f.SetCellDefault(op.Sheet, key, value)
			}
		case "removeColumn":
			fmt.Fprintf(os.Stderr, "\nSheet: %s  ->  ", op.Sheet)
			fmt.Fprintf(os.Stderr, "removing Column: %s", op.Column)
			if err := f.RemoveCol(op.Sheet, op.Column); err != nil {
				panic(err)
			}
		case "insertColumn":
			fmt.Fprintf(os.Stderr, "\nSheet: %s  ->  ", op.Sheet)
			colStyle, _ := f.GetColStyle(op.Sheet, op.Column)
			for i := 0; i < op.Count; i++ {
				fmt.Fprintf(os.Stderr, "Inserting Column %d before col: %s | ", i+1, op.Column)
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
		default:
			fmt.Fprintf(os.Stderr, "Unknown operation type: %s\n", op.Type)
		}
	}

	// Write the Excel file to standard output if the current app is the compiled exe else if its the go file then save the output to output.xlsx
	if strings.HasSuffix(args[0], "exe\\main.exe") {
		if err := f.SaveAs("output.xlsx"); err != nil {
			panic(err)
		}
	} else {
		if err := f.Write(os.Stdout); err != nil {
			panic(err)
		}
	}
}
