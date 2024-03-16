package main

import (
	"encoding/json"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"

	"github.com/tealeg/xlsx/v3"
)

func v2(args []string) {
	fmt.Fprintln(os.Stderr, "using V2...")
	var f *xlsx.File

	// If operations was not provided, use some default values
	if len(args) < 2 || args[1] == "" {
		operations := []string{
			`{"type": "updateCells", "sheet": "START", "mappings": {"C06": "Test", "C07": "Test Position", "C08": "GoLang", "C09": "22GO01", "C10": "CSE/AI", "C11": "4", "C12": "2024"}}`,
			// `{"type": "insertColumn", "sheet": "IA", "column": "O", "count": 3}`,
			// `{"type": "removeColumn", "sheet": "IA", "column": "R"}`,
			`{"type": "updateCells", "sheet": "IA", "mappings": {"E08": "CO1", "F08": "CO2", "G08": "CO3", "H08": "CO4", "I08": "CO5", "J08": "CO6", "K08": "CO1", "L08": "CO2", "M08": "CO3", "N08": "CO4", "O08": "CO5"}}`,
			`{"type": "updateCells", "sheet": "IA", "mappings": {"E09": "5", "F09": "5", "G09": "5", "H09": "5", "I09": "5", "J09": "5", "K09": "5", "L09": "5", "M09": "5", "N09": "5", "O09": "5"}}`,
			`{"type": "updateCells", "sheet": "IA", "mappings": {"E10": "3", "F10": "3", "G10": "3", "H10": "3", "I10": "3", "J10": "3", "K10": "3", "L10": "3", "M10": "3", "N10": "3", "O10": "3"}}`,
			`{"type": "updateCells", "sheet": "IA", "mappings": {"E11": "4", "F11": "4", "G11": "4", "H11": "4", "I11": "4", "J11": "4", "K11": "4", "L11": "4", "M11": "4", "N11": "4", "O11": "4"}}`,
			`{"type": "updateCells", "sheet": "IA", "mappings": {"E12": "5", "F12": "5", "G12": "5", "H12": "5", "I12": "5", "J12": "5", "K12": "5", "L12": "5", "M12": "5", "N12": "5", "O12": "5"}}`,
		}
		args = append(args, "["+strings.Join(operations, ",")+"]")
		if len(args) < 3 {
			args = append(args, "input.xlsx")
		}
	}

	// args[1] = `[{\"sheet\":\"START\",\"mappings\":{\"C6\":\"2\",\"C7\":\"2\",\"C8\":\"2\",\"C9\":\"2\",\"C10\":\"2\",\"C11\":\"2\"}},{\"sheet\":\"IA\",\"mappings\":{\"E7\":\"Class Test\",\"E8\":\"2\",\"E9\":\"222\",\"E10\":\"2\",\"E11\":\"2\"}},{\"sheet\":\"LAB\",\"mappings\":{\"E7\":\"Lab Experiment\",\"E8\":\"3\",\"E9\":\"3\",\"E10\":\"3\",\"E11\":\"3\"}}]`

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
		f, err = xlsx.OpenFile(args[2])
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
		f, err = xlsx.OpenBinary(inputBytes)
		if err != nil {
			panic(err)
		}
	}

	m := f.Sheet["IA"]
	f.Sheet["IA2"] = m

	// Perform the operations on the Excel file
	for _, op := range operations {
		switch op.Type {
		case "updateCells":
			fmt.Fprintf(os.Stderr, "\nSheet: %s  ->  ", op.Sheet)
			for key, value := range op.Mappings {
				fmt.Fprintf(os.Stderr, "%s : %s  |  ", key, value)
				x, y, _ := xlsx.GetCoordsFromCellIDString(key)
				cell, _ := f.Sheet[op.Sheet].Cell(y, x)
				if _, err := strconv.ParseInt(value, 10, 64); err == nil {
					cell.SetNumeric(value)
				} else {
					cell.SetValue(value)
				}
			}
		case "removeColumn":
			fmt.Fprintf(os.Stderr, "\nSheet: %s  ->  ", op.Sheet)
			fmt.Fprintf(os.Stderr, "removing Column: %s", op.Column)
			colIndex := xlsx.ColLettersToIndex(op.Column)
			hidden := true
			f.Sheet[op.Sheet].Col(colIndex).Hidden = &hidden
		case "insertColumn":
			fmt.Fprintf(os.Stderr, "\nSheet: %s  ->  ", op.Sheet)
			for i := 0; i < op.Count; i++ {
				fmt.Fprintf(os.Stderr, "Inserting Column %d before col: %s | ", i+1, op.Column)
				newCol := f.Sheet[op.Sheet].Col(xlsx.ColLettersToIndex(op.Column) + i)
				f.Sheet[op.Sheet].Cols.Add(newCol)
				// f.Sheet[op.Sheet].Cols = append(f.Sheet[op.Sheet].Cols, xlsx.Col{})
				// f.InsertCols(op.Sheet, op.Column, 1)
				// f.SetColStyle(op.Sheet, op.Column, colStyle)
			}
		default:
			fmt.Fprintf(os.Stderr, "Unknown operation type: %s\n", op.Type)
		}
	}

	// Write the Excel file to standard output if the current app is the compiled exe else if its the go file then save the output to output.xlsx
	if strings.HasSuffix(os.Args[0], "exe\\excel-writer.exe") {
		if err := f.Save("output.xlsx"); err != nil {
			panic(err)
		}
	} else {
		// if err := f.Write(os.Stdout); err != nil {
		// 	panic(err)
		// }
		f.Save(args[3])
	}
}
