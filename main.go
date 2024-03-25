package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"os"
	"strings"
)

// main is the entry point of the program.
// It parses command line flags, reads the operations from JSON,
// and performs the specified operations on the input Excel file.
// The variant flag determines which version of the operations to use.
// The inputFile flag specifies the path of the input Excel file.
// The outputFile flag specifies the path of the output Excel file.
func main() {
	// define the command line flags
	variant := flag.String("v", "v1", "Select the variant (v1 or v2)")
	operationsJson := flag.String("op", "", "Operations to perform on the excel sheet")
	inputFile := flag.String("i", "input.xlsx", "Input Excel file path")
	outputFile := flag.String("o", "output.xlsx", "Output Excel file path")

	flag.Parse()

	// parse the operations from JSON
	operations := parseOperations(*operationsJson)

	switch *variant {
	case "v1":
		v1(operations, *inputFile, *outputFile)
	case "v2":
		v2(operations, *inputFile, *outputFile)
	default:
		fmt.Println("Invalid variant. Please use 'v1' or 'v2'.")
	}
}

// parseOperations parses the given JSON string and returns a slice of Operation.
// If the string "test" is passed to opJson, it uses a test string for debugging.
// It unmarshals the operations JSON and returns the parsed operations.
func parseOperations(opJson string) []Operation {
	// if 'test' is passed to [-op], use a test string for debugging
	if opJson == "test" {
		opJson = getTestOperations()
	}

	// Unmarshal the operations json
	var operations []Operation
	err := json.Unmarshal([]byte(opJson), &operations)
	if err != nil {
		fmt.Fprintln(os.Stderr, "Error unmarshalling JSON: ", err)
		os.Exit(1)
	}

	return operations
}

// getTestOperations returns a string containing a JSON array of test operations.
// This function is used for debugging and testing purposes.
// The function returns the JSON array as a string.
func getTestOperations() string {
	opStrings := []string{
		`{"type": "updateCells", "sheet": "START", "mappings": {"C06": "Test", "C07": "Test Position", "C08": "GoLang", "C09": "22GO01", "C10": "CSE/AI", "C11": "4", "C12": "2024"}}`,
		`{"type": "removeColumn", "sheet": "IA", "column": "R", "count": 3}`,
		`{"type": "insertColumn", "sheet": "IA", "column": "O", "count": 3}`,
		`{"type": "hideColumn", "sheet": "CES", "column": "O", "count": 3}`,
		`{"type": "showColumn", "sheet": "SEE", "column": "H", "count": 3}`,
		`{"type": "hideSheet", "sheet": "SEE"}`,
		`{"type": "showSheet", "sheet": "IA"}`,
	}
	return "[" + strings.Join(opStrings, ",") + "]"
}
