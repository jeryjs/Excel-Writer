# CO-PO-Attainment-Generator Helper App

This repository contains a helper application for the main [CO-PO-Attainment-Generator](github.com/jeryjs/co_po_attainment_generator) app. It's designed to handle specific operations like updating cells and inserting columns in Excel files, providing essential functionality to the main application.

## Usage

This helper app is meant to be used in conjunction with the main [CO-PO-Generator](github.com/jeryjs/co_po_attainment_generator) app. It takes a JSON string as input, which specifies the operations to perform on the Excel file.

The JSON string (as given by [struct_operation.go](struct_operation.go)) should be in the following format:

```json
{
    "type": "updateCells",
    "sheet": "Sheet1",
    "column": "A",
    "count": 3,
    "mappings": {
        "A1": "New Value"
    }
}
```

Each operation object should have the following properties:

- `type`: The type of operation. Supported types are `updateCells`, `removeColumn`, `insertColumn`, `hideColumn`, `showColumn`, `hideSheet`, and `showSheet`.
- `sheet`: The name of the sheet on which to perform the operation.
- `column`: The name of the column for column-specific operations. This is optional and not needed for `updateCells`, `hideSheet`, and `showSheet` operations.
- `count`: The number of times to perform the operation. This is optional and not needed for `updateCells`, `hideSheet`, and `showSheet` operations.
- `mappings`: A map of cell addresses to values for the `updateCells` operation. This is optional and not needed for other operations.

## Command-Line Flags

The program supports the following command-line flags:

- `-v`: Select the variant (`v1` or `v2`).
  - `v1` uses the [excelize package by xuri](github.com/xuri/excelize)
  - `v2` uses the [xlsx package by tealeg](github.com/tealeg/xlsx)
- `-op`: Operations to perform on the excel sheet in JSON string format.
  - This is the operations object mentioned in the [usage section](#usage).
- `-i`: Input Excel file path.
- `-o`: Output Excel file path.

## Examples

Here are a few examples of how to use the application:

- To perform operations specified in a JSON string:

  ```bat
  go run main.go -v v1 -op '[{"type": "updateCells", "sheet": "START", "mappings": {"C06": "Test", "C07": "Test Position"}}]' -i input.xlsx -o output.xlsx
  ```

- To perform operations specified in a file:

  ```bat
  go run main.go -v v1 -op "$(cat operations.json)" -i input.xlsx -o output.xlsx
  ```

- To use the test operations for debugging:

  ```bat
  go run main.go -v v1 -op test -i input.xlsx -o output.xlsx
  ```
