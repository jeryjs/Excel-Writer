## CO-PO-Attainment-Generator Helper App

This repository contains a helper application for the main CO-PO-Attainment-Generator app. It's designed to handle specific operations like updating cells and inserting columns in Excel files, providing essential functionality to the main application.

### Usage

This helper app is meant to be used in conjunction with the main CO-PO-Attainment-Generator app. It takes a JSON string as input, which specifies the operations to perform on the Excel file.

The JSON string should be in the following format:

```json
{
    "type": "updateCells",
    "sheet": "Sheet1",
    "mappings": {
        "A1": "New Value"
    }
}
```
