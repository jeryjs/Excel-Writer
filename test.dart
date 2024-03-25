import 'dart:convert';
import 'dart:io';

void main() async {
  // The path to the Go program
  var goProgramPath = 'excel-writer.exe';

  // The path to the Excel file
  var inputFilePath = 'input.xlsx';

  // The path to the output file
  var outputFilePath = 'output2.xlsx';

  // The map to send to the Go program
  var map = [
    {"type": "updateCells", "sheet": "START", "mappings": {"C06": "Test", "C07": "Test Position", "C08": "GoLang", "C09": "22GO01", "C10": "CSE/AI", "C11": "4", "C12": "2024"}},
    {"type": "removeColumn", "sheet": "IA", "column": "R", "count": 3},
    {"type": "insertColumn", "sheet": "IA", "column": "O", "count": 3},
    {"type": "hideColumn", "sheet": "CES", "column": "O", "count": 3},
    {"type": "showColumn", "sheet": "SEE", "column": "H", "count": 3},
    {"type": "hideSheet", "sheet": "SEE"},
    {"type": "showSheet", "sheet": "IA"},
  ];

  // encode the map as a JSON string
  var operationJson = jsonEncode(map);

  // Start the Go program as a separate process
  var process = await Process.start(goProgramPath, ["-v", "v2", "-op", operationJson, "-i", inputFilePath, "-o", outputFilePath]);

  // write the stdout and stderr of the process to the console
  stdout.addStream(process.stdout);
  stderr.addStream(process.stderr);
}