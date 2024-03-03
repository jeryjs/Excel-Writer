import 'dart:convert';
import 'dart:io';

void main() async {
  // The path to the Go program
  var goProgramPath = 'excel-writer.exe';

  // The path to the Excel file
  var excelFilePath = 'input.xlsx';

  // The path to the output file
  var outputFilePath = 'output2.xlsx';

  // The map to send to the Go program
  var map = [
			{"type": "updateCells", "sheet": "START", "mappings": {"C06": "Test", "C07": "Test Position", "C08": "GoLang", "C09": "22GO01", "C10": "CSE/AI", "C11": "4", "C12": "2024"}},
			{"type": "insertColumn", "sheet": "IA", "column": "O", "count": 3},
			{"type": "removeColumn", "sheet": "IA", "column": "R"},
			{"type": "updateCells", "sheet": "IA", "mappings": {"E08": "CO1", "F08": "CO2", "G08": "CO3", "H08": "CO4", "I08": "CO5", "J08": "CO6", "K08": "CO1", "L08": "CO2", "M08": "CO3", "N08": "CO4", "O08": "CO5"}},
			{"type": "updateCells", "sheet": "IA", "mappings": {"E09": "5", "F09": "5", "G09": "5", "H09": "5", "I09": "5", "J09": "5", "K09": "5", "L09": "5", "M09": "5", "N09": "5", "O09": "5"}},
			{"type": "updateCells", "sheet": "IA", "mappings": {"E10": "3", "F10": "3", "G10": "3", "H10": "3", "I10": "3", "J10": "3", "K10": "3", "L10": "3", "M10": "3", "N10": "3", "O10": "3"}},
			{"type": "updateCells", "sheet": "IA", "mappings": {"E11": "4", "F11": "4", "G11": "4", "H11": "4", "I11": "4", "J11": "4", "K11": "4", "L11": "4", "M11": "4", "N11": "4", "O11": "4"}},
			{"type": "updateCells", "sheet": "IA", "mappings": {"E12": "5", "F12": "5", "G12": "5", "H12": "5", "I12": "5", "J12": "5", "K12": "5", "L12": "5", "M12": "5", "N12": "5", "O12": "5"}},
  ];

  // encode the map as a JSON string
  var jsonString = jsonEncode(map);

  // Start the Go program as a separate process
  var process = await Process.start(goProgramPath, [jsonString]);

  // Listen to the stderr stream and print each line
  process.stderr.transform(utf8.decoder).transform(LineSplitter()).listen((line) {
      print('excel-writer> $line');
  });

  // Open the Excel file
  var excelFile = File(excelFilePath);
  var excelFileBytes = await excelFile.readAsBytes();

  // Send the Excel file to the Go program via standard input
  process.stdin.add(excelFileBytes);
  await process.stdin.close();

  // Open the output file
  var outputFile = File(outputFilePath);
  var outputSink = outputFile.openWrite();

  // Write the output from the Go program to the output file
  await process.stdout.pipe(outputSink);

  // Close the output file
  await outputSink.close();
}