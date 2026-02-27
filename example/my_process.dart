import 'dart:io';
import 'package:excel/excel.dart';

void main(List<String> args) {
  if (args.isEmpty) {
    print('Usage: dart example/my_process.dart <path_to_excel_file>');
    print('Example: dart example/my_process.dart test/test_resources/example.xlsx');
    return;
  }

  var filePath = args[0];
  var file = File(filePath);
  if (!file.existsSync()) {
    print('Error: File not found: $filePath');
    return;
  }

  var bytes = file.readAsBytesSync();
  var excel = Excel.decodeBytes(bytes);

  print('=== File: $filePath ===');
  print('Sheets: ${excel.tables.keys.toList()}');
  print('');

  for (var table in excel.tables.keys) {
    var sheet = excel.tables[table]!;
    print('--- Sheet: $table (${sheet.maxRows} rows x ${sheet.maxColumns} cols) ---');
    for (var row in sheet.rows) {
      var values = row.map((cell) => cell?.value?.toString() ?? '').toList();
      print(values.join('\t'));
    }
    print('');
  }
}
