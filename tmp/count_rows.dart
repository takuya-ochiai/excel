import 'dart:io';
import 'package:archive/archive.dart';
import 'package:xml/xml.dart';

void main() {
  // Compare row counts between original and exported
  var files = {
    'original': 'test/test_resources/report-template.xlsm',
    'exported': 'tmp/passthrough_out.xlsm',
  };

  files.forEach((label, path) {
    var bytes = File(path).readAsBytesSync();
    var archive = ZipDecoder().decodeBytes(bytes);
    
    for (var file in archive) {
      if (!file.name.contains('worksheets/sheet')) continue;
      file.decompress();
      var content = String.fromCharCodes(file.content);
      var doc = XmlDocument.parse(content);
      var sheetData = doc.findAllElements('sheetData').first;
      var rows = sheetData.findElements('row').toList();
      var cellCount = 0;
      for (var row in rows) {
        cellCount += row.findElements('c').length;
      }
      print('$label ${file.name}: ${rows.length} rows, $cellCount cells');
    }
  });
}
