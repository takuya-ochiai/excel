import 'dart:io';
import 'package:archive/archive.dart';
import 'package:xml/xml.dart';

void main() {
  var expBytes = File('tmp/passthrough_out.xlsm').readAsBytesSync();
  var expArchive = ZipDecoder().decodeBytes(expBytes);

  for (var file in expArchive) {
    if (file.name == 'xl/worksheets/sheet1.xml') {
      file.decompress();
      var content = String.fromCharCodes(file.content);
      
      try {
        var doc = XmlDocument.parse(content);
        print('XML parse succeeded');
        
        // Count rows
        var sheetData = doc.findAllElements('sheetData').first;
        var rows = sheetData.findElements('row');
        print('Number of rows: ${rows.length}');
        
        // Check for empty/suspicious content
        var rowList = rows.toList();
        if (rowList.isNotEmpty) {
          print('First row: ${rowList.first.toXmlString().substring(0, 200)}');
          print('Last row: ${rowList.last.toXmlString().substring(0, 200)}');
        }
      } catch (e) {
        print('XML parse FAILED: $e');
        
        // Find the line/col of the error
        var lines = content.split('\n');
        print('Number of lines: ${lines.length}');
        for (var i = 0; i < lines.length && i < 5; i++) {
          var line = lines[i];
          print('Line ${i+1} (${line.length} chars): ${line.substring(0, line.length > 200 ? 200 : line.length)}');
        }
      }
      break;
    }
  }
}
