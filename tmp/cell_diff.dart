import 'dart:io';
import 'package:archive/archive.dart';
import 'package:xml/xml.dart';

void main() {
  var origBytes = File('test/test_resources/report-template.xlsm').readAsBytesSync();
  var origArchive = ZipDecoder().decodeBytes(origBytes);
  var expBytes = File('tmp/passthrough_out.xlsm').readAsBytesSync();
  var expArchive = ZipDecoder().decodeBytes(expBytes);

  // Get sheet1.xml from both
  ArchiveFile? origFile, expFile;
  for (var f in origArchive) {
    if (f.name == 'xl/worksheets/sheet1.xml') { f.decompress(); origFile = f; break; }
  }
  for (var f in expArchive) {
    if (f.name == 'xl/worksheets/sheet1.xml') { f.decompress(); expFile = f; break; }
  }

  var origDoc = XmlDocument.parse(String.fromCharCodes(origFile!.content));
  var expDoc = XmlDocument.parse(String.fromCharCodes(expFile!.content));

  var origRows = origDoc.findAllElements('sheetData').first.findElements('row').toList();
  var expRows = expDoc.findAllElements('sheetData').first.findElements('row').toList();

  // Check first few rows for differences
  var totalMissing = 0;
  var missingTypes = <String, int>{};
  
  for (var i = 0; i < origRows.length; i++) {
    var origCells = origRows[i].findElements('c').toList();
    var expCells = expRows[i].findElements('c').toList();
    
    if (origCells.length != expCells.length) {
      var origRefs = origCells.map((c) => c.getAttribute('r')).toSet();
      var expRefs = expCells.map((c) => c.getAttribute('r')).toSet();
      var missing = origRefs.difference(expRefs);
      totalMissing += missing.length;
      
      // Check what types of cells are missing
      for (var cell in origCells) {
        if (missing.contains(cell.getAttribute('r'))) {
          var hasValue = cell.findElements('v').isNotEmpty;
          var hasFormula = cell.findElements('f').isNotEmpty;
          var hasStyle = cell.getAttribute('s') != null;
          var type = cell.getAttribute('t') ?? 'none';
          var key = 'type=$type,val=$hasValue,formula=$hasFormula,style=$hasStyle';
          missingTypes[key] = (missingTypes[key] ?? 0) + 1;
        }
      }
    }
  }
  
  print('Total missing cells: $totalMissing');
  print('Missing cell types:');
  missingTypes.forEach((k, v) => print('  $k: $v'));
  
  // Show first 5 missing cells in detail
  print('\nFirst 5 missing cells:');
  var shown = 0;
  for (var i = 0; i < origRows.length && shown < 5; i++) {
    var origCells = origRows[i].findElements('c').toList();
    var expCells = expRows[i].findElements('c').toList();
    var expRefs = expCells.map((c) => c.getAttribute('r')).toSet();
    
    for (var cell in origCells) {
      if (!expRefs.contains(cell.getAttribute('r')) && shown < 5) {
        print('  Row ${i+1}, ${cell.getAttribute("r")}: ${cell.toXmlString()}');
        shown++;
      }
    }
  }
}
