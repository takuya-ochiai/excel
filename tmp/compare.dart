import 'dart:io';
import 'package:archive/archive.dart';

void main() {
  // Original
  var origBytes = File('test/test_resources/report-template.xlsm').readAsBytesSync();
  var origArchive = ZipDecoder().decodeBytes(origBytes);

  // Exported
  var expBytes = File('tmp/passthrough_out.xlsm').readAsBytesSync();
  var expArchive = ZipDecoder().decodeBytes(expBytes);

  // Compare file lists
  var origNames = origArchive.map((f) => f.name).toSet();
  var expNames = expArchive.map((f) => f.name).toSet();
  
  var onlyInOrig = origNames.difference(expNames);
  var onlyInExp = expNames.difference(origNames);
  
  if (onlyInOrig.isNotEmpty) print('Only in original: $onlyInOrig');
  if (onlyInExp.isNotEmpty) print('Only in exported: $onlyInExp');

  // Compare sizes of worksheet files
  for (var origFile in origArchive) {
    if (!origFile.name.contains('worksheets/sheet')) continue;
    origFile.decompress();
    var origContent = String.fromCharCodes(origFile.content);
    
    for (var expFile in expArchive) {
      if (expFile.name == origFile.name) {
        expFile.decompress();
        var expContent = String.fromCharCodes(expFile.content);
        print('${origFile.name}: orig=${origContent.length} chars, exp=${expContent.length} chars');
        
        // Find first difference
        var minLen = origContent.length < expContent.length ? origContent.length : expContent.length;
        for (var i = 0; i < minLen; i++) {
          if (origContent[i] != expContent[i]) {
            var start = i > 30 ? i - 30 : 0;
            print('  First diff at position $i');
            print('  Orig: ...${origContent.substring(start, i + 50 > origContent.length ? origContent.length : i + 50)}...');
            print('  Exp:  ...${expContent.substring(start, i + 50 > expContent.length ? expContent.length : i + 50)}...');
            break;
          }
        }
        if (origContent.length != expContent.length && origContent.substring(0, minLen) == expContent.substring(0, minLen)) {
          print('  Content identical up to shorter length, but sizes differ');
          if (expContent.length > origContent.length) {
            print('  Extra content in exported: ${expContent.substring(origContent.length, origContent.length + 200 > expContent.length ? expContent.length : origContent.length + 200)}');
          }
        }
        break;
      }
    }
  }
}
