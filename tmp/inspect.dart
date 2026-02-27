import 'dart:io';
import 'package:archive/archive.dart';

void main() {
  var bytes = File('tmp/passthrough_out.xlsm').readAsBytesSync();
  var archive = ZipDecoder().decodeBytes(bytes);
  for (var file in archive) {
    if (file.name == 'xl/worksheets/sheet1.xml') {
      file.decompress();
      var content = String.fromCharCodes(file.content);
      print('=== First 800 chars ===');
      print(content.substring(0, content.length > 800 ? 800 : content.length));
      print('---');
      var contentBytes = file.content as List<int>;
      print('=== First 50 bytes ===');
      print(contentBytes.take(50).toList());
      break;
    }
  }
}
