import 'dart:io';
import 'package:archive/archive.dart';
import 'package:xml/xml.dart';

void main() {
  var files = {
    'original': 'test/test_resources/report-template.xlsm',
    'exported': 'tmp/passthrough_out.xlsm',
  };

  files.forEach((label, path) {
    var bytes = File(path).readAsBytesSync();
    var archive = ZipDecoder().decodeBytes(bytes);
    
    for (var f in archive) {
      if (f.name != 'xl/worksheets/sheet1.xml') continue;
      f.decompress();
      var content = String.fromCharCodes(f.content);
      var doc = XmlDocument.parse(content);
      var worksheet = doc.findAllElements('worksheet').first;
      
      print('--- $label worksheet child elements (top-level order) ---');
      for (var child in worksheet.children) {
        if (child is XmlElement) {
          var attrs = child.attributes.map((a) => '${a.name}="${a.value}"').join(' ');
          var childCount = child.children.whereType<XmlElement>().length;
          print('  <${child.name}${attrs.isNotEmpty ? " $attrs" : ""}> ($childCount children)');
        }
      }
      print('');
      break;
    }
  });
}
