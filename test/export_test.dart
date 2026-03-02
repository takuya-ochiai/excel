import 'dart:convert';
import 'dart:io';
import 'package:archive/archive.dart';
import 'package:excel/excel.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

void main() {
  test('xlsmファイルを読み込めること', () {
    var file = './test/test_resources/report-template.xlsm';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);

    expect(excel.tables.keys, isNotEmpty);
    for (var sheetName in excel.tables.keys) {
      var sheet = excel.tables[sheetName]!;
      print('$sheetName: ${sheet.maxRows} rows x ${sheet.maxColumns} cols');
    }
  });

  test('xlsmファイルを読み込み、変更せずそのままエクスポートする', () {
    var file = './test/test_resources/report-template.xlsm';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);

    // 何も変更せずエンコード
    var exportedBytes = excel.encode()!;

    // エクスポートしたファイルを保存
    var outPath = Directory.current.path + '/tmp/passthrough_out.xlsm';
    File(outPath)
      ..createSync(recursive: true)
      ..writeAsBytesSync(exportedBytes);
    print('Exported to: $outPath');

    // 再度読み込んで元データと一致するか検証
    var excelAgain = Excel.decodeBytes(exportedBytes);

    // シート数が同じ
    expect(excelAgain.tables.keys.toSet(), equals(excel.tables.keys.toSet()));

    // 各シートの行数・列数が同じ
    for (var sheetName in excel.tables.keys) {
      var original = excel.tables[sheetName]!;
      var exported = excelAgain.tables[sheetName]!;

      expect(exported.maxRows, equals(original.maxRows),
          reason: '$sheetName: row count mismatch');
      expect(exported.maxColumns, equals(original.maxColumns),
          reason: '$sheetName: column count mismatch');

      // セル値を比較（A～W列、1～21行に限定）
      var maxRow = original.maxRows < 21 ? original.maxRows : 21;
      var maxCol = original.maxColumns < 23 ? original.maxColumns : 23;
      var mismatches = <String>[];
      for (var r = 0; r < maxRow; r++) {
        for (var c = 0; c < maxCol; c++) {
          var origVal = original.rows[r][c]?.value?.toString() ?? '';
          var expVal = exported.rows[r][c]?.value?.toString() ?? '';
          if (origVal != expVal) {
            mismatches.add('$sheetName($r,$c): "$origVal" != "$expVal"');
          }
        }
      }
      expect(mismatches, isEmpty, reason: mismatches.take(5).join('\n'));
    }
    print('Passthrough export verified successfully.');
  }, timeout: Timeout(Duration(minutes: 5)));

  test('xlsmファイルを読み込み、一部変更してエクスポートする', () {
    var file = './test/test_resources/report-template.xlsm';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    var sheetName = excel.tables.keys.first;
    var sheet = excel.tables[sheetName]!;

    // 5行目DE列の結合セルに`2021026400`
    sheet.updateCell(
        CellIndex.indexByString('D5'), TextCellValue('2021026400'));
    // 5行目G列のセルに`00`
    sheet.updateCell(CellIndex.indexByString('G5'), TextCellValue('00'));
    // 5行目I列のセルに`1`
    sheet.updateCell(CellIndex.indexByString('I5'), IntCellValue(1));
    // 5行目MNO列の結合セルに`968`
    sheet.updateCell(CellIndex.indexByString('M5'), IntCellValue(968));
    // 5行目RS列の結合セルに`943`
    sheet.updateCell(CellIndex.indexByString('R5'), IntCellValue(943));
    // 5行目VW列の結合セルに`0`
    sheet.updateCell(CellIndex.indexByString('V5'), IntCellValue(0));
    // 6行目F~W列の結合セルに`春日神社`
    sheet.updateCell(
        CellIndex.indexByString('F6'), TextCellValue('春日神社'));
    // 9行目P列のセルに`段切り`
    sheet.updateCell(
        CellIndex.indexByString('P9'), TextCellValue('段切り'));
    // 12行目CD列の結合セルに`杉`
    sheet.updateCell(CellIndex.indexByString('C12'), TextCellValue('杉'));
    // 12行目E列のセルに`34`
    sheet.updateCell(CellIndex.indexByString('E12'), IntCellValue(34));
    // 12行目F列のセルに`地役権外`
    sheet.updateCell(
        CellIndex.indexByString('F12'), TextCellValue('地役権外'));
    // 12行目G列のセルに`根切り`
    sheet.updateCell(
        CellIndex.indexByString('G12'), TextCellValue('根切り'));
    // 12行目H列のセルに`要`
    sheet.updateCell(CellIndex.indexByString('H12'), TextCellValue('要'));
    // 12行目I列のセルに`伐触木以外`
    sheet.updateCell(
        CellIndex.indexByString('I12'), TextCellValue('伐触木以外'));
    // 12行目P列のセルに`1.00`
    sheet.updateCell(
        CellIndex.indexByString('P12'), DoubleCellValue(1.00));

    // エンコード
    var exportedBytes = excel.encode()!;

    // エクスポートしたファイルを保存
    var outPath = Directory.current.path + '/tmp/modified_out.xlsm';
    File(outPath)
      ..createSync(recursive: true)
      ..writeAsBytesSync(exportedBytes);
    print('Exported to: $outPath');

    // 再読み込みして検証
    var excelAgain = Excel.decodeBytes(exportedBytes);
    var sheetAgain = excelAgain.tables[sheetName]!;

    // 変更したセルの値を検証
    var expectations = <String, String>{
      'D5': '2021026400',
      'G5': '00',
      'I5': '1',
      'M5': '968',
      'R5': '943',
      'V5': '0',
      'F6': '春日神社',
      'P9': '段切り',
      'C12': '杉',
      'E12': '34',
      'F12': '地役権外',
      'G12': '根切り',
      'H12': '要',
      'I12': '伐触木以外',
      'P12': '1',
    };

    var mismatches = <String>[];
    for (var entry in expectations.entries) {
      var cell = sheetAgain.cell(CellIndex.indexByString(entry.key));
      var actual = cell.value?.toString() ?? '';
      if (actual != entry.value) {
        mismatches.add('${entry.key}: expected "${entry.value}" but got "$actual"');
      }
    }
    expect(mismatches, isEmpty, reason: mismatches.join('\n'));
    print('Modified export verified successfully.');
  }, timeout: Timeout(Duration(minutes: 5)));

  test('headerFooterがextLstより前に配置されること', () {
    var file = './test/test_resources/report-template.xlsm';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);

    var exportedBytes = excel.encode()!;

    var archive = ZipDecoder().decodeBytes(exportedBytes);
    var sheetFiles = archive.files
        .where((f) => f.name.startsWith('xl/worksheets/sheet') && f.name.endsWith('.xml'));

    for (var sheetFile in sheetFiles) {
      var xmlContent = utf8.decode(sheetFile.content as List<int>);
      var document = XmlDocument.parse(xmlContent);
      var worksheet = document.findAllElements('worksheet').first;

      var children = worksheet.children.whereType<XmlElement>().toList();
      var headerFooterIndex = children.indexWhere((e) => e.name.local == 'headerFooter');
      var extLstIndex = children.indexWhere((e) => e.name.local == 'extLst');

      if (headerFooterIndex != -1 && extLstIndex != -1) {
        expect(headerFooterIndex, lessThan(extLstIndex),
            reason: '${sheetFile.name}: headerFooter must appear before extLst');
      }

      if (extLstIndex != -1) {
        expect(extLstIndex, equals(children.length - 1),
            reason: '${sheetFile.name}: extLst must be the last child element');
      }
    }
  });
}
