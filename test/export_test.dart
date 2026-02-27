import 'dart:io';
import 'package:excel/excel.dart';
import 'package:test/test.dart';

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

      // セル値を一括比較（個別expectだとタイムアウトするため）
      var mismatches = <String>[];
      for (var r = 0; r < original.maxRows; r++) {
        for (var c = 0; c < original.maxColumns; c++) {
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
}
