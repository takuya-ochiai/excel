import 'dart:io';
import 'package:excel/excel.dart';
import 'package:test/test.dart';

void main() {
  test('任意のExcelファイルを読み込んでシート情報を表示する', () {
    var file = './test/test_resources/example.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);

    expect(excel.tables, isNotEmpty);

    for (var table in excel.tables.keys) {
      var sheet = excel.tables[table]!;
      print('Sheet: $table (${sheet.maxRows} rows x ${sheet.maxColumns} cols)');
      for (var row in sheet.rows) {
        var values = row.map((cell) => cell?.value?.toString() ?? '').toList();
        print(values.join('\t'));
      }
    }
  });

  test('セル値を個別に取得する', () {
    var file = './test/test_resources/example.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);

    var sheet = excel['Sheet1'];
    var cellA1 = sheet.cell(CellIndex.indexByString('A1'));
    print('A1: ${cellA1.value}');
    expect(cellA1.value, isNotNull);
  });

  test('Excelファイルを読み込み、変更せずそのままエクスポートする', () {
    var file = './test/test_resources/example.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);

    // 何も変更せずエンコード
    var exportedBytes = excel.encode()!;

    // エクスポートしたファイルを保存
    var outPath = Directory.current.path + '/tmp/passthrough_out.xlsx';
    File(outPath)
      ..createSync(recursive: true)
      ..writeAsBytesSync(exportedBytes);
    print('Exported to: $outPath');

    // 再度読み込んで元データと一致するか検証
    var excelAgain = Excel.decodeBytes(exportedBytes);

    // シート数が同じ
    expect(excelAgain.tables.keys.toSet(), equals(excel.tables.keys.toSet()));

    // 各シートの行数・列数・セル値が同じ
    for (var sheetName in excel.tables.keys) {
      var original = excel.tables[sheetName]!;
      var exported = excelAgain.tables[sheetName]!;

      expect(exported.maxRows, equals(original.maxRows),
          reason: '$sheetName: row count mismatch');
      expect(exported.maxColumns, equals(original.maxColumns),
          reason: '$sheetName: column count mismatch');

      for (var r = 0; r < original.maxRows; r++) {
        for (var c = 0; c < original.maxColumns; c++) {
          var origVal = original.rows[r][c]?.value?.toString() ?? '';
          var expVal = exported.rows[r][c]?.value?.toString() ?? '';
          expect(expVal, equals(origVal),
              reason: '$sheetName: cell ($r, $c) mismatch');
        }
      }
    }
    print('Passthrough export verified successfully.');
  });

  test('全テストリソースファイルを読み込めることを確認する', () {
    var dir = Directory('./test/test_resources');
    var xlsxFiles = dir
        .listSync()
        .whereType<File>()
        .where((f) => f.path.endsWith('.xlsx'))
        .toList();

    expect(xlsxFiles, isNotEmpty);
    print('Found ${xlsxFiles.length} xlsx files');

    for (var file in xlsxFiles) {
      var bytes = file.readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      var sheetCount = excel.tables.keys.length;
      print('  ${file.path} -> $sheetCount sheets');
      expect(sheetCount, greaterThan(0));
    }
  });
}
