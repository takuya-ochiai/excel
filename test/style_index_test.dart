import 'dart:convert';
import 'dart:io';

import 'package:archive/archive.dart';
import 'package:excel/excel.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

/// エクスポートされた Excel バイト列から styles.xml の XmlDocument を取得する
XmlDocument _getStylesXml(List<int> bytes) {
  var archive = ZipDecoder().decodeBytes(bytes);
  var stylesFile =
      archive.files.firstWhere((f) => f.name == 'xl/styles.xml');
  var xmlContent = utf8.decode(stylesFile.content as List<int>);
  return XmlDocument.parse(xmlContent);
}

void main() {
  // ========================================
  // 未変更セルの s 属性保持 (要件 7.3)
  // ========================================
  group('未変更セルの s 属性保持', () {
    test('ロード→別セルにスタイル変更→エクスポート→再読込で未変更セルのスタイルが保持される',
        () {
      var file = './test/test_resources/example.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);

      var sheetName = excel.tables.keys.first;
      var sheet = excel[sheetName];

      // 元のセルのスタイル情報を記録（最初の数行のフォントスタイル）
      var originalStyles = <String, CellStyle>{};
      for (var row = 0; row < sheet.maxRows && row < 5; row++) {
        for (var col = 0; col < sheet.maxColumns && col < 5; col++) {
          var cell = sheet.rows[row][col];
          if (cell?.cellStyle != null) {
            var key = '${row}_$col';
            originalStyles[key] = cell!.cellStyle!;
          }
        }
      }

      // 別のセルにスタイル変更を加える
      sheet.updateCell(
        CellIndex.indexByString('Z99'),
        TextCellValue('new styled cell'),
        cellStyle: CellStyle(bold: true, fontSize: 20, fontFamily: 'Verdana'),
      );

      // エクスポート→再読込
      var exportedBytes = excel.encode()!;
      var reloaded = Excel.decodeBytes(exportedBytes);
      var reloadedSheet = reloaded[sheetName];

      // 元のセルのスタイルが保持されていることを確認
      for (var entry in originalStyles.entries) {
        var parts = entry.key.split('_');
        var row = int.parse(parts[0]);
        var col = int.parse(parts[1]);
        var cell = reloadedSheet.rows[row][col];
        expect(cell, isNotNull, reason: 'Cell at ($row, $col) should exist');
        expect(cell!.cellStyle, isNotNull,
            reason: 'Style at ($row, $col) should exist');
        expect(cell.cellStyle!.isBold, equals(entry.value.isBold),
            reason: 'Bold at ($row, $col) should be preserved');
        expect(cell.cellStyle!.isItalic, equals(entry.value.isItalic),
            reason: 'Italic at ($row, $col) should be preserved');
        expect(cell.cellStyle!.fontSize, equals(entry.value.fontSize),
            reason: 'FontSize at ($row, $col) should be preserved');
      }
    });
  });

  // ========================================
  // 新規スタイルの追加 (要件 7.4)
  // ========================================
  group('新規スタイルの追加', () {
    test('cellXfs count が正確に増加する', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      // 元の styles.xml の cellXfs count を記録
      var originalBytes = excel.encode()!;
      var originalStylesXml = _getStylesXml(originalBytes);
      var originalCount = int.parse(originalStylesXml
          .findAllElements('cellXfs')
          .first
          .getAttribute('count')!);

      // 新しいスタイルを2つ追加
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('style1'),
        cellStyle: CellStyle(bold: true, fontSize: 16),
      );
      sheet.updateCell(
        CellIndex.indexByString('A2'),
        TextCellValue('style2'),
        cellStyle: CellStyle(italic: true, fontSize: 18, fontFamily: 'Arial'),
      );

      var bytes = excel.encode()!;
      var stylesXml = _getStylesXml(bytes);
      var newCount = int.parse(stylesXml
          .findAllElements('cellXfs')
          .first
          .getAttribute('count')!);
      var xfElements =
          stylesXml.findAllElements('cellXfs').first.findAllElements('xf');

      // count 属性が実際の xf 要素数と一致
      expect(newCount, equals(xfElements.length));
      // 2つのスタイルが追加されたので count は +2
      expect(newCount, equals(originalCount + 2));
    });

    test('新規スタイルが再インポートで保持される', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      var style = CellStyle(
        bold: true,
        italic: true,
        fontSize: 14,
        fontFamily: 'Times New Roman',
        horizontalAlign: HorizontalAlign.Right,
      );

      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('test'),
        cellStyle: style,
      );

      var bytes = excel.encode()!;
      var decoded = Excel.decodeBytes(bytes);
      var cell = decoded['Sheet1'].rows[0][0];
      var cs = cell!.cellStyle!;

      expect(cs.isBold, isTrue);
      expect(cs.isItalic, isTrue);
      expect(cs.fontSize, equals(14));
      expect(cs.fontFamily, equals('Times New Roman'));
      expect(cs.horizontalAlignment, equals(HorizontalAlign.Right));
    });
  });

  // ========================================
  // フォントID解決 (要件 7.3, 7.5)
  // ========================================
  group('フォントID解決', () {
    test('同一フォントの複数スタイルが同じ fontId を共有する', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      // 同じフォントだが異なる背景色のスタイル
      var style1 = CellStyle(
        bold: true,
        fontSize: 14,
        fontFamily: 'Arial',
        backgroundColorHex: ExcelColor.fromHexString('#FF0000'),
      );
      var style2 = CellStyle(
        bold: true,
        fontSize: 14,
        fontFamily: 'Arial',
        backgroundColorHex: ExcelColor.fromHexString('#00FF00'),
      );

      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('red bg'),
        cellStyle: style1,
      );
      sheet.updateCell(
        CellIndex.indexByString('A2'),
        TextCellValue('green bg'),
        cellStyle: style2,
      );

      var bytes = excel.encode()!;
      var stylesXml = _getStylesXml(bytes);

      // cellXfs 内の最後の2つの xf 要素を取得
      var xfList =
          stylesXml.findAllElements('cellXfs').first.findAllElements('xf').toList();
      var lastXf = xfList.last;
      var secondLastXf = xfList[xfList.length - 2];

      // 同じフォントなので fontId が同じであるべき
      expect(lastXf.getAttribute('fontId'),
          equals(secondLastXf.getAttribute('fontId')));
      // 異なる背景色なので fillId は異なるべき
      expect(lastXf.getAttribute('fillId'),
          isNot(equals(secondLastXf.getAttribute('fillId'))));
    });
  });

  // ========================================
  // ボーダーID解決 (要件 7.3, 7.5)
  // ========================================
  group('ボーダーID解決', () {
    test('同一ボーダーの複数スタイルが同じ borderId を共有する', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      var border = Border(borderStyle: BorderStyle.Thin);

      // 同じボーダーだが異なるフォントのスタイル
      var style1 = CellStyle(
        bold: true,
        leftBorder: border,
        rightBorder: border,
      );
      var style2 = CellStyle(
        italic: true,
        leftBorder: border,
        rightBorder: border,
      );

      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('bold'),
        cellStyle: style1,
      );
      sheet.updateCell(
        CellIndex.indexByString('A2'),
        TextCellValue('italic'),
        cellStyle: style2,
      );

      var bytes = excel.encode()!;
      var stylesXml = _getStylesXml(bytes);

      var xfList =
          stylesXml.findAllElements('cellXfs').first.findAllElements('xf').toList();
      var lastXf = xfList.last;
      var secondLastXf = xfList[xfList.length - 2];

      // 同じボーダーなので borderId が同じであるべき
      expect(lastXf.getAttribute('borderId'),
          equals(secondLastXf.getAttribute('borderId')));
      // 異なるフォントなので fontId は異なるべき
      expect(lastXf.getAttribute('fontId'),
          isNot(equals(secondLastXf.getAttribute('fontId'))));
    });
  });

  // ========================================
  // フルラウンドトリップ (要件 7.1, 7.2)
  // ========================================
  group('フルラウンドトリップ', () {
    test('テーマカラー・保護・アライメント拡張・塗りつぶし・下線を含む完全ラウンドトリップ',
        () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      // テーマカラーフォント + 塗りつぶし
      var style1 = CellStyle(bold: true);
      style1.fontColorValue = ColorValue.fromTheme(4, tint: -0.25);
      style1.fill = FillValue(
        patternType: 'solid',
        fgColor: ColorValue.fromTheme(6, tint: 0.4),
      );

      // 保護 + アライメント拡張
      var style2 = CellStyle(
        horizontalAlign: HorizontalAlign.Center,
        verticalAlign: VerticalAlign.Top,
      );
      style2.protection = CellProtection(locked: false, hidden: true);
      style2.indent = 2;
      style2.readingOrder = 1;
      style2.justifyLastLine = true;

      // 下線 + 取り消し線
      var style3 = CellStyle(
        underline: Underline.Double,
      );
      style3.isStrikethrough = true;
      style3.fontVerticalAlign = FontVerticalAlign.superscript;

      sheet.updateCell(CellIndex.indexByString('A1'),
          TextCellValue('theme+fill'), cellStyle: style1);
      sheet.updateCell(CellIndex.indexByString('A2'),
          TextCellValue('prot+align'), cellStyle: style2);
      sheet.updateCell(CellIndex.indexByString('A3'),
          TextCellValue('underline+strike'), cellStyle: style3);

      var bytes = excel.encode()!;
      var decoded = Excel.decodeBytes(bytes);
      var reSheet = decoded['Sheet1'];

      // Cell A1: テーマカラー + 塗りつぶし
      var cs1 = reSheet.rows[0][0]!.cellStyle!;
      expect(cs1.isBold, isTrue);
      expect(cs1.fontColorValue, isNotNull);
      expect(cs1.fontColorValue!.theme, equals(4));
      expect(cs1.fontColorValue!.tint, equals(-0.25));
      expect(cs1.fill, isNotNull);
      expect(cs1.fill!.patternType, equals('solid'));
      expect(cs1.fill!.fgColor, isNotNull);
      expect(cs1.fill!.fgColor!.theme, equals(6));

      // Cell A2: 保護 + アライメント
      var cs2 = reSheet.rows[1][0]!.cellStyle!;
      expect(cs2.horizontalAlignment, equals(HorizontalAlign.Center));
      expect(cs2.verticalAlignment, equals(VerticalAlign.Top));
      expect(cs2.protection, isNotNull);
      expect(cs2.protection!.locked, isFalse);
      expect(cs2.protection!.hidden, isTrue);
      expect(cs2.indent, equals(2));
      expect(cs2.readingOrder, equals(1));
      expect(cs2.justifyLastLine, isTrue);

      // Cell A3: 下線 + 取り消し線
      var cs3 = reSheet.rows[2][0]!.cellStyle!;
      expect(cs3.underline, equals(Underline.Double));
      expect(cs3.isStrikethrough, isTrue);
      expect(cs3.fontVerticalAlign, equals(FontVerticalAlign.superscript));
    });
  });

  // ========================================
  // 変更なしパススルー (要件 7.1)
  // ========================================
  group('変更なしパススルー', () {
    test('スタイル変更なしで cellXfs count が実際の xf 要素数と一致する', () {
      var file = './test/test_resources/example.xlsx';
      var bytes = File(file).readAsBytesSync();

      // ロード→変更なし→エクスポート
      var excel = Excel.decodeBytes(bytes);
      var exportedBytes = excel.encode()!;
      var exportedStylesXml = _getStylesXml(exportedBytes);
      var exportedCellXfs =
          exportedStylesXml.findAllElements('cellXfs').first;
      var exportedCount =
          int.parse(exportedCellXfs.getAttribute('count')!);
      var exportedXfCount =
          exportedCellXfs.findAllElements('xf').length;

      // count 属性が実際の xf 要素数と一致（構造的整合性）
      expect(exportedCount, equals(exportedXfCount));

      // 再読込でクラッシュしない
      var reloaded = Excel.decodeBytes(exportedBytes);
      expect(reloaded.tables, isNotEmpty);
    });
  });

  // ========================================
  // フォントカウント正確性 (要件 7.5)
  // ========================================
  group('フォントカウント正確性', () {
    test('fonts count 属性が実際の font 要素数と一致する', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      // 複数の異なるフォントを使用
      sheet.updateCell(CellIndex.indexByString('A1'),
          TextCellValue('bold'), cellStyle: CellStyle(bold: true));
      sheet.updateCell(
          CellIndex.indexByString('A2'),
          TextCellValue('big'),
          cellStyle: CellStyle(fontSize: 24, fontFamily: 'Courier'));
      sheet.updateCell(
          CellIndex.indexByString('A3'),
          TextCellValue('italic'),
          cellStyle: CellStyle(italic: true, fontFamily: 'Georgia'));

      var bytes = excel.encode()!;
      var stylesXml = _getStylesXml(bytes);

      var fontsEl = stylesXml.findAllElements('fonts').first;
      var countAttr = int.parse(fontsEl.getAttribute('count')!);
      var actualFontCount = fontsEl.findAllElements('font').length;

      expect(countAttr, equals(actualFontCount));
    });

    test('既存ファイルのフォントカウントも正確', () {
      var file = './test/test_resources/example.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);

      // スタイル変更を加える
      var sheetName = excel.tables.keys.first;
      excel[sheetName].updateCell(
        CellIndex.indexByString('Z1'),
        TextCellValue('new font'),
        cellStyle: CellStyle(bold: true, fontSize: 22, fontFamily: 'Impact'),
      );

      var exportedBytes = excel.encode()!;
      var stylesXml = _getStylesXml(exportedBytes);

      var fontsEl = stylesXml.findAllElements('fonts').first;
      var countAttr = int.parse(fontsEl.getAttribute('count')!);
      var actualFontCount = fontsEl.findAllElements('font').length;

      expect(countAttr, equals(actualFontCount));
    });
  });

  // ========================================
  // テストリソース回帰 (要件 8.3)
  // ========================================
  group('テストリソース回帰', () {
    var testFiles = Directory('./test/test_resources')
        .listSync()
        .whereType<File>()
        .where((f) => f.path.endsWith('.xlsx'))
        .toList();

    for (var testFile in testFiles) {
      var fileName = testFile.path.split(Platform.pathSeparator).last;
      test('$fileName のラウンドトリップでクラッシュしない', () {
        var bytes = testFile.readAsBytesSync();
        var excel = Excel.decodeBytes(bytes);

        // エクスポートがクラッシュしないこと
        var exportedBytes = excel.encode();
        expect(exportedBytes, isNotNull);

        // 再読込がクラッシュしないこと
        var reloaded = Excel.decodeBytes(exportedBytes!);
        expect(reloaded.tables, isNotEmpty);

        // fonts count が正確であること（直接子要素のみカウント）
        var stylesXml = _getStylesXml(exportedBytes);
        var fontsEl = stylesXml.findAllElements('fonts').first;
        var countAttr = int.parse(fontsEl.getAttribute('count')!);
        var actualFontCount = fontsEl.findElements('font').length;
        expect(countAttr, equals(actualFontCount),
            reason: '$fileName: fonts count mismatch');

        // cellXfs count が xf 要素数と一致すること
        var cellXfs = stylesXml.findAllElements('cellXfs').first;
        var xfCount = int.parse(cellXfs.getAttribute('count')!);
        var actualXfCount = cellXfs.findAllElements('xf').length;
        expect(xfCount, equals(actualXfCount),
            reason: '$fileName: cellXfs count mismatch');
      });
    }
  });
}
