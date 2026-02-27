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

/// cellXfs セクションから最後の xf 要素を取得する（新規追加されたスタイル）
XmlElement _getLastXf(XmlDocument stylesXml) {
  var cellXfs = stylesXml.findAllElements('cellXfs').first;
  var xfElements = cellXfs.findAllElements('xf').toList();
  return xfElements.last;
}

/// スタイルを設定してエクスポートし、styles.xml を返すヘルパー
XmlDocument _exportWithStyle(CellStyle style) {
  var excel = Excel.createExcel();
  var sheet = excel['Sheet1'];
  sheet.updateCell(
    CellIndex.indexByString('A1'),
    TextCellValue('test'),
    cellStyle: style,
  );
  var bytes = excel.encode()!;
  return _getStylesXml(bytes);
}

void main() {
  // ========================================
  // Task 4.1: apply* 属性と xfId テスト
  // ========================================
  group('Task 4.1: apply* 属性と xfId 出力', () {
    test('applyBorder: ボーダー付きセルで applyBorder="1" が出力される', () {
      var style = CellStyle(
        leftBorder: Border(borderStyle: BorderStyle.Thin),
      );
      var stylesXml = _exportWithStyle(style);
      var xf = _getLastXf(stylesXml);

      expect(xf.getAttribute('applyBorder'), equals('1'));
    });

    test('applyFont: フォント変更セルで applyFont="1" が出力される', () {
      var style = CellStyle(
        bold: true,
        fontSize: 14,
        fontFamily: 'Arial',
      );
      var stylesXml = _exportWithStyle(style);
      var xf = _getLastXf(stylesXml);

      expect(xf.getAttribute('applyFont'), equals('1'));
    });

    test('applyNumberFormat: 数値書式セルで applyNumberFormat="1" が出力される',
        () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle(
        bold: true,
        numberFormat: NumFormat.standard_2,
      );
      // 数値フォーマットには数値セル値を使う（テキスト値だと accepts() が false になりデフォルトにフォールバックされる）
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        DoubleCellValue(3.14),
        cellStyle: style,
      );
      var bytes = excel.encode()!;
      var stylesXml = _getStylesXml(bytes);
      var xf = _getLastXf(stylesXml);

      expect(xf.getAttribute('numFmtId'), equals('2'));
      expect(xf.getAttribute('applyNumberFormat'), equals('1'));
    });

    test('applyFill: 背景色セルで applyFill="1" が出力される', () {
      var style = CellStyle(
        backgroundColorHex: ExcelColor.fromHexString('#FF0000'),
      );
      var stylesXml = _exportWithStyle(style);
      var xf = _getLastXf(stylesXml);

      expect(xf.getAttribute('applyFill'), equals('1'));
    });

    test('applyAlignment: アライメント設定セルで applyAlignment="1" が出力される',
        () {
      var style = CellStyle(
        horizontalAlign: HorizontalAlign.Center,
        verticalAlign: VerticalAlign.Top,
      );
      var stylesXml = _exportWithStyle(style);
      var xf = _getLastXf(stylesXml);

      expect(xf.getAttribute('applyAlignment'), equals('1'));

      // alignment 子要素も検証
      var alignment = xf.findAllElements('alignment').first;
      expect(alignment.getAttribute('horizontal'), equals('center'));
      expect(alignment.getAttribute('vertical'), equals('top'));
    });

    test('xfId: xfId が元の値で出力される（デフォルト 0）', () {
      var style = CellStyle(bold: true);
      var stylesXml = _exportWithStyle(style);
      var xf = _getLastXf(stylesXml);

      // xfId 属性が存在すること
      expect(xf.getAttribute('xfId'), isNotNull);
      expect(xf.getAttribute('xfId'), equals('0'));
    });

    test('xfId: カスタム xfId が保持される', () {
      var style = CellStyle(bold: true);
      style.xfId = 3;
      var stylesXml = _exportWithStyle(style);
      var xf = _getLastXf(stylesXml);

      expect(xf.getAttribute('xfId'), equals('3'));
    });
  });

  // ========================================
  // Task 4.2: テーマカラー・塗りつぶし・保護・アライメント・フォント出力
  // ========================================
  group('Task 4.2: 拡張スタイル出力', () {
    test('テーマカラーフォント: color 要素に theme/tint が出力される', () {
      var style = CellStyle(bold: true);
      style.fontColorValue = ColorValue.fromTheme(4, tint: -0.25);

      var stylesXml = _exportWithStyle(style);

      // フォントセクションから最後のフォントを取得
      var fonts = stylesXml.findAllElements('fonts').first;
      var fontElements = fonts.findAllElements('font').toList();
      var lastFont = fontElements.last;
      var colorEl = lastFont.findAllElements('color').first;

      expect(colorEl.getAttribute('theme'), equals('4'));
      expect(colorEl.getAttribute('tint'), equals('-0.25'));
    });

    test('FillValue パターン塗りつぶし: patternType と fgColor/bgColor が出力される',
        () {
      var style = CellStyle();
      style.fill = FillValue(
        patternType: 'darkGrid',
        fgColor: ColorValue.rgb('FF0000FF'),
        bgColor: ColorValue.rgb('FFFFFFFF'),
      );

      var stylesXml = _exportWithStyle(style);

      // fills セクションから最後の fill を取得
      var fills = stylesXml.findAllElements('fills').first;
      var fillElements = fills.findAllElements('fill').toList();
      var lastFill = fillElements.last;
      var patternFill = lastFill.findAllElements('patternFill').first;

      expect(patternFill.getAttribute('patternType'), equals('darkGrid'));

      var fgColor = patternFill.findAllElements('fgColor').first;
      expect(fgColor.getAttribute('rgb'), equals('FF0000FF'));

      var bgColor = patternFill.findAllElements('bgColor').first;
      expect(bgColor.getAttribute('rgb'), equals('FFFFFFFF'));
    });

    test('テーマカラー塗りつぶし: fgColor に theme/tint が出力される', () {
      var style = CellStyle();
      style.fill = FillValue(
        patternType: 'solid',
        fgColor: ColorValue.fromTheme(6, tint: 0.4),
      );

      var stylesXml = _exportWithStyle(style);

      var fills = stylesXml.findAllElements('fills').first;
      var fillElements = fills.findAllElements('fill').toList();
      var lastFill = fillElements.last;
      var patternFill = lastFill.findAllElements('patternFill').first;
      var fgColor = patternFill.findAllElements('fgColor').first;

      expect(fgColor.getAttribute('theme'), equals('6'));
      expect(fgColor.getAttribute('tint'), equals('0.4'));
    });

    test('Protection: protection 要素に locked/hidden が出力される', () {
      var style = CellStyle(bold: true);
      style.protection = CellProtection(locked: false, hidden: true);

      var stylesXml = _exportWithStyle(style);
      var xf = _getLastXf(stylesXml);

      // applyProtection 属性
      expect(xf.getAttribute('applyProtection'), equals('1'));

      // protection 子要素
      var protectionEl = xf.findAllElements('protection').first;
      expect(protectionEl.getAttribute('locked'), equals('0'));
      expect(protectionEl.getAttribute('hidden'), equals('1'));
    });

    test('拡張アライメント: indent, readingOrder が alignment に出力される', () {
      var style = CellStyle(
        horizontalAlign: HorizontalAlign.Center,
      );
      style.indent = 2;
      style.readingOrder = 2;

      var stylesXml = _exportWithStyle(style);
      var xf = _getLastXf(stylesXml);

      expect(xf.getAttribute('applyAlignment'), equals('1'));

      var alignment = xf.findAllElements('alignment').first;
      expect(alignment.getAttribute('indent'), equals('2'));
      expect(alignment.getAttribute('readingOrder'), equals('2'));
    });

    test('拡張アライメント: justifyLastLine が alignment に出力される', () {
      var style = CellStyle(
        horizontalAlign: HorizontalAlign.Center,
      );
      style.justifyLastLine = true;

      var stylesXml = _exportWithStyle(style);
      var xf = _getLastXf(stylesXml);

      var alignment = xf.findAllElements('alignment').first;
      expect(alignment.getAttribute('justifyLastLine'), equals('1'));
    });

    test('取り消し線: strike 要素がフォントに出力される', () {
      var style = CellStyle(bold: true);
      style.isStrikethrough = true;

      var stylesXml = _exportWithStyle(style);

      var fonts = stylesXml.findAllElements('fonts').first;
      var fontElements = fonts.findAllElements('font').toList();
      var lastFont = fontElements.last;

      var strikeElements = lastFont.findAllElements('strike');
      expect(strikeElements, isNotEmpty);
    });

    test('vertAlign: vertAlign 要素がフォントに出力される', () {
      var style = CellStyle(bold: true);
      style.fontVerticalAlign = FontVerticalAlign.superscript;

      var stylesXml = _exportWithStyle(style);

      var fonts = stylesXml.findAllElements('fonts').first;
      var fontElements = fonts.findAllElements('font').toList();
      var lastFont = fontElements.last;

      var vertAlignEl = lastFont.findAllElements('vertAlign').first;
      expect(vertAlignEl.getAttribute('val'), equals('superscript'));
    });

    test('vertAlign subscript: subscript が正しく出力される', () {
      var style = CellStyle(bold: true);
      style.fontVerticalAlign = FontVerticalAlign.subscript;

      var stylesXml = _exportWithStyle(style);

      var fonts = stylesXml.findAllElements('fonts').first;
      var fontElements = fonts.findAllElements('font').toList();
      var lastFont = fontElements.last;

      var vertAlignEl = lastFont.findAllElements('vertAlign').first;
      expect(vertAlignEl.getAttribute('val'), equals('subscript'));
    });

    test('cellStyleXfs/cellStyles 透過: スタイル変更してもセクションが保持される', () {
      // テーマやスタイル名を含むテストファイルを読み込む
      var file = './test/test_resources/example.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);

      // 元のstyles.xmlからcellStyleXfs/cellStylesを確認
      var originalStylesXml = _getStylesXml(bytes);
      var hasCellStyleXfs =
          originalStylesXml.findAllElements('cellStyleXfs').isNotEmpty;
      var hasCellStyles =
          originalStylesXml.findAllElements('cellStyles').isNotEmpty;

      // スタイル変更を加えてエクスポート
      var sheet = excel.tables.values.first;
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('modified'),
        cellStyle: CellStyle(bold: true, fontSize: 16),
      );

      var exportedBytes = excel.encode()!;
      var exportedStylesXml = _getStylesXml(exportedBytes);

      // 元に存在していたセクションがエクスポート後も保持されること
      if (hasCellStyleXfs) {
        expect(
          exportedStylesXml.findAllElements('cellStyleXfs'),
          isNotEmpty,
          reason: 'cellStyleXfs セクションがエクスポート後に失われた',
        );
      }
      if (hasCellStyles) {
        expect(
          exportedStylesXml.findAllElements('cellStyles'),
          isNotEmpty,
          reason: 'cellStyles セクションがエクスポート後に失われた',
        );
      }
    });
  });

  // ========================================
  // 統合テスト: ラウンドトリップ + XML 検証
  // ========================================
  group('統合: ラウンドトリップスタイル保持', () {
    test('複合スタイルのラウンドトリップで全属性が保持される', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      var style = CellStyle(
        bold: true,
        italic: true,
        fontSize: 14,
        fontFamily: 'Arial',
        horizontalAlign: HorizontalAlign.Center,
        verticalAlign: VerticalAlign.Top,
        rotation: 45,
        textWrapping: TextWrapping.WrapText,
        leftBorder: Border(borderStyle: BorderStyle.Thin),
        topBorder: Border(borderStyle: BorderStyle.Double),
      );
      style.isStrikethrough = true;
      style.fontVerticalAlign = FontVerticalAlign.superscript;
      style.indent = 3;
      style.readingOrder = 1;
      style.protection = CellProtection(locked: false, hidden: true);

      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('styled'),
        cellStyle: style,
      );

      var bytes = excel.encode()!;
      var decoded = Excel.decodeBytes(bytes);
      var cell = decoded['Sheet1'].rows[0][0];
      var cs = cell!.cellStyle!;

      // 基本属性
      expect(cs.isBold, isTrue);
      expect(cs.isItalic, isTrue);
      expect(cs.fontSize, equals(14));
      expect(cs.fontFamily, equals('Arial'));
      expect(cs.horizontalAlignment, equals(HorizontalAlign.Center));
      expect(cs.verticalAlignment, equals(VerticalAlign.Top));
      expect(cs.rotation, equals(45));
      expect(cs.wrap, equals(TextWrapping.WrapText));

      // 拡張属性
      expect(cs.isStrikethrough, isTrue);
      expect(cs.fontVerticalAlign, equals(FontVerticalAlign.superscript));
      expect(cs.indent, equals(3));
      expect(cs.readingOrder, equals(1));

      // 保護
      expect(cs.protection, isNotNull);
      expect(cs.protection!.locked, isFalse);
      expect(cs.protection!.hidden, isTrue);
    });

    test('FillValue のラウンドトリップで塗りつぶし属性が保持される', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      var style = CellStyle();
      style.fill = FillValue(
        patternType: 'solid',
        fgColor: ColorValue.rgb('FF00FF00'),
      );

      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('fill test'),
        cellStyle: style,
      );

      var bytes = excel.encode()!;
      var decoded = Excel.decodeBytes(bytes);
      var cell = decoded['Sheet1'].rows[0][0];
      var cs = cell!.cellStyle!;

      expect(cs.fill, isNotNull);
      expect(cs.fill!.patternType, equals('solid'));
      expect(cs.fill!.fgColor, isNotNull);
      expect(cs.fill!.fgColor!.hexColor, equals('FF00FF00'));
    });

    test('テーマカラーフォントのラウンドトリップ', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];

      var style = CellStyle(bold: true);
      style.fontColorValue = ColorValue.fromTheme(5, tint: 0.6);

      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('theme font'),
        cellStyle: style,
      );

      var bytes = excel.encode()!;
      var decoded = Excel.decodeBytes(bytes);
      var cell = decoded['Sheet1'].rows[0][0];
      var cs = cell!.cellStyle!;

      expect(cs.fontColorValue, isNotNull);
      expect(cs.fontColorValue!.theme, equals(5));
      expect(cs.fontColorValue!.tint, equals(0.6));
    });
  });
}
