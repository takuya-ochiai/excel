import 'dart:io';
import 'package:excel/excel.dart';
import 'package:test/test.dart';

void main() {
  // ── Task 3.1: Alignment bug fix tests ──
  group('Alignment parsing bug fix', () {
    test('horizontal alignment Center is preserved in round-trip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle(horizontalAlign: HorizontalAlign.Center);
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('test'),
        cellStyle: style,
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var cell = decoded['Sheet1']
          .rows[0][0];
      expect(cell, isNotNull);
      expect(cell!.cellStyle?.horizontalAlignment, equals(HorizontalAlign.Center));
    });

    test('horizontal alignment Right is preserved in round-trip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle(horizontalAlign: HorizontalAlign.Right);
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('test'),
        cellStyle: style,
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var cell = decoded['Sheet1'].rows[0][0];
      expect(cell, isNotNull);
      expect(cell!.cellStyle?.horizontalAlignment, equals(HorizontalAlign.Right));
    });

    test('vertical alignment Top is preserved in round-trip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle(verticalAlign: VerticalAlign.Top);
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('test'),
        cellStyle: style,
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var cell = decoded['Sheet1'].rows[0][0];
      expect(cell, isNotNull);
      expect(cell!.cellStyle?.verticalAlignment, equals(VerticalAlign.Top));
    });

    test('vertical alignment Center is preserved in round-trip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle(verticalAlign: VerticalAlign.Center);
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('test'),
        cellStyle: style,
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var cell = decoded['Sheet1'].rows[0][0];
      expect(cell, isNotNull);
      expect(cell!.cellStyle?.verticalAlignment, equals(VerticalAlign.Center));
    });

    test('textRotation is preserved in round-trip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle(rotation: 45);
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('test'),
        cellStyle: style,
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var cell = decoded['Sheet1'].rows[0][0];
      expect(cell, isNotNull);
      expect(cell!.cellStyle?.rotation, equals(45));
    });
  });

  // ── Task 3.1: Underline parsing bug fix tests ──
  group('Underline parsing bug fix', () {
    test('Double underline is preserved in round-trip (not overwritten to Single)', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle(underline: Underline.Double);
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('test'),
        cellStyle: style,
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var cell = decoded['Sheet1'].rows[0][0];
      expect(cell, isNotNull);
      expect(cell!.cellStyle?.underline, equals(Underline.Double));
    });

    test('Single underline is preserved in round-trip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle(underline: Underline.Single);
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('test'),
        cellStyle: style,
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var cell = decoded['Sheet1'].rows[0][0];
      expect(cell, isNotNull);
      expect(cell!.cellStyle?.underline, equals(Underline.Single));
    });

    test('No underline is preserved in round-trip', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle(underline: Underline.None);
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('test'),
        cellStyle: style,
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var cell = decoded['Sheet1'].rows[0][0];
      expect(cell, isNotNull);
      expect(cell!.cellStyle?.underline, equals(Underline.None));
    });
  });

  // ── Task 3.2: Extended attribute parsing tests ──
  group('Extended attribute parsing', () {
    test('strikethrough is parsed from font and preserved on CellStyle', () {
      // Create an Excel, set strikethrough via CellStyle extended property, encode, decode
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle();
      style.isStrikethrough = true;
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('strikethrough'),
        cellStyle: style,
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var cell = decoded['Sheet1'].rows[0][0];
      expect(cell, isNotNull);
      expect(cell!.cellStyle?.isStrikethrough, isTrue);
    });

    test('fontVerticalAlign superscript is parsed and preserved', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle();
      style.fontVerticalAlign = FontVerticalAlign.superscript;
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('super'),
        cellStyle: style,
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var cell = decoded['Sheet1'].rows[0][0];
      expect(cell, isNotNull);
      expect(cell!.cellStyle?.fontVerticalAlign, equals(FontVerticalAlign.superscript));
    });

    test('fontVerticalAlign subscript is parsed and preserved', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle();
      style.fontVerticalAlign = FontVerticalAlign.subscript;
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('sub'),
        cellStyle: style,
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var cell = decoded['Sheet1'].rows[0][0];
      expect(cell, isNotNull);
      expect(cell!.cellStyle?.fontVerticalAlign, equals(FontVerticalAlign.subscript));
    });

    test('indent is parsed from alignment and preserved', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle();
      style.indent = 2;
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('indented'),
        cellStyle: style,
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var cell = decoded['Sheet1'].rows[0][0];
      expect(cell, isNotNull);
      expect(cell!.cellStyle?.indent, equals(2));
    });

    test('readingOrder is parsed from alignment and preserved', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle();
      style.readingOrder = 2;
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('rtl'),
        cellStyle: style,
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var cell = decoded['Sheet1'].rows[0][0];
      expect(cell, isNotNull);
      expect(cell!.cellStyle?.readingOrder, equals(2));
    });

    test('theme color on font is parsed and preserved as ColorValue', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle();
      style.fontColorValue = ColorValue.fromTheme(4, tint: 0.5);
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('themed'),
        cellStyle: style,
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var cell = decoded['Sheet1'].rows[0][0];
      expect(cell, isNotNull);
      expect(cell!.cellStyle?.fontColorValue, isNotNull);
      expect(cell!.cellStyle!.fontColorValue!.isThemeColor, isTrue);
      expect(cell!.cellStyle!.fontColorValue!.theme, equals(4));
      expect(cell!.cellStyle!.fontColorValue!.tint, equals(0.5));
    });

    test('fill with patternType and fgColor/bgColor is parsed correctly', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle();
      style.fill = FillValue(
        patternType: 'darkGrid',
        fgColor: ColorValue.rgb('FF0000FF'),
        bgColor: ColorValue.rgb('FFFFFFFF'),
      );
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('patterned'),
        cellStyle: style,
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var cell = decoded['Sheet1'].rows[0][0];
      expect(cell, isNotNull);
      expect(cell!.cellStyle?.fill, isNotNull);
      expect(cell!.cellStyle!.fill!.patternType, equals('darkGrid'));
      expect(cell!.cellStyle!.fill!.fgColor?.hexColor, equals('FF0000FF'));
      expect(cell!.cellStyle!.fill!.bgColor?.hexColor, equals('FFFFFFFF'));
    });

    test('protection attributes are parsed and preserved', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle();
      style.protection = CellProtection(locked: false, hidden: true);
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('protected'),
        cellStyle: style,
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var cell = decoded['Sheet1'].rows[0][0];
      expect(cell, isNotNull);
      expect(cell!.cellStyle?.protection, isNotNull);
      expect(cell!.cellStyle!.protection!.locked, isFalse);
      expect(cell!.cellStyle!.protection!.hidden, isTrue);
    });

    test('xfId is parsed and preserved (not hardcoded to 0)', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle();
      style.xfId = 3;
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('xfId3'),
        cellStyle: style,
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var cell = decoded['Sheet1'].rows[0][0];
      expect(cell, isNotNull);
      expect(cell!.cellStyle?.xfId, equals(3));
    });

    test('cellStyleXfs raw XML is preserved through round-trip', () {
      // Load a file that has cellStyleXfs, modify a style, re-export, and verify valid output
      var file = './test/test_resources/example.xlsx';
      if (!File(file).existsSync()) return;
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      // Trigger style changes
      var style = CellStyle(bold: true);
      excel['Sheet1'].updateCell(
        CellIndex.indexByString('Z1'),
        TextCellValue('test'),
        cellStyle: style,
      );
      // Encode and re-decode should not throw
      var encoded = excel.encode();
      expect(encoded, isNotNull);
      var decoded = Excel.decodeBytes(encoded!);
      expect(decoded.sheets, isNotEmpty);
    });

    test('border theme color is parsed via ColorValue', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle(
        leftBorder: Border(
          borderStyle: BorderStyle.Thin,
          borderColor: ColorValue.fromTheme(1),
        ),
      );
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('bordered'),
        cellStyle: style,
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var cell = decoded['Sheet1'].rows[0][0];
      expect(cell, isNotNull);
      expect(cell!.cellStyle?.leftBorder.borderColor, isNotNull);
      expect(cell!.cellStyle!.leftBorder.borderColor!.isThemeColor, isTrue);
      expect(cell!.cellStyle!.leftBorder.borderColor!.theme, equals(1));
    });

    test('fill with theme color references is parsed correctly', () {
      var excel = Excel.createExcel();
      var sheet = excel['Sheet1'];
      var style = CellStyle();
      style.fill = FillValue(
        patternType: 'solid',
        fgColor: ColorValue.fromTheme(4, tint: -0.25),
      );
      sheet.updateCell(
        CellIndex.indexByString('A1'),
        TextCellValue('themed fill'),
        cellStyle: style,
      );

      var bytes = excel.encode();
      var decoded = Excel.decodeBytes(bytes!);
      var cell = decoded['Sheet1'].rows[0][0];
      expect(cell, isNotNull);
      expect(cell!.cellStyle?.fill, isNotNull);
      expect(cell!.cellStyle!.fill!.fgColor?.isThemeColor, isTrue);
      expect(cell!.cellStyle!.fill!.fgColor?.theme, equals(4));
      expect(cell!.cellStyle!.fill!.fgColor?.tint, equals(-0.25));
    });
  });
}
