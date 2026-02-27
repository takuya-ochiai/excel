import 'package:excel/excel.dart';
import 'package:test/test.dart';

void main() {
  // ── Task 1.1: ColorValue Tests ──
  group('ColorValue', () {
    test('RGB constructor creates RGB color', () {
      final color = ColorValue.rgb('FF000000');
      expect(color.hexColor, equals('FF000000'));
      expect(color.theme, isNull);
      expect(color.tint, isNull);
      expect(color.indexed, isNull);
      expect(color.auto, isNull);
    });

    test('fromTheme constructor creates theme color', () {
      final color = ColorValue.fromTheme(1, tint: -0.25);
      expect(color.theme, equals(1));
      expect(color.tint, equals(-0.25));
      expect(color.hexColor, isNull);
      expect(color.indexed, isNull);
      expect(color.auto, isNull);
    });

    test('fromTheme without tint', () {
      final color = ColorValue.fromTheme(4);
      expect(color.theme, equals(4));
      expect(color.tint, isNull);
    });

    test('named constructor with all fields', () {
      final color = ColorValue(
        hexColor: 'FFFF0000',
        theme: 3,
        tint: 0.5,
        indexed: 10,
        auto: true,
      );
      expect(color.hexColor, equals('FFFF0000'));
      expect(color.theme, equals(3));
      expect(color.tint, equals(0.5));
      expect(color.indexed, equals(10));
      expect(color.auto, isTrue);
    });

    test('isThemeColor returns true when theme is set', () {
      final color = ColorValue.fromTheme(1);
      expect(color.isThemeColor, isTrue);
      expect(color.isRgbColor, isFalse);
    });

    test('isRgbColor returns true when hexColor is set without theme', () {
      final color = ColorValue.rgb('FF000000');
      expect(color.isRgbColor, isTrue);
      expect(color.isThemeColor, isFalse);
    });

    test('isRgbColor returns false when both hexColor and theme are set', () {
      final color = ColorValue(hexColor: 'FF000000', theme: 1);
      expect(color.isRgbColor, isFalse);
      expect(color.isThemeColor, isTrue);
    });

    test('equality: same RGB colors are equal', () {
      final a = ColorValue.rgb('FF000000');
      final b = ColorValue.rgb('FF000000');
      expect(a, equals(b));
    });

    test('equality: different RGB colors are not equal', () {
      final a = ColorValue.rgb('FF000000');
      final b = ColorValue.rgb('FFFF0000');
      expect(a, isNot(equals(b)));
    });

    test('equality: same theme colors are equal', () {
      final a = ColorValue.fromTheme(1, tint: -0.25);
      final b = ColorValue.fromTheme(1, tint: -0.25);
      expect(a, equals(b));
    });

    test('equality: different theme colors are not equal', () {
      final a = ColorValue.fromTheme(1, tint: -0.25);
      final b = ColorValue.fromTheme(2, tint: -0.25);
      expect(a, isNot(equals(b)));
    });

    test('equality: RGB and theme colors are not equal', () {
      final a = ColorValue.rgb('FF000000');
      final b = ColorValue.fromTheme(1);
      expect(a, isNot(equals(b)));
    });

    test('indexed color', () {
      final color = ColorValue(indexed: 64);
      expect(color.indexed, equals(64));
      expect(color.isThemeColor, isFalse);
      expect(color.isRgbColor, isFalse);
    });

    test('auto color', () {
      final color = ColorValue(auto: true);
      expect(color.auto, isTrue);
      expect(color.isThemeColor, isFalse);
      expect(color.isRgbColor, isFalse);
    });
  });

  // ── Task 1.2: CellProtection Tests ──
  group('CellProtection', () {
    test('default values: locked=true, hidden=false', () {
      final protection = CellProtection();
      expect(protection.locked, isTrue);
      expect(protection.hidden, isFalse);
    });

    test('custom values', () {
      final protection = CellProtection(locked: false, hidden: true);
      expect(protection.locked, isFalse);
      expect(protection.hidden, isTrue);
    });

    test('equality: same values are equal', () {
      final a = CellProtection(locked: true, hidden: false);
      final b = CellProtection(locked: true, hidden: false);
      expect(a, equals(b));
    });

    test('equality: different values are not equal', () {
      final a = CellProtection(locked: true, hidden: false);
      final b = CellProtection(locked: false, hidden: true);
      expect(a, isNot(equals(b)));
    });

    test('default constructor equals explicit defaults', () {
      final a = CellProtection();
      final b = CellProtection(locked: true, hidden: false);
      expect(a, equals(b));
    });
  });

  // ── Task 1.2: FontVerticalAlign Tests ──
  group('FontVerticalAlign', () {
    test('enum has three values', () {
      expect(FontVerticalAlign.values.length, equals(3));
    });

    test('enum values exist', () {
      expect(FontVerticalAlign.none, isNotNull);
      expect(FontVerticalAlign.superscript, isNotNull);
      expect(FontVerticalAlign.subscript, isNotNull);
    });
  });

  // ── Task 1.3: FillValue Tests ──
  group('FillValue', () {
    test('solid fill with fgColor only', () {
      final fill = FillValue(
        patternType: 'solid',
        fgColor: ColorValue.rgb('FFFF0000'),
      );
      expect(fill.patternType, equals('solid'));
      expect(fill.fgColor, equals(ColorValue.rgb('FFFF0000')));
      expect(fill.bgColor, isNull);
    });

    test('pattern fill with both fgColor and bgColor', () {
      final fill = FillValue(
        patternType: 'darkGrid',
        fgColor: ColorValue.fromTheme(4),
        bgColor: ColorValue.rgb('FFFFFFFF'),
      );
      expect(fill.patternType, equals('darkGrid'));
      expect(fill.fgColor, equals(ColorValue.fromTheme(4)));
      expect(fill.bgColor, equals(ColorValue.rgb('FFFFFFFF')));
    });

    test('none pattern type', () {
      final fill = FillValue(patternType: 'none');
      expect(fill.patternType, equals('none'));
      expect(fill.fgColor, isNull);
      expect(fill.bgColor, isNull);
    });

    test('gray125 pattern type', () {
      final fill = FillValue(patternType: 'gray125');
      expect(fill.patternType, equals('gray125'));
    });

    test('equality: same fills are equal', () {
      final a = FillValue(
        patternType: 'solid',
        fgColor: ColorValue.rgb('FFFF0000'),
      );
      final b = FillValue(
        patternType: 'solid',
        fgColor: ColorValue.rgb('FFFF0000'),
      );
      expect(a, equals(b));
    });

    test('equality: different pattern types are not equal', () {
      final a = FillValue(patternType: 'solid');
      final b = FillValue(patternType: 'gray125');
      expect(a, isNot(equals(b)));
    });

    test('equality: different colors are not equal', () {
      final a = FillValue(
        patternType: 'solid',
        fgColor: ColorValue.rgb('FFFF0000'),
      );
      final b = FillValue(
        patternType: 'solid',
        fgColor: ColorValue.rgb('FF0000FF'),
      );
      expect(a, isNot(equals(b)));
    });

    test('fill with theme color fgColor', () {
      final fill = FillValue(
        patternType: 'solid',
        fgColor: ColorValue.fromTheme(4, tint: 0.5),
      );
      expect(fill.fgColor!.isThemeColor, isTrue);
      expect(fill.fgColor!.theme, equals(4));
      expect(fill.fgColor!.tint, equals(0.5));
    });

    test('fill preserves bgColor separately from fgColor', () {
      final fg = ColorValue.rgb('FFFF0000');
      final bg = ColorValue.rgb('FF0000FF');
      final fill = FillValue(
        patternType: 'solid',
        fgColor: fg,
        bgColor: bg,
      );
      expect(fill.fgColor, isNot(equals(fill.bgColor)));
      expect(fill.fgColor, equals(fg));
      expect(fill.bgColor, equals(bg));
    });
  });
}
