part of excel;

/// Styling class for cells
// ignore: must_be_immutable
class _FontStyle extends Equatable {
  ExcelColor? _fontColorHex = ExcelColor.black;
  String? _fontFamily;
  FontScheme _fontScheme = FontScheme.Unset;
  bool _bold = false, _italic = false;
  Underline _underline = Underline.None;
  int? _fontSize;

  // ── Extended font properties ──
  bool _isStrikethrough = false;
  FontVerticalAlign _fontVerticalAlign = FontVerticalAlign.none;
  ColorValue? _fontColor;

  _FontStyle(
      {ExcelColor? fontColorHex = ExcelColor.black,
      int? fontSize,
      String? fontFamily,
      FontScheme fontScheme = FontScheme.Unset,
      bool bold = false,
      Underline underline = Underline.None,
      bool italic = false}) {
    _bold = bold;

    _fontSize = fontSize;

    _italic = italic;

    _fontFamily = fontFamily;

    _fontScheme = fontScheme;

    _underline = underline;

    if (fontColorHex != null) {
      _fontColorHex = _isColorAppropriate(fontColorHex.colorHex).excelColor;
    } else {
      _fontColorHex = ExcelColor.black;
    }
  }

  /// Get Font Color
  ExcelColor get fontColor {
    return _fontColorHex ?? ExcelColor.black;
  }

  /// Set Font Color
  set fontColor(ExcelColor? fontColorHex) {
    if (fontColorHex != null) {
      _fontColorHex = _isColorAppropriate(fontColorHex.colorHex).excelColor;
    } else {
      _fontColorHex = ExcelColor.black;
    }
  }

  /// `Get FontFamily`
  String? get fontFamily {
    return _fontFamily;
  }

  /// `Set FontFamily`
  set fontFamily(String? family) {
    _fontFamily = family;
  }

  ///`Get FontScheme`
  ///
  FontScheme get fontScheme {
    return _fontScheme;
  }

  ///`Set FontScheme`
  ///
  set fontScheme(FontScheme scheme) {
    _fontScheme = scheme;
  }

  /// Get Font Size
  int? get fontSize {
    return _fontSize;
  }

  /// Set Font Size
  set fontSize(int? _fs) {
    _fontSize = _fs;
  }

  /// Get `Underline`
  Underline get underline {
    return _underline;
  }

  /// set `Underline`
  set underline(Underline underline) {
    _underline = underline;
  }

  /// Get `Bold`
  bool get isBold {
    return _bold;
  }

  /// Set `Bold`
  set isBold(bool bold) {
    _bold = bold;
  }

  /// Get `Italic`
  bool get isItalic {
    return _italic;
  }

  /// Set `Italic`
  set isItalic(bool italic) {
    _italic = italic;
  }

  /// Get `Strikethrough`
  bool get isStrikethrough => _isStrikethrough;

  /// Set `Strikethrough`
  set isStrikethrough(bool value) => _isStrikethrough = value;

  /// Get `FontVerticalAlign`
  FontVerticalAlign get fontVerticalAlign => _fontVerticalAlign;

  /// Set `FontVerticalAlign`
  set fontVerticalAlign(FontVerticalAlign value) => _fontVerticalAlign = value;

  /// Get font color as ColorValue (theme color support)
  ColorValue? get fontColorCV => _fontColor;

  /// Set font color as ColorValue (theme color support)
  set fontColorCV(ColorValue? value) => _fontColor = value;

  @override
  List<Object?> get props => [
        _bold,
        _italic,
        _fontSize,
        _underline,
        _fontFamily,
        _fontScheme,
        _fontColorHex,
        _isStrikethrough,
        _fontVerticalAlign,
        _fontColor,
      ];
}
