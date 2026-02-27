part of excel;

/// Represents a color as either an RGB hex value or a theme color reference.
///
/// This immutable value object unifies multiple OOXML color representations:
/// - RGB hex ("AARRGGBB" format, e.g., "FF000000")
/// - Theme color reference (theme index + optional tint)
/// - Legacy indexed color (0-63)
/// - Automatic color flag
// ignore: must_be_immutable
class ColorValue extends Equatable {
  /// "AARRGGBB" format RGB value (e.g., "FF000000")
  final String? hexColor;

  /// Theme color index (0-11)
  final int? theme;

  /// Theme color tint adjustment (-1.0 to 1.0)
  final double? tint;

  /// Legacy indexed color (0-63)
  final int? indexed;

  /// Automatic color flag
  final bool? auto;

  /// Whether this color is a theme color reference
  bool get isThemeColor => theme != null;

  /// Whether this color is an RGB color (hex set, no theme)
  bool get isRgbColor => hexColor != null && theme == null;

  const ColorValue({
    this.hexColor,
    this.theme,
    this.tint,
    this.indexed,
    this.auto,
  });

  /// Creates an RGB color from a hex string
  const ColorValue.rgb(String hex)
      : hexColor = hex,
        theme = null,
        tint = null,
        indexed = null,
        auto = null;

  /// Creates a theme color reference
  const ColorValue.fromTheme(int themeIndex, {double? tint})
      : theme = themeIndex,
        tint = tint,
        hexColor = null,
        indexed = null,
        auto = null;

  @override
  List<Object?> get props => [hexColor, theme, tint, indexed, auto];
}
