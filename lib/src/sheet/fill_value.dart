part of excel;

/// Represents an OOXML pattern fill with pattern type, foreground color, and background color.
///
/// Supports all 19 OOXML pattern types: none, solid, mediumGray, darkGray,
/// lightGray, darkHorizontal, darkVertical, darkDown, darkUp, darkGrid,
/// darkTrellis, lightHorizontal, lightVertical, lightDown, lightUp,
/// lightGrid, lightTrellis, gray125, gray0625.
// ignore: must_be_immutable
class FillValue extends Equatable {
  /// OOXML pattern type name (e.g., "solid", "gray125", "darkGrid")
  final String patternType;

  /// Pattern foreground color
  final ColorValue? fgColor;

  /// Pattern background color
  final ColorValue? bgColor;

  const FillValue({
    required this.patternType,
    this.fgColor,
    this.bgColor,
  });

  @override
  List<Object?> get props => [patternType, fgColor, bgColor];
}
