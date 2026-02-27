part of excel;

/// Represents cell protection attributes (locked / hidden).
///
/// Default values follow Excel's standard behavior:
/// - locked: true (cells are locked by default)
/// - hidden: false (formulas are visible by default)
// ignore: must_be_immutable
class CellProtection extends Equatable {
  /// Whether the cell is locked
  final bool locked;

  /// Whether the cell's formula is hidden
  final bool hidden;

  const CellProtection({
    this.locked = true,
    this.hidden = false,
  });

  @override
  List<Object?> get props => [locked, hidden];
}
