================================================
FILE: pubspec.yaml
================================================
name: excel
description: A flutter and dart library for reading, creating, editing and updating excel sheets with compatible both on client and server side.
version: 5.0.0
homepage: https://github.com/justkawal/excel
topics:
  - excel
  - office
  - storage
  - sheets
  - spreadsheet

environment:
  sdk: ">=3.6.0 <4.0.0"

dependencies:
  web: ^1.1.1
  archive: ^4.0.4
  xml: ">=5.0.0 <7.0.0"
  collection: ^1.15.0
  equatable: ^2.0.0

dev_dependencies:
  test: ^1.23.0
  lints: ^5.1.1



================================================
FILE: lib/excel.dart
================================================
library excel;

import 'dart:convert';
import 'dart:math';
import 'package:archive/archive.dart';
import 'package:collection/collection.dart';
import 'package:equatable/equatable.dart';
import 'package:xml/xml.dart';
import 'src/web_helper/client_save_excel.dart'
    if (dart.library.html) 'src/web_helper/web_save_excel_browser.dart'
    as helper;

/// main directory
part 'src/excel.dart';

/// sharedStrigns
part 'src/sharedStrings/shared_strings.dart';

/// Number Format
part 'src/number_format/num_format.dart';

/// Utilities
part 'src/utilities/span.dart';
part 'src/utilities/fast_list.dart';
part 'src/utilities/utility.dart';
part 'src/utilities/constants.dart';
part 'src/utilities/enum.dart';
part 'src/utilities/archive.dart';
part 'src/utilities/colors.dart';

/// Save
part 'src/save/save_file.dart';
part 'src/save/self_correct_span.dart';
part 'src/parser/parse.dart';

/// Sheet
part 'src/sheet/sheet.dart';
part 'src/sheet/font_family.dart';
part 'src/sheet/data_model.dart';
part 'src/sheet/cell_index.dart';
part 'src/sheet/cell_style.dart';
part 'src/sheet/font_style.dart';
part 'src/sheet/header_footer.dart';
part 'src/sheet/border_style.dart';



================================================
FILE: lib/src/excel.dart
================================================
part of excel;

Excel _newExcel(Archive archive) {
  // Lookup at file format
  String? format;

  var mimetype = archive.findFile('mimetype');
  if (mimetype == null) {
    var xl = archive.findFile('xl/workbook.xml');
    if (xl != null) {
      format = _spreasheetXlsx;
    }
  }

  switch (format) {
    case _spreasheetXlsx:
      return Excel._(archive);
    default:
      throw UnsupportedError(
          'Excel format unsupported. Only .xlsx files are supported');
  }
}

/// Decode a excel file.
class Excel {
  bool _styleChanges = false;
  bool _mergeChanges = false;
  bool _rtlChanges = false;

  Archive _archive;

  final Map<String, XmlNode> _sheets = {};
  final Map<String, XmlDocument> _xmlFiles = {};
  final Map<String, String> _xmlSheetId = {};
  final Map<String, Map<String, int>> _cellStyleReferenced = {};
  final Map<String, Sheet> _sheetMap = {};

  List<CellStyle> _cellStyleList = [];
  List<String> _patternFill = [];
  final List<String> _mergeChangeLook = [];
  final List<String> _rtlChangeLook = [];
  List<_FontStyle> _fontStyleList = [];
  final List<int> _numFmtIds = [];
  final NumFormatMaintainer _numFormats = NumFormatMaintainer();
  List<_BorderSet> _borderSetList = [];

  _SharedStringsMaintainer _sharedStrings = _SharedStringsMaintainer._();

  String _stylesTarget = '';
  String _sharedStringsTarget = '';
  String get _absSharedStringsTarget {
    if (_sharedStringsTarget.isNotEmpty && _sharedStringsTarget[0] == "/") {
      return _sharedStringsTarget.substring(1);
    }
    return "xl/${_sharedStringsTarget}";
  }

  String? _defaultSheet;
  late Parser parser;

  Excel._(this._archive) {
    parser = Parser._(this);
    parser._startParsing();
  }

  factory Excel.createExcel() {
    return Excel.decodeBytes(Base64Decoder().convert(_newSheet));
  }

  factory Excel.decodeBytes(List<int> data) {
    final Archive archive;
    try {
      archive = ZipDecoder().decodeBytes(data);
    } catch (e) {
      throw UnsupportedError(
          'Excel format unsupported. Only .xlsx files are supported');
    }
    return _newExcel(archive);
  }

  factory Excel.decodeBuffer(InputStream input) {
    return _newExcel(ZipDecoder().decodeStream(input));
  }

  ///
  ///It will return `tables` as map in order to mimic the previous versions reading the data.
  ///
  Map<String, Sheet> get tables {
    if (this._sheetMap.isEmpty) {
      _damagedExcel(text: "Corrupted Excel file.");
    }
    return Map<String, Sheet>.from(this._sheetMap);
  }

  ///
  ///It will return the SheetObject of `sheet`.
  ///
  ///If the `sheet` does not exist then it will create `sheet` with `New Sheet Object`
  ///
  Sheet operator [](String sheet) {
    _availSheet(sheet);
    return _sheetMap[sheet]!;
  }

  ///
  ///Returns the `Map<String, Sheet>`
  ///
  ///where `key` is the `Sheet Name` and the `value` is the `Sheet Object`
  ///
  Map<String, Sheet> get sheets {
    return Map<String, Sheet>.from(_sheetMap);
  }

  ///
  ///If `sheet` does not exist then it will be automatically created with contents of `sheetObject`
  ///
  ///Newly created sheet with name = `sheet` will have seperate reference and will not be linked to sheetObject.
  ///
  operator []=(String sheet, Sheet sheetObject) {
    _availSheet(sheet);

    _sheetMap[sheet] = Sheet._clone(this, sheet, sheetObject);
  }

  ///
  ///`sheet2Object` will be linked with `sheet1`.
  ///
  ///If `sheet1` does not exist then it will be automatically created.
  ///
  ///Important Note: After linkage the operations performed on `sheet1`, will also get performed on `sheet2Object` and `vica-versa`.
  ///
  void link(String sheet1, Sheet existingSheetObject) {
    if (_sheetMap[existingSheetObject.sheetName] != null) {
      _availSheet(sheet1);

      _sheetMap[sheet1] = _sheetMap[existingSheetObject.sheetName]!;

      if (_cellStyleReferenced[existingSheetObject.sheetName] != null) {
        _cellStyleReferenced[sheet1] = Map<String, int>.from(
            _cellStyleReferenced[existingSheetObject.sheetName]!);
      }
    }
  }

  ///
  ///If `sheet` is linked with any other sheet's object then it's link will be broke
  ///
  void unLink(String sheet) {
    if (_sheetMap[sheet] != null) {
      ///
      /// copying the sheet into itself thus resulting in breaking the linkage as Sheet._clone() will provide new reference;
      copy(sheet, sheet);
    }
  }

  ///
  ///Copies the content of `fromSheet` into `toSheet`.
  ///
  ///In order to successfully copy: `fromSheet` should exist in `excel.tables.keys`.
  ///
  ///If `toSheet` does not exist then it will be automatically created.
  ///
  void copy(String fromSheet, String toSheet) {
    _availSheet(toSheet);

    if (_sheetMap[fromSheet] != null) {
      this[toSheet] = this[fromSheet];
    }
    if (_cellStyleReferenced[fromSheet] != null) {
      _cellStyleReferenced[toSheet] =
          Map<String, int>.from(_cellStyleReferenced[fromSheet]!);
    }
  }

  ///
  ///Changes the name from `oldSheetName` to `newSheetName`.
  ///
  ///In order to rename : `oldSheetName` should exist in `excel.tables.keys` and `newSheetName` must not exist.
  ///
  void rename(String oldSheetName, String newSheetName) {
    if (_sheetMap[oldSheetName] != null && _sheetMap[newSheetName] == null) {
      ///
      /// rename from _defaultSheet var also
      if (_defaultSheet == oldSheetName) {
        _defaultSheet = newSheetName;
      }

      copy(oldSheetName, newSheetName);

      ///
      /// delete the `oldSheetName` as sheet with `newSheetName` is having cloned `SheetObject of oldSheetName` with new reference,
      delete(oldSheetName);
    }
  }

  ///
  ///If `sheet` exist in `excel.tables.keys` and `excel.tables.keys.length >= 2` then it will be `deleted`.
  ///
  void delete(String sheet) {
    ///
    /// remove the sheet `name` or `key` from the below locations if they exist.

    ///
    /// If it is not the last sheet then `delete` otherwise `return`;
    if (_sheetMap.length <= 1) {
      return;
    }

    ///
    ///remove from _defaultSheet var also
    if (_defaultSheet == sheet) {
      _defaultSheet = null;
    }

    ///
    /// remove the `Sheet Object` from `_sheetMap`.
    if (_sheetMap[sheet] != null) {
      _sheetMap.remove(sheet);
    }

    ///
    /// remove from `_mergeChangeLook`.
    if (_mergeChangeLook.contains(sheet)) {
      _mergeChangeLook.remove(sheet);
    }

    ///
    /// remove from `_rtlChangeLook`.
    if (_rtlChangeLook.contains(sheet)) {
      _rtlChangeLook.remove(sheet);
    }

    ///
    /// remove from `_xmlSheetId`.
    if (_xmlSheetId[sheet] != null) {
      String sheetId1 =
              "worksheets" + _xmlSheetId[sheet]!.split('worksheets')[1],
          sheetId2 = _xmlSheetId[sheet]!;

      _xmlFiles['xl/_rels/workbook.xml.rels']
          ?.rootElement
          .children
          .removeWhere((_sheetName) {
        return _sheetName.getAttribute('Target') != null &&
            _sheetName.getAttribute('Target') == sheetId1;
      });

      _xmlFiles['[Content_Types].xml']
          ?.rootElement
          .children
          .removeWhere((_sheetName) {
        return _sheetName.getAttribute('PartName') != null &&
            _sheetName.getAttribute('PartName') == '/' + sheetId2;
      });

      ///
      /// Also remove from the _xmlFiles list as we might want to create this sheet again from new starting.
      if (_xmlFiles[_xmlSheetId[sheet]] != null) {
        _xmlFiles.remove(_xmlSheetId[sheet]);
      }

      ///
      /// Maybe overkill and unsafe to do this, but works for now especially
      /// delete or renaming default sheet name (`Sheet1`),
      /// another safer method preferred
      _archive = _cloneArchive(
        _archive,
        _xmlFiles.map((k, v) {
          final encode = utf8.encode(v.toString());
          final value = ArchiveFile(k, encode.length, encode);
          return MapEntry(k, value);
        }),
        excludedFile: _xmlSheetId[sheet],
      );

      _xmlSheetId.remove(sheet);
    }

    ///
    /// remove from key = `sheet` from `_sheets`
    if (_sheets[sheet] != null) {
      ///
      /// Remove from `xl/workbook.xml`
      ///
      _xmlFiles['xl/workbook.xml']
          ?.findAllElements('sheets')
          .first
          .children
          .removeWhere((element) {
        return element.getAttribute('name') != null &&
            element.getAttribute('name').toString() == sheet;
      });

      _sheets.remove(sheet);
    }

    ///
    /// remove the cellStlye Referencing as it would be useless to have cellStyleReferenced saved
    if (_cellStyleReferenced[sheet] != null) {
      _cellStyleReferenced.remove(sheet);
    }
  }

  ///
  ///It will start setting the edited values of `sheets` into the `files` and then `exports the file`.
  ///
  List<int>? encode() {
    Save s = Save._(this, parser);
    return s._save();
  }

  /// Starts Saving the file.
  /// `On Web`
  /// ```
  /// // Call function save() to download the file
  /// var bytes = excel.save(fileName: "My_Excel_File_Name.xlsx");
  ///
  ///
  /// ```
  /// `On Android / iOS`
  ///
  /// For getting directory on Android or iOS, Use: [path_provider](https://pub.dev/packages/path_provider)
  /// ```
  /// // Call function save() to download the file
  /// var fileBytes = excel.save();
  /// var directory = await getApplicationDocumentsDirectory();
  ///
  /// File(join("$directory/output_file_name.xlsx"))
  ///   ..createSync(recursive: true)
  ///   ..writeAsBytesSync(fileBytes);
  ///
  ///```
  List<int>? save({String fileName = 'FlutterExcel.xlsx'}) {
    Save s = Save._(this, parser);
    var onValue = s._save();
    return helper.SavingHelper.saveFile(onValue, fileName);
  }

  ///
  ///returns the name of the `defaultSheet` (the sheet which opens firstly when xlsx file is opened in `excel based software`).
  ///
  String? getDefaultSheet() {
    if (_defaultSheet != null) {
      return _defaultSheet;
    } else {
      String? re = _getDefaultSheet();
      return re;
    }
  }

  ///
  ///Internal function which returns the defaultSheet-Name by reading from `workbook.xml`
  ///
  String? _getDefaultSheet() {
    Iterable<XmlElement>? elements =
        _xmlFiles['xl/workbook.xml']?.findAllElements('sheet');
    XmlElement? _sheet;
    if (elements?.isNotEmpty ?? false) {
      _sheet = elements?.first;
    }

    if (_sheet != null) {
      var defaultSheet = _sheet.getAttribute('name');
      if (defaultSheet != null) {
        return defaultSheet;
      } else {
        _damagedExcel(
            text: 'Excel sheet corrupted!! Try creating new excel file.');
      }
    }
    return null;
  }

  ///
  ///It returns `true` if the passed `sheetName` is successfully set to `default opening sheet` otherwise returns `false`.
  ///
  bool setDefaultSheet(String sheetName) {
    if (_sheetMap[sheetName] != null) {
      _defaultSheet = sheetName;
      return true;
    }
    return false;
  }

  ///
  ///Inserts an empty `column` in sheet at position = `columnIndex`.
  ///
  ///If `columnIndex == null` or `columnIndex < 0` if will not execute
  ///
  ///If the `sheet` does not exists then it will be created automatically.
  ///
  void insertColumn(String sheet, int columnIndex) {
    if (columnIndex < 0) {
      return;
    }
    _availSheet(sheet);
    _sheetMap[sheet]!.insertColumn(columnIndex);
  }

  ///
  ///If `sheet` exists and `columnIndex < maxColumns` then it removes column at index = `columnIndex`
  ///
  void removeColumn(String sheet, int columnIndex) {
    if (columnIndex >= 0 && _sheetMap[sheet] != null) {
      _sheetMap[sheet]!.removeColumn(columnIndex);
    }
  }

  ///
  ///Inserts an empty row in `sheet` at position = `rowIndex`.
  ///
  ///If `rowIndex == null` or `rowIndex < 0` if will not execute
  ///
  ///If the `sheet` does not exists then it will be created automatically.
  ///
  void insertRow(String sheet, int rowIndex) {
    if (rowIndex < 0) {
      return;
    }
    _availSheet(sheet);
    _sheetMap[sheet]!.insertRow(rowIndex);
  }

  ///
  ///If `sheet` exists and `rowIndex < maxRows` then it removes row at index = `rowIndex`
  ///
  void removeRow(String sheet, int rowIndex) {
    if (rowIndex >= 0 && _sheetMap[sheet] != null) {
      _sheetMap[sheet]!.removeRow(rowIndex);
    }
  }

  ///
  ///Appends [row] iterables just post the last filled index in the [sheet]
  ///
  ///If `sheet` does not exist then it will be automatically created.
  ///
  void appendRow(String sheet, List<CellValue?> row) {
    if (row.isEmpty) {
      return;
    }
    _availSheet(sheet);
    int targetRow = _sheetMap[sheet]!.maxRows;
    insertRowIterables(sheet, row, targetRow);
  }

  ///
  ///If `sheet` does not exist then it will be automatically created.
  ///
  ///Adds the [row] iterables in the given rowIndex = [rowIndex] in [sheet]
  ///
  ///[startingColumn] tells from where we should start putting the [row] iterables
  ///
  ///[overwriteMergedCells] when set to [true] will over-write mergedCell and does not jumps to next unqiue cell.
  ///
  ///[overwriteMergedCells] when set to [false] puts the cell value to next unique cell available by putting the value in merged cells only once and jumps to next unique cell.
  ///
  void insertRowIterables(String sheet, List<CellValue?> row, int rowIndex,
      {int startingColumn = 0, bool overwriteMergedCells = true}) {
    if (rowIndex < 0) {
      return;
    }
    _availSheet(sheet);
    _sheetMap['$sheet']!.insertRowIterables(row, rowIndex,
        startingColumn: startingColumn,
        overwriteMergedCells: overwriteMergedCells);
  }

  ///
  ///Returns the `count` of replaced `source` with `target`
  ///
  ///`source` is Pattern which allows you to pass your custom `RegExp` or a `String` providing more control over it.
  ///
  ///optional argument `first` is used to replace the number of first earlier occurrences
  ///
  ///If `first` is set to `3` then it will replace only first `3 occurrences` of the `source` with `target`.
  ///
  ///       excel.findAndReplace('MySheetName', 'sad', 'happy', first: 3);
  ///
  ///       or
  ///
  ///       var mySheet = excel['mySheetName'];
  ///       mySheet.findAndReplace('MySheetName', 'sad', 'happy', first: 3);
  ///
  ///In the above example it will replace all the occurences of `sad` with `happy` in the cells
  ///
  ///Other `options` are used to `narrow down` the `starting and ending ranges of cells`.
  ///
  int findAndReplace(String sheet, Pattern source, dynamic target,
      {int first = -1,
      int startingRow = -1,
      int endingRow = -1,
      int startingColumn = -1,
      int endingColumn = -1}) {
    int replaceCount = 0;
    if (_sheetMap[sheet] == null) return replaceCount;

    _sheetMap['$sheet']!.findAndReplace(
      source,
      target,
      first: first,
      startingRow: startingRow,
      endingRow: endingRow,
      startingColumn: startingColumn,
      endingColumn: endingColumn,
    );

    return replaceCount;
  }

  ///
  ///Make `sheet` available if it does not exist in `_sheetMap`
  ///
  void _availSheet(String sheet) {
    if (_sheetMap[sheet] == null) {
      _sheetMap[sheet] = Sheet._(this, sheet);
    }
  }

  ///
  ///Updates the contents of `sheet` of the `cellIndex: CellIndex.indexByColumnRow(0, 0);` where indexing starts from 0
  ///
  ///----or---- by `cellIndex: CellIndex.indexByString("A3");`.
  ///
  ///Styling of cell can be done by passing the CellStyle object to `cellStyle`.
  ///
  ///If `sheet` does not exist then it will be automatically created.
  ///
  void updateCell(String sheet, CellIndex cellIndex, CellValue? value,
      {CellStyle? cellStyle}) {
    _availSheet(sheet);

    _sheetMap[sheet]!.updateCell(cellIndex, value, cellStyle: cellStyle);
  }

  ///
  ///Merges the cells starting from `start` to `end`.
  ///
  ///If `custom value` is not defined then it will look for the very first available value in range `start` to `end` by searching row-wise from left to right.
  ///
  ///If `sheet` does not exist then it will be automatically created.
  ///
  void merge(String sheet, CellIndex start, CellIndex end,
      {CellValue? customValue}) {
    _availSheet(sheet);
    _sheetMap[sheet]!.merge(start, end, customValue: customValue);
  }

  ///
  ///returns an Iterable of `cell-Id` for the previously merged cell-Ids.
  ///
  List<String> getMergedCells(String sheet) {
    return List<String>.from(
        _sheetMap[sheet] != null ? _sheetMap[sheet]!.spannedItems : <String>[]);
  }

  ///
  ///unMerge the merged cells.
  ///
  ///       var sheet = 'DesiredSheet';
  ///       List<String> spannedCells = excel.getMergedCells(sheet);
  ///       var cellToUnMerge = "A1:A2";
  ///       excel.unMerge(sheet, cellToUnMerge);
  ///
  void unMerge(String sheet, String unmergeCells) {
    if (_sheetMap[sheet] != null) {
      _sheetMap[sheet]!.unMerge(unmergeCells);
    }
  }

  ///
  ///Internal function taking care of adding the `sheetName` to the `mergeChangeLook` List
  ///So that merging function will be only called on `sheetNames of mergeChangeLook`
  ///
  set _mergeChangeLookup(String value) {
    if (!_mergeChangeLook.contains(value)) {
      _mergeChangeLook.add(value);
      //_mergeChanges = true;
    }
  }

  set _rtlChangeLookup(String value) {
    if (!_rtlChangeLook.contains(value)) {
      _rtlChangeLook.add(value);
      _rtlChanges = true;
    }
  }
}



================================================
FILE: lib/src/number_format/num_format.dart
================================================
part of excel;

Map<V, K> _createInverseMap<K, V>(Map<K, V> map) {
  final inverse = <V, K>{};
  for (var entry in map.entries) {
    assert(!inverse.containsKey(entry.value), 'map values are not unique');
    inverse[entry.value] = entry.key;
  }
  return inverse;
}

class NumFormatMaintainer {
  static const int _firstCustomFmtId = 164;
  int _nextFmtId = _firstCustomFmtId;
  Map<int, NumFormat> _map = {..._standardNumFormats};
  Map<NumFormat, int> _inverseMap = _createInverseMap(_standardNumFormats);

  void add(int numFmtId, CustomNumFormat format) {
    if (_map.containsKey(numFmtId)) {
      throw Exception('numFmtId $numFmtId already exists');
    }
    if (numFmtId < _firstCustomFmtId) {
      throw Exception(
          'invalid numFmtId $numFmtId, custom numFmtId must be $_firstCustomFmtId or greater');
    }
    _map[numFmtId] = format;
    _inverseMap[format] = numFmtId;
    if (numFmtId >= _nextFmtId) {
      _nextFmtId = numFmtId + 1;
    }
  }

  int findOrAdd(CustomNumFormat format) {
    var fmtId = _inverseMap[format];
    if (fmtId != null) {
      return fmtId;
    }
    fmtId = _nextFmtId;
    _nextFmtId++;
    _map[fmtId] = format;
    return fmtId;
  }

  void clear() {
    _nextFmtId = _firstCustomFmtId;
    _map = {..._standardNumFormats};
    _inverseMap = _createInverseMap(_standardNumFormats);
  }

  NumFormat? getByNumFmtId(int numFmtId) {
    return _map[numFmtId];
  }
}

sealed class NumFormat {
  final String formatCode;

  static const defaultNumeric = standard_1;
  static const defaultFloat = standard_2;
  static const defaultBool = standard_0;
  static const defaultDate = standard_14;
  static const defaultTime = standard_20;
  static const defaultDateTime = standard_22;

  static const standard_0 =
      StandardNumericNumFormat._(numFmtId: 0, formatCode: 'General');
  static const standard_1 =
      StandardNumericNumFormat._(numFmtId: 1, formatCode: "0");
  static const standard_2 =
      StandardNumericNumFormat._(numFmtId: 2, formatCode: "0.00");
  static const standard_3 =
      StandardNumericNumFormat._(numFmtId: 3, formatCode: "#,##0");
  static const standard_4 =
      StandardNumericNumFormat._(numFmtId: 4, formatCode: "#,##0.00");
  static const standard_9 =
      StandardNumericNumFormat._(numFmtId: 9, formatCode: "0%");
  static const standard_10 =
      StandardNumericNumFormat._(numFmtId: 10, formatCode: "0.00%");
  static const standard_11 =
      StandardNumericNumFormat._(numFmtId: 11, formatCode: "0.00E+00");
  static const standard_12 =
      StandardNumericNumFormat._(numFmtId: 12, formatCode: "# ?/?");
  static const standard_13 =
      StandardNumericNumFormat._(numFmtId: 13, formatCode: "# ??/??");
  static const standard_14 =
      StandardDateTimeNumFormat._(numFmtId: 14, formatCode: "mm-dd-yy");
  static const standard_15 =
      StandardDateTimeNumFormat._(numFmtId: 15, formatCode: "d-mmm-yy");
  static const standard_16 =
      StandardDateTimeNumFormat._(numFmtId: 16, formatCode: "d-mmm");
  static const standard_17 =
      StandardDateTimeNumFormat._(numFmtId: 17, formatCode: "mmm-yy");
  static const standard_18 =
      StandardTimeNumFormat._(numFmtId: 18, formatCode: "h:mm AM/PM");
  static const standard_19 =
      StandardTimeNumFormat._(numFmtId: 19, formatCode: "h:mm:ss AM/PM");
  static const standard_20 =
      StandardTimeNumFormat._(numFmtId: 20, formatCode: "h:mm");
  static const standard_21 =
      StandardTimeNumFormat._(numFmtId: 21, formatCode: "h:mm:dd");
  static const standard_22 =
      StandardDateTimeNumFormat._(numFmtId: 22, formatCode: "m/d/yy h:mm");
  static const standard_37 =
      StandardNumericNumFormat._(numFmtId: 37, formatCode: "#,##0 ;(#,##0)");
  static const standard_38 = StandardNumericNumFormat._(
      numFmtId: 38, formatCode: "#,##0 ;[Red](#,##0)");
  static const standard_39 = StandardNumericNumFormat._(
      numFmtId: 39, formatCode: "#,##0.00;(#,##0.00)");
  static const standard_40 = StandardNumericNumFormat._(
      numFmtId: 40, formatCode: "#,##0.00;[Red](#,#)");
  static const standard_45 =
      StandardTimeNumFormat._(numFmtId: 45, formatCode: "mm:ss");
  static const standard_46 =
      StandardTimeNumFormat._(numFmtId: 46, formatCode: "[h]:mm:ss");
  static const standard_47 =
      StandardTimeNumFormat._(numFmtId: 47, formatCode: "mmss.0");
  static const standard_48 =
      StandardNumericNumFormat._(numFmtId: 48, formatCode: "##0.0");
  static const standard_49 =
      StandardNumericNumFormat._(numFmtId: 49, formatCode: "@");

  const NumFormat({
    required this.formatCode,
  });

  static CustomNumFormat custom({
    required String formatCode,
  }) {
    if (formatCode == 'General') {
      return CustomNumericNumFormat(formatCode: 'General');
    }

    //const dateParts = ['m', 'mm', 'mmm', 'mmmm', 'mmmmm', 'd', 'dd', 'ddd', 'yy', 'yyyy'];
    //const timeParts = ['h', 'hh', 'm', 'mm', 's', 'ss', 'AM/PM'];

    /// mm appears in dateParts and timeParts, about this from the microsoft website:
    /// > If you use "m" immediately after the "h" or "hh" code or immediately before
    /// > the "ss" code, Excel displays minutes instead of the month.

    /// a very rudamentary check if we're talking date/time/numeric
    /// https://support.microsoft.com/en-us/office/format-numbers-as-dates-or-times-418bd3fe-0577-47c8-8caa-b4d30c528309
    /// or: https://www.ablebits.com/office-addins-blog/custom-excel-number-format/
    /// about dates: https://www.ablebits.com/office-addins-blog/change-date-format-excel/#custom-date-format
    /// about times: https://www.ablebits.com/office-addins-blog/excel-time-format/#custom
    /// [Green]#,##0.00\ \X\X"POSITIV";[Red]\-#\ "Negativ"\.##0.00

    if (_formatCodeLooksLikeDateTime(formatCode)) {
      return CustomDateTimeNumFormat(formatCode: formatCode);
    } else {
      return CustomNumericNumFormat(formatCode: formatCode);
    }
  }

  CellValue read(String v);

  @override
  int get hashCode => Object.hash(runtimeType, formatCode);

  @override
  operator ==(Object other) =>
      other.runtimeType == runtimeType &&
      (other as NumFormat).formatCode == formatCode;

  bool accepts(CellValue? value);

  static NumFormat defaultFor(CellValue? value) => switch (value) {
        null || FormulaCellValue() || TextCellValue() => NumFormat.standard_0,
        IntCellValue() => NumFormat.defaultNumeric,
        DoubleCellValue() => NumFormat.defaultFloat,
        DateCellValue() => NumFormat.defaultDate,
        BoolCellValue() => NumFormat.defaultBool,
        TimeCellValue() => NumFormat.defaultTime,
        DateTimeCellValue() => NumFormat.defaultDateTime,
      };
}

const Map<int, NumFormat> _standardNumFormats = {
  0: NumFormat.standard_0,
  1: NumFormat.standard_1,
  2: NumFormat.standard_2,
  3: NumFormat.standard_3,
  4: NumFormat.standard_4,
  9: NumFormat.standard_9,
  10: NumFormat.standard_10,
  11: NumFormat.standard_11,
  12: NumFormat.standard_12,
  13: NumFormat.standard_13,
  14: NumFormat.standard_14,
  15: NumFormat.standard_15,
  16: NumFormat.standard_16,
  17: NumFormat.standard_17,
  18: NumFormat.standard_18,
  19: NumFormat.standard_19,
  20: NumFormat.standard_20,
  21: NumFormat.standard_21,
  22: NumFormat.standard_22,
  37: NumFormat.standard_37,
  38: NumFormat.standard_38,
  39: NumFormat.standard_39,
  40: NumFormat.standard_40,
  45: NumFormat.standard_45,
  46: NumFormat.standard_46,
  47: NumFormat.standard_47,
  48: NumFormat.standard_48,
  49: NumFormat.standard_49,
};

bool _formatCodeLooksLikeDateTime(String formatCode) {
  // for comparison, remove any character that is quoted or escaped
  var inEscape = false;
  var inQuotes = false;
  for (var i = 0; i < formatCode.length; ++i) {
    final c = formatCode[i];
    if (inEscape) {
      inEscape = false;
      continue;
    } else if (c == '\\') {
      inEscape = true;
      continue;
    }
    if (inQuotes) {
      if (c == '"') {
        inQuotes = false;
      }
      continue;
    } else if (c == '"') {
      inQuotes = true;
      continue;
    }

    switch (c) {
      case 'y':
      case 'm':
      case 'd':
      case 'h':
      case 's':
        return true;
      case ';':
        // separator only exists for decimal formats
        return false;
      default:
        break;
    }
  }
  return false;
}

sealed class StandardNumFormat implements NumFormat {
  int get numFmtId;
}

sealed class CustomNumFormat implements NumFormat {
  String get formatCode;
}

sealed class NumericNumFormat extends NumFormat {
  const NumericNumFormat({
    required super.formatCode,
  });

  @override
  CellValue read(String v) {
    // check if scientific notation e.g. 1E-3
    final eIdx = v.indexOf('E');
    final decimalSeparatorIdx = v.indexOf('.');

    if (decimalSeparatorIdx == -1 && eIdx == -1) {
      return IntCellValue(int.parse(v));
    }

    // also read .0 (or even .00) as an int
    bool noActualDecimalPlaces = true;
    for (var idx = decimalSeparatorIdx + 1; idx < v.length; ++idx) {
      if (v[idx] != '0') {
        noActualDecimalPlaces = false;
        break;
      }
    }
    if (noActualDecimalPlaces) {
      return IntCellValue(int.parse(v.substring(0, decimalSeparatorIdx)));
    }

    return DoubleCellValue(double.parse(v));
  }

  String writeDouble(DoubleCellValue value) {
    return value.value.toString();
  }

  String writeInt(IntCellValue value) {
    return value.value.toString();
  }
}

class StandardNumericNumFormat extends NumericNumFormat
    implements StandardNumFormat {
  @override
  final int numFmtId;

  const StandardNumericNumFormat._({
    required this.numFmtId,
    required super.formatCode,
  });

  @override
  bool accepts(CellValue? value) => switch (value) {
        null => true,
        FormulaCellValue() => true,
        IntCellValue() => true,
        TextCellValue() => numFmtId == 0,
        BoolCellValue() => true,
        DoubleCellValue() => true,
        DateCellValue() => false,
        TimeCellValue() => false,
        DateTimeCellValue() => false,
      };

  @override
  String toString() {
    return 'StandardNumericNumFormat($numFmtId, "$formatCode")';
  }
}

class CustomNumericNumFormat extends NumericNumFormat
    implements CustomNumFormat {
  const CustomNumericNumFormat({
    required super.formatCode,
  });

  @override
  bool accepts(CellValue? value) => switch (value) {
        null => true,
        FormulaCellValue() => true,
        IntCellValue() => true,
        TextCellValue() => false,
        BoolCellValue() => true,
        DoubleCellValue() => true,
        DateCellValue() => false,
        TimeCellValue() => false,
        DateTimeCellValue() => false,
      };

  @override
  String toString() {
    return 'CustomNumericNumFormat("$formatCode")';
  }
}

sealed class DateTimeNumFormat extends NumFormat {
  const DateTimeNumFormat({
    required super.formatCode,
  });

  @override
  CellValue read(String v) {
    if (v == '0') {
      return const TimeCellValue(
        hour: 0,
        minute: 0,
        second: 0,
        millisecond: 0,
        microsecond: 0,
      );
    }
    final value = num.parse(v);
    if (value < 1) {
      return TimeCellValue.fromFractionOfDay(value);
    }
    var delta = value * 24 * 3600 * 1000;
    var dateOffset = DateTime.utc(1899, 12, 30);
    final utcDate = dateOffset.add(Duration(milliseconds: delta.round()));
    if (!v.contains('.') || v.endsWith('.0')) {
      return DateCellValue.fromDateTime(utcDate);
    } else {
      return DateTimeCellValue.fromDateTime(utcDate);
    }
  }

  String writeDate(DateCellValue value) {
    var dateOffset = DateTime.utc(1899, 12, 30);
    final delta = value.asDateTimeUtc().difference(dateOffset);
    final dayFractions = delta.inMilliseconds.toDouble() / (1000 * 3600 * 24);
    return dayFractions.toString();
  }

  String writeDateTime(DateTimeCellValue value) {
    var dateOffset = DateTime.utc(1899, 12, 30);
    final delta = value.asDateTimeUtc().difference(dateOffset);
    final dayFractions = delta.inMilliseconds.toDouble() / (1000 * 3600 * 24);
    return dayFractions.toString();
  }

  @override
  bool accepts(CellValue? value) => switch (value) {
        null => true,
        FormulaCellValue() => true,
        IntCellValue() => false,
        TextCellValue() => false,
        BoolCellValue() => false,
        DoubleCellValue() => false,
        DateCellValue() => true,
        DateTimeCellValue() => true,
        TimeCellValue() => false,
      };
}

class StandardDateTimeNumFormat extends DateTimeNumFormat
    implements StandardNumFormat {
  final int numFmtId;

  const StandardDateTimeNumFormat._({
    required this.numFmtId,
    required super.formatCode,
  });

  @override
  String toString() {
    return 'StandardDateTimeNumFormat($numFmtId, "$formatCode")';
  }
}

class CustomDateTimeNumFormat extends DateTimeNumFormat
    implements CustomNumFormat {
  const CustomDateTimeNumFormat({
    required super.formatCode,
  });

  @override
  String toString() {
    return 'CustomDateTimeNumFormat("$formatCode")';
  }
}

sealed class TimeNumFormat extends NumFormat {
  const TimeNumFormat({
    required super.formatCode,
  });

  @override
  CellValue read(String v) {
    if (v == '0') {
      return const TimeCellValue(
        hour: 0,
        minute: 0,
        second: 0,
        millisecond: 0,
        microsecond: 0,
      );
    }
    var value = num.parse(v);
    if (value < 1) {
      var delta = value * 24 * 3600 * 1000;
      final time = Duration(milliseconds: delta.round());
      final date = DateTime.utc(0).add(time);
      return TimeCellValue(
        hour: date.hour,
        minute: date.minute,
        second: date.second,
        millisecond: date.millisecond,
        microsecond: date.microsecond,
      );
    }
    var delta = value * 24 * 3600 * 1000;
    var dateOffset = DateTime.utc(1899, 12, 30);
    final utcDate = dateOffset.add(Duration(milliseconds: delta.round()));
    if (!v.contains('.') || v.endsWith('.0')) {
      return DateCellValue(
        year: utcDate.year,
        month: utcDate.month,
        day: utcDate.day,
      );
    } else {
      return DateTimeCellValue(
        year: utcDate.year,
        month: utcDate.month,
        day: utcDate.day,
        hour: utcDate.hour,
        minute: utcDate.minute,
        second: utcDate.second,
        millisecond: utcDate.millisecond,
        microsecond: utcDate.microsecond,
      );
    }
  }

  String writeTime(TimeCellValue value) {
    final fractionOfDay =
        value.asDuration().inMilliseconds.toDouble() / (1000 * 3600 * 24);
    return fractionOfDay.toString();
  }

  @override
  bool accepts(CellValue? value) => switch (value) {
        null => true,
        FormulaCellValue() => true,
        IntCellValue() => false,
        TextCellValue() => false,
        BoolCellValue() => false,
        DoubleCellValue() => false,
        DateCellValue() => false,
        DateTimeCellValue() => false,
        TimeCellValue() => true,
      };
}

class StandardTimeNumFormat extends TimeNumFormat implements StandardNumFormat {
  final int numFmtId;

  const StandardTimeNumFormat._({
    required this.numFmtId,
    required super.formatCode,
  });

  @override
  String toString() {
    return 'StandardTimeNumFormat($numFmtId, "$formatCode")';
  }
}

class CustomTimeNumFormat extends TimeNumFormat implements CustomNumFormat {
  const CustomTimeNumFormat({
    required super.formatCode,
  });

  @override
  String toString() {
    return 'CustomTimeNumFormat("$formatCode")';
  }
}



================================================
FILE: lib/src/parser/parse.dart
================================================
part of excel;

class Parser {
  final Excel _excel;
  final List<String> _rId = [];
  final Map<String, String> _worksheetTargets = {};

  Parser._(this._excel);

  void _startParsing() {
    _putContentXml();
    _parseRelations();
    _parseStyles(_excel._stylesTarget);
    _parseSharedStrings();
    _parseContent();
    _parseMergedCells();
  }

  void _normalizeTable(Sheet sheet) {
    if (sheet._maxRows == 0 || sheet._maxColumns == 0) {
      sheet._sheetData.clear();
    }
    sheet._countRowsAndColumns();
  }

  void _putContentXml() {
    var file = _excel._archive.findFile("[Content_Types].xml");

    if (file == null) {
      _damagedExcel();
    }
    file!.decompress();
    _excel._xmlFiles["[Content_Types].xml"] =
        XmlDocument.parse(utf8.decode(file.content));
  }

  void _parseRelations() {
    var relations = _excel._archive.findFile('xl/_rels/workbook.xml.rels');
    if (relations != null) {
      relations.decompress();
      var document = XmlDocument.parse(utf8.decode(relations.content));
      _excel._xmlFiles['xl/_rels/workbook.xml.rels'] = document;

      document.findAllElements('Relationship').forEach((node) {
        String? id = node.getAttribute('Id');
        String? target = node.getAttribute('Target');
        if (target != null) {
          switch (node.getAttribute('Type')) {
            case _relationshipsStyles:
              _excel._stylesTarget = target;
              break;
            case _relationshipsWorksheet:
              if (id != null) _worksheetTargets[id] = target;
              break;
            case _relationshipsSharedStrings:
              _excel._sharedStringsTarget = target;
              break;
          }
        }
        if (id != null && !_rId.contains(id)) {
          _rId.add(id);
        }
      });
    } else {
      _damagedExcel();
    }
  }

  void _parseSharedStrings() {
    var sharedStrings =
        _excel._archive.findFile(_excel._absSharedStringsTarget);
    if (sharedStrings == null) {
      _excel._sharedStringsTarget = 'sharedStrings.xml';

      /// Running it with false will collect all the `rid` and will
      /// help us to get the available rid to assign it to `sharedStrings.xml` back
      _parseContent(run: false);

      if (_excel._xmlFiles.containsKey("xl/_rels/workbook.xml.rels")) {
        int rIdNumber = _getAvailableRid();

        _excel._xmlFiles["xl/_rels/workbook.xml.rels"]
            ?.findAllElements('Relationships')
            .first
            .children
            .add(XmlElement(
              XmlName('Relationship'),
              <XmlAttribute>[
                XmlAttribute(XmlName('Id'), 'rId$rIdNumber'),
                XmlAttribute(XmlName('Type'),
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'),
                XmlAttribute(XmlName('Target'), 'sharedStrings.xml')
              ],
            ));
        if (!_rId.contains('rId$rIdNumber')) {
          _rId.add('rId$rIdNumber');
        }
        String content =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        bool contain = true;

        _excel._xmlFiles["[Content_Types].xml"]
            ?.findAllElements('Override')
            .forEach((node) {
          var value = node.getAttribute('ContentType');
          if (value == content) {
            contain = false;
          }
        });
        if (contain) {
          _excel._xmlFiles["[Content_Types].xml"]
              ?.findAllElements('Types')
              .first
              .children
              .add(XmlElement(
                XmlName('Override'),
                <XmlAttribute>[
                  XmlAttribute(XmlName('PartName'), '/xl/sharedStrings.xml'),
                  XmlAttribute(XmlName('ContentType'), content),
                ],
              ));
        }
      }

      var content = utf8.encode(
          "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"/>");
      _excel._archive.addFile(
          ArchiveFile("xl/sharedStrings.xml", content.length, content));
      sharedStrings = _excel._archive.findFile("xl/sharedStrings.xml");
    }
    sharedStrings!.decompress();
    var document = XmlDocument.parse(utf8.decode(sharedStrings.content));
    _excel._xmlFiles["xl/${_excel._sharedStringsTarget}"] = document;

    document.findAllElements('si').forEach((node) {
      _parseSharedString(node);
    });
  }

  void _parseSharedString(XmlElement node) {
    final sharedString = SharedString(node: node);
    _excel._sharedStrings.add(sharedString, sharedString.stringValue);
  }

  void _parseContent({bool run = true}) {
    var workbook = _excel._archive.findFile('xl/workbook.xml');
    if (workbook == null) {
      _damagedExcel();
    }
    workbook!.decompress();
    var document = XmlDocument.parse(utf8.decode(workbook.content));
    _excel._xmlFiles["xl/workbook.xml"] = document;

    document.findAllElements('sheet').forEach((node) {
      if (run) {
        _parseTable(node);
      } else {
        var rid = node.getAttribute('r:id');
        if (rid != null && !_rId.contains(rid)) {
          _rId.add(rid);
        }
      }
    });
  }

  /// Parses and processes merged cells within the spreadsheet.
  ///
  /// This method identifies merged cell regions in each sheet of the spreadsheet
  /// and handles them accordingly. It removes all cells within a merged cell region
  /// except for the top-left cell, preserving its content.
  void _parseMergedCells() {
    Map spannedCells = <String, List<String>>{};
    _excel._sheets.forEach((sheetName, node) {
      _excel._availSheet(sheetName);
      XmlElement sheetDataNode = node as XmlElement;
      List spanList = <String>[];
      final sheet = _excel._sheetMap[sheetName]!;

      final worksheetNode = sheetDataNode.parent;
      worksheetNode!.findAllElements('mergeCell').forEach((element) {
        String? ref = element.getAttribute('ref');
        if (ref != null && ref.contains(':') && ref.split(':').length == 2) {
          if (!sheet._spannedItems.contains(ref)) {
            sheet._spannedItems.add(ref);
          }

          String startCell = ref.split(':')[0], endCell = ref.split(':')[1];

          if (!spanList.contains(startCell)) {
            spanList.add(startCell);
          }
          spannedCells[sheetName] = spanList;

          CellIndex startIndex = CellIndex.indexByString(startCell),
              endIndex = CellIndex.indexByString(endCell);
          _Span spanObj = _Span.fromCellIndex(
            start: startIndex,
            end: endIndex,
          );
          if (!sheet._spanList.contains(spanObj)) {
            sheet._spanList.add(spanObj);

            _deleteAllButTopLeftCellsOfSpanObj(spanObj, sheet);
          }
          _excel._mergeChangeLookup = sheetName;
        }
      });
    });
  }

  /// Deletes all cells within the span of the given [_Span] object
  /// except for the top-left cell.
  ///
  /// This method is used internally by [_parseMergedCells] to remove
  /// cells within merged cell regions.
  ///
  /// Parameters:
  ///   - [spanObj]: The span object representing the merged cell region.
  ///   - [sheet]: The sheet object from which cells are to be removed.
  void _deleteAllButTopLeftCellsOfSpanObj(_Span spanObj, Sheet sheet) {
    final columnSpanStart = spanObj.columnSpanStart;
    final columnSpanEnd = spanObj.columnSpanEnd;
    final rowSpanStart = spanObj.rowSpanStart;
    final rowSpanEnd = spanObj.rowSpanEnd;

    for (var columnI = columnSpanStart; columnI <= columnSpanEnd; columnI++) {
      for (var rowI = rowSpanStart; rowI <= rowSpanEnd; rowI++) {
        bool isTopLeftCellThatShouldNotBeDeleted =
            columnI == columnSpanStart && rowI == rowSpanStart;

        if (isTopLeftCellThatShouldNotBeDeleted) {
          continue;
        }
        sheet._removeCell(rowI, columnI);
      }
    }
  }

  // Reading the styles from the excel file.
  void _parseStyles(String _stylesTarget) {
    var styles = _excel._archive.findFile('xl/$_stylesTarget');
    if (styles != null) {
      styles.decompress();
      var document = XmlDocument.parse(utf8.decode(styles.content));
      _excel._xmlFiles['xl/$_stylesTarget'] = document;

      _excel._fontStyleList = <_FontStyle>[];
      _excel._patternFill = <String>[];
      _excel._cellStyleList = <CellStyle>[];
      _excel._borderSetList = <_BorderSet>[];

      Iterable<XmlElement> fontList = document.findAllElements('font');

      document.findAllElements('patternFill').forEach((node) {
        String patternType = node.getAttribute('patternType') ?? '', rgb;
        if (node.children.isNotEmpty) {
          node.findElements('fgColor').forEach((child) {
            rgb = child.getAttribute('rgb') ?? '';
            _excel._patternFill.add(rgb);
          });
        } else {
          _excel._patternFill.add(patternType);
        }
      });

      document.findAllElements('border').forEach((node) {
        final diagonalUp = !['0', 'false', null]
            .contains(node.getAttribute('diagonalUp')?.trim());
        final diagonalDown = !['0', 'false', null]
            .contains(node.getAttribute('diagonalDown')?.trim());

        const List<String> borderElementNamesList = [
          'left',
          'right',
          'top',
          'bottom',
          'diagonal'
        ];
        Map<String, Border> borderElements = {};
        for (var elementName in borderElementNamesList) {
          XmlElement? element;
          try {
            element = node.findElements(elementName).single;
          } on StateError catch (_) {
            // Either there is no element, or there are too many ones.
            // Silently ignore this element.
          }

          final borderStyleAttribute = element?.getAttribute('style')?.trim();
          final borderStyle = borderStyleAttribute != null
              ? getBorderStyleByName(borderStyleAttribute)
              : null;

          String? borderColorHex;
          try {
            final color = element?.findElements('color').single;
            borderColorHex = color?.getAttribute('rgb')?.trim();
          } on StateError catch (_) {}

          borderElements[elementName] = Border(
              borderStyle: borderStyle,
              borderColorHex: borderColorHex?.excelColor);
        }

        final borderSet = _BorderSet(
          leftBorder: borderElements['left']!,
          rightBorder: borderElements['right']!,
          topBorder: borderElements['top']!,
          bottomBorder: borderElements['bottom']!,
          diagonalBorder: borderElements['diagonal']!,
          diagonalBorderDown: diagonalDown,
          diagonalBorderUp: diagonalUp,
        );
        _excel._borderSetList.add(borderSet);
      });

      document.findAllElements('numFmts').forEach((node1) {
        node1.findAllElements('numFmt').forEach((node) {
          final numFmtId = int.parse(node.getAttribute('numFmtId')!);
          final formatCode = node.getAttribute('formatCode')!;
          if (numFmtId >= 164) {
            _excel._numFormats
                .add(numFmtId, NumFormat.custom(formatCode: formatCode));
          }
        });
      });

      document.findAllElements('cellXfs').forEach((node1) {
        node1.findAllElements('xf').forEach((node) {
          final numFmtId = _getFontIndex(node, 'numFmtId');
          _excel._numFmtIds.add(numFmtId);

          String fontColor = ExcelColor.black.colorHex,
              backgroundColor = ExcelColor.none.colorHex;
          String? fontFamily;
          FontScheme fontScheme = FontScheme.Unset;
          _BorderSet? borderSet;

          int fontSize = 12;
          bool isBold = false, isItalic = false;
          Underline underline = Underline.None;
          HorizontalAlign horizontalAlign = HorizontalAlign.Left;
          VerticalAlign verticalAlign = VerticalAlign.Bottom;
          TextWrapping? textWrapping;
          int rotation = 0;
          int fontId = _getFontIndex(node, 'fontId');
          _FontStyle _fontStyle = _FontStyle();

          /// checking for other font values
          if (fontId < fontList.length) {
            XmlElement font = fontList.elementAt(fontId);

            /// Checking for font Size.
            var _clr = _nodeChildren(font, 'color', attribute: 'rgb');
            if (_clr != null && !(_clr is bool)) {
              fontColor = _clr.toString();
            }

            /// Checking for font Size.
            String? _size = _nodeChildren(font, 'sz', attribute: 'val');
            if (_size != null) {
              fontSize = double.parse(_size).round();
            }

            /// Checking for bold
            var _bold = _nodeChildren(font, 'b');
            if (_bold != null && _bold is bool && _bold) {
              isBold = true;
            }

            /// Checking for italic
            var _italic = _nodeChildren(font, 'i');
            if (_italic != null && _italic) {
              isItalic = true;
            }

            /// Checking for double underline
            var _underline = _nodeChildren(font, 'u', attribute: 'val');
            if (_underline != null) {
              underline = Underline.Double;
            }

            /// Checking for single underline
            var _singleUnderline = _nodeChildren(font, 'u');
            if (_singleUnderline != null) {
              underline = Underline.Single;
            }

            /// Checking for font Family
            var _family = _nodeChildren(font, 'name', attribute: 'val');
            if (_family != null && _family != true) {
              fontFamily = _family;
            }

            /// Checking for font Scheme
            var _scheme = _nodeChildren(font, 'scheme', attribute: 'val');
            if (_scheme != null) {
              fontScheme =
                  _scheme == "major" ? FontScheme.Major : FontScheme.Minor;
            }

            _fontStyle.isBold = isBold;
            _fontStyle.isItalic = isItalic;
            _fontStyle.fontSize = fontSize;
            _fontStyle.fontFamily = fontFamily;
            _fontStyle.fontScheme = fontScheme;
            _fontStyle._fontColorHex = fontColor.excelColor;
          }

          /// If `-1` is returned then it indicates that `_fontStyle` is not present in the `_fontStyleList`
          if (_fontStyleIndex(_excel._fontStyleList, _fontStyle) == -1) {
            _excel._fontStyleList.add(_fontStyle);
          }

          int fillId = _getFontIndex(node, 'fillId');
          if (fillId < _excel._patternFill.length) {
            backgroundColor = _excel._patternFill[fillId];
          }

          int borderId = _getFontIndex(node, 'borderId');
          if (borderId < _excel._borderSetList.length) {
            borderSet = _excel._borderSetList[borderId];
          }

          if (node.children.isNotEmpty) {
            node.findElements('alignment').forEach((child) {
              if (_getFontIndex(child, 'wrapText') == 1) {
                textWrapping = TextWrapping.WrapText;
              } else if (_getFontIndex(child, 'shrinkToFit') == 1) {
                textWrapping = TextWrapping.Clip;
              }

              var vertical = node.getAttribute('vertical');
              if (vertical != null) {
                if (vertical.toString() == 'top') {
                  verticalAlign = VerticalAlign.Top;
                } else if (vertical.toString() == 'center') {
                  verticalAlign = VerticalAlign.Center;
                }
              }

              var horizontal = node.getAttribute('horizontal');
              if (horizontal != null) {
                if (horizontal.toString() == 'center') {
                  horizontalAlign = HorizontalAlign.Center;
                } else if (horizontal.toString() == 'right') {
                  horizontalAlign = HorizontalAlign.Right;
                }
              }

              var rotationString = node.getAttribute('textRotation');
              if (rotationString != null) {
                rotation = (double.tryParse(rotationString) ?? 0.0).floor();
              }
            });
          }

          var numFormat = _excel._numFormats.getByNumFmtId(numFmtId);
          if (numFormat == null) {
            assert(false, 'missing numFmt for $numFmtId');
            numFormat = NumFormat.standard_0;
          }

          CellStyle cellStyle = CellStyle(
            fontColorHex: fontColor.excelColor,
            fontFamily: fontFamily,
            fontSize: fontSize,
            bold: isBold,
            italic: isItalic,
            underline: underline,
            backgroundColorHex:
                backgroundColor == 'none' || backgroundColor.isEmpty
                    ? ExcelColor.none
                    : backgroundColor.excelColor,
            horizontalAlign: horizontalAlign,
            verticalAlign: verticalAlign,
            textWrapping: textWrapping,
            rotation: rotation,
            leftBorder: borderSet?.leftBorder,
            rightBorder: borderSet?.rightBorder,
            topBorder: borderSet?.topBorder,
            bottomBorder: borderSet?.bottomBorder,
            diagonalBorder: borderSet?.diagonalBorder,
            diagonalBorderUp: borderSet?.diagonalBorderUp ?? false,
            diagonalBorderDown: borderSet?.diagonalBorderDown ?? false,
            numberFormat: numFormat,
          );

          _excel._cellStyleList.add(cellStyle);
        });
      });
    } else {
      _damagedExcel(text: 'styles');
    }
  }

  dynamic _nodeChildren(XmlElement node, String child, {var attribute}) {
    Iterable<XmlElement> ele = node.findElements(child);
    if (ele.isNotEmpty) {
      if (attribute != null) {
        var attr = ele.first.getAttribute(attribute);
        if (attr != null) {
          return attr;
        }
        return null; // pretending that attribute is not found so sending null.
      }
      return true; // mocking to be found the children in case of bold and italic.
    }
    return null; // pretending that the node's children is not having specified child.
  }

  int _getFontIndex(var node, String text) {
    String? applyFont = node.getAttribute(text)?.trim();
    if (applyFont != null) {
      try {
        return int.parse(applyFont.toString());
      } catch (e) {
        if (applyFont.toLowerCase() == 'true') {
          return 1;
        }
      }
    }
    return 0;
  }

  void _parseTable(XmlElement node) {
    var name = node.getAttribute('name')!;
    var target = _worksheetTargets[node.getAttribute('r:id')];

    if (_excel._sheetMap['$name'] == null) {
      _excel._sheetMap['$name'] = Sheet._(_excel, '$name');
    }

    Sheet sheetObject = _excel._sheetMap['$name']!;

    var file = _excel._archive.findFile('xl/$target');
    file!.decompress();

    var content = XmlDocument.parse(utf8.decode(file.content));
    var worksheet = content.findElements('worksheet').first;

    ///
    /// check for right to left view
    ///
    var sheetView = worksheet.findAllElements('sheetView').toList();
    if (sheetView.isNotEmpty) {
      var sheetViewNode = sheetView.first;
      var rtl = sheetViewNode.getAttribute('rightToLeft');
      sheetObject.isRTL = rtl != null && rtl == '1';
    }
    var sheet = worksheet.findElements('sheetData').first;

    _findRows(sheet).forEach((child) {
      _parseRow(child, sheetObject, name);
    });

    _parseHeaderFooter(worksheet, sheetObject);
    _parseColWidthsRowHeights(worksheet, sheetObject);

    _excel._sheets[name] = sheet;

    _excel._xmlFiles['xl/$target'] = content;
    _excel._xmlSheetId[name] = 'xl/$target';

    _normalizeTable(sheetObject);
  }

  _parseRow(XmlElement node, Sheet sheetObject, String name) {
    var rowIndex = (_getRowNumber(node) ?? -1) - 1;
    if (rowIndex < 0) {
      return;
    }

    _findCells(node).forEach((child) {
      _parseCell(child, sheetObject, rowIndex, name);
    });
  }

  void _parseCell(
      XmlElement node, Sheet sheetObject, int rowIndex, String name) {
    int? columnIndex = _getCellNumber(node);
    if (columnIndex == null) {
      return;
    }

    var s1 = node.getAttribute('s');
    int s = 0;
    if (s1 != null) {
      try {
        s = int.parse(s1.toString());
      } catch (_) {}

      String rC = node.getAttribute('r').toString();

      if (_excel._cellStyleReferenced[name] == null) {
        _excel._cellStyleReferenced[name] = {rC: s};
      } else {
        _excel._cellStyleReferenced[name]![rC] = s;
      }
    }

    CellValue? value;
    String? type = node.getAttribute('t');

    switch (type) {
      // sharedString
      case 's':
        final sharedString = _excel._sharedStrings
            .value(int.parse(_parseValue(node.findElements('v').first)));
        value = TextCellValue.span(sharedString!.textSpan);
        break;
      // boolean
      case 'b':
        value = BoolCellValue(_parseValue(node.findElements('v').first) == '1');
        break;
      // error
      case 'e':
      // formula
      case 'str':
        value = FormulaCellValue(_parseValue(node.findElements('v').first));
        break;
      // inline string
      case 'inlineStr':
        // <c r='B2' t='inlineStr'>
        // <is><t>Dartonico</t></is>
        // </c>
        value = TextCellValue(_parseValue(node.findAllElements('t').first));
        break;
      // number
      case 'n':
      default:
        var formulaNode = node.findElements('f');
        if (formulaNode.isNotEmpty) {
          value = FormulaCellValue(_parseValue(formulaNode.first).toString());
        } else {
          final vNode = node.findElements('v').firstOrNull;
          if (vNode == null) {
            value = null;
          } else if (s1 != null) {
            final v = _parseValue(vNode);
            var numFmtId = _excel._numFmtIds[s];
            final numFormat = _excel._numFormats.getByNumFmtId(numFmtId);
            if (numFormat == null) {
              assert(
                  false, 'found no number format spec for numFmtId $numFmtId');
              value = NumFormat.defaultNumeric.read(v);
            } else {
              value = numFormat.read(v);
            }
          } else {
            final v = _parseValue(vNode);
            value = NumFormat.defaultNumeric.read(v);
          }
        }
    }

    sheetObject.updateCell(
      CellIndex.indexByColumnRow(columnIndex: columnIndex, rowIndex: rowIndex),
      value,
      cellStyle: _excel._cellStyleList[s],
    );
  }

  static String _parseValue(XmlElement node) {
    var buffer = StringBuffer();

    node.children.forEach((child) {
      if (child is XmlText) {
        buffer.write(_normalizeNewLine(child.value));
      }
    });

    return buffer.toString();
  }

  int _getAvailableRid() {
    _rId.sort((a, b) {
      return int.parse(a.substring(3)).compareTo(int.parse(b.substring(3)));
    });

    List<String> got = List<String>.from(_rId.last.split(''));
    got.removeWhere((item) {
      return !'0123456789'.split('').contains(item);
    });
    return int.parse(got.join().toString()) + 1;
  }

  ///Uses the [newSheet] as the name of the sheet and also adds it to the [ xl/worksheets/ ] directory
  ///
  ///Creates the sheet with name `newSheet` as file output and then adds it to the archive directory.
  ///
  ///
  void _createSheet(String newSheet) {
    /*
    List<XmlNode> list = _excel._xmlFiles['xl/workbook.xml']
        .findAllElements('sheets')
        .first
        .children;
    if (list.isEmpty) {
      throw ArgumentError('');
    } */

    int _sheetId = -1;
    List<int> sheetIdList = <int>[];

    _excel._xmlFiles['xl/workbook.xml']
        ?.findAllElements('sheet')
        .forEach((sheetIdNode) {
      var sheetId = sheetIdNode.getAttribute('sheetId');
      if (sheetId != null) {
        int t = int.parse(sheetId.toString());
        if (!sheetIdList.contains(t)) {
          sheetIdList.add(t);
        }
      } else {
        _damagedExcel(text: 'Corrupted Sheet Indexing');
      }
    });

    sheetIdList.sort();

    for (int i = 0; i < sheetIdList.length; i++) {
      if ((i + 1) != sheetIdList[i]) {
        _sheetId = i + 1;
        break;
      }
    }
    if (_sheetId == -1) {
      if (sheetIdList.isEmpty) {
        _sheetId = 1;
      } else {
        _sheetId = sheetIdList.length + 1;
      }
    }

    int sheetNumber = _sheetId;
    int ridNumber = _getAvailableRid();

    _excel._xmlFiles['xl/_rels/workbook.xml.rels']
        ?.findAllElements('Relationships')
        .first
        .children
        .add(XmlElement(XmlName('Relationship'), <XmlAttribute>[
          XmlAttribute(XmlName('Id'), 'rId$ridNumber'),
          XmlAttribute(XmlName('Type'), '$_relationships/worksheet'),
          XmlAttribute(XmlName('Target'), 'worksheets/sheet$sheetNumber.xml'),
        ]));

    if (!_rId.contains('rId$ridNumber')) {
      _rId.add('rId$ridNumber');
    }

    _excel._xmlFiles['xl/workbook.xml']
        ?.findAllElements('sheets')
        .first
        .children
        .add(XmlElement(
          XmlName('sheet'),
          <XmlAttribute>[
            XmlAttribute(XmlName('state'), 'visible'),
            XmlAttribute(XmlName('name'), newSheet),
            XmlAttribute(XmlName('sheetId'), '$sheetNumber'),
            XmlAttribute(XmlName('r:id'), 'rId$ridNumber')
          ],
        ));

    _worksheetTargets['rId$ridNumber'] = 'worksheets/sheet$sheetNumber.xml';

    var content = utf8.encode(
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac xr xr2 xr3\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\"> <dimension ref=\"A1\"/> <sheetViews> <sheetView workbookViewId=\"0\"/> </sheetViews> <sheetData/> <pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/> </worksheet>");

    _excel._archive.addFile(ArchiveFile(
        'xl/worksheets/sheet$sheetNumber.xml', content.length, content));
    var _newSheet =
        _excel._archive.findFile('xl/worksheets/sheet$sheetNumber.xml');

    _newSheet!.decompress();
    var document = XmlDocument.parse(utf8.decode(_newSheet.content));
    _excel._xmlFiles['xl/worksheets/sheet$sheetNumber.xml'] = document;
    _excel._xmlSheetId[newSheet] = 'xl/worksheets/sheet$sheetNumber.xml';

    _excel._xmlFiles['[Content_Types].xml']
        ?.findAllElements('Types')
        .first
        .children
        .add(XmlElement(
          XmlName('Override'),
          <XmlAttribute>[
            XmlAttribute(XmlName('ContentType'),
                'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
            XmlAttribute(
                XmlName('PartName'), '/xl/worksheets/sheet$sheetNumber.xml'),
          ],
        ));
    if (_excel._xmlFiles['xl/workbook.xml'] != null) {
      _parseTable(
          _excel._xmlFiles['xl/workbook.xml']!.findAllElements('sheet').last);
    }
  }

  void _parseHeaderFooter(XmlElement worksheet, Sheet sheetObject) {
    final results = worksheet.findAllElements("headerFooter");
    if (results.isEmpty) return;

    final headerFooterElement = results.first;

    sheetObject.headerFooter = HeaderFooter.fromXmlElement(headerFooterElement);
  }

  void _parseColWidthsRowHeights(XmlElement worksheet, Sheet sheetObject) {
    /* parse default column width and default row height
      example XML content
      <sheetFormatPr baseColWidth="10" defaultColWidth="26.33203125" defaultRowHeight="13" x14ac:dyDescent="0.15" />
    */
    Iterable<XmlElement> results;
    results = worksheet.findAllElements("sheetFormatPr");
    if (results.isNotEmpty) {
      results.forEach((element) {
        double? defaultColWidth;
        double? defaultRowHeight;
        // default column width
        String? widthAttribute = element.getAttribute("defaultColWidth");
        if (widthAttribute != null) {
          defaultColWidth = double.tryParse(widthAttribute);
        }
        // default row height
        String? rowHeightAttribute = element.getAttribute("defaultRowHeight");
        if (rowHeightAttribute != null) {
          defaultRowHeight = double.tryParse(rowHeightAttribute);
        }

        // both values valid ?
        if (defaultColWidth != null && defaultRowHeight != null) {
          sheetObject._defaultColumnWidth = defaultColWidth;
          sheetObject._defaultRowHeight = defaultRowHeight;
        }
      });
    }

    /* parse custom column height
      example XML content
      <col min="2" max="2" width="71.83203125" customWidth="1"/>, 
      <col min="4" max="4" width="26.5" customWidth="1"/>, 
      <col min="6" max="6" width="31.33203125" customWidth="1"/>
    */
    results = worksheet.findAllElements("col");
    if (results.isNotEmpty) {
      results.forEach((element) {
        String? colAttribute =
            element.getAttribute("min"); // i think min refers to the column
        String? widthAttribute = element.getAttribute("width");
        if (colAttribute != null && widthAttribute != null) {
          int? col = int.tryParse(colAttribute);
          double? width = double.tryParse(widthAttribute);
          if (col != null && width != null) {
            col -= 1; // first col in _columnWidths is index 0
            if (col >= 0) {
              sheetObject._columnWidths[col] = width;
            }
          }
        }
      });
    }

    /* parse custom row height
      example XML content
      <row r="1" spans="1:2" ht="44" customHeight="1" x14ac:dyDescent="0.15">
    */
    results = worksheet.findAllElements("row");
    if (results.isNotEmpty) {
      results.forEach((element) {
        String? rowAttribute =
            element.getAttribute("r"); // i think min refers to the column
        String? heightAttribute = element.getAttribute("ht");
        if (rowAttribute != null && heightAttribute != null) {
          int? row = int.tryParse(rowAttribute);
          double? height = double.tryParse(heightAttribute);
          if (row != null && height != null) {
            row -= 1; // first col in _rowHeights is index 0
            if (row >= 0) {
              sheetObject._rowHeights[row] = height;
            }
          }
        }
      });
    }
  }
}



================================================
FILE: lib/src/save/save_file.dart
================================================
part of excel;

class Save {
  final Excel _excel;
  final Map<String, ArchiveFile> _archiveFiles = {};
  final List<CellStyle> _innerCellStyle = [];
  final Parser parser;

  Save._(this._excel, this.parser);

  void _addNewColumn(XmlElement columns, int min, int max, double width) {
    columns.children.add(XmlElement(XmlName('col'), [
      XmlAttribute(XmlName('min'), (min + 1).toString()),
      XmlAttribute(XmlName('max'), (max + 1).toString()),
      XmlAttribute(XmlName('width'), width.toStringAsFixed(2)),
      XmlAttribute(XmlName('bestFit'), "1"),
      XmlAttribute(XmlName('customWidth'), "1"),
    ], []));
  }

  double _calcAutoFitColumnWidth(Sheet sheet, int column) {
    var maxNumOfCharacters = 0;
    sheet._sheetData.forEach((key, value) {
      if (value.containsKey(column) &&
          value[column]!.value is! FormulaCellValue) {
        maxNumOfCharacters =
            max(value[column]!.value.toString().length, maxNumOfCharacters);
      }
    });

    return ((maxNumOfCharacters * 7.0 + 9.0) / 7.0 * 256).truncate() / 256;
  }

  /*   XmlElement _replaceCell(String sheet, XmlElement row, XmlElement lastCell,
      int columnIndex, int rowIndex, CellValue? value) {
    var index = lastCell == null ? 0 : row.children.indexOf(lastCell);
    var cell = _createCell(sheet, columnIndex, rowIndex, value);
    row.children
      ..removeAt(index)
      ..insert(index, cell);
    return cell;
  } */

  // Manage value's type
  XmlElement _createCell(String sheet, int columnIndex, int rowIndex,
      CellValue? value, NumFormat? numberFormat) {
    SharedString? sharedString;
    if (value is TextCellValue) {
      sharedString = _excel._sharedStrings.tryFind(value.toString());
      if (sharedString != null) {
        _excel._sharedStrings.add(sharedString, value.toString());
      } else {
        sharedString = _excel._sharedStrings.addFromString(value.toString());
      }
    }

    String rC = getCellId(columnIndex, rowIndex);

    var attributes = <XmlAttribute>[
      XmlAttribute(XmlName('r'), rC),
      if (value is TextCellValue) XmlAttribute(XmlName('t'), 's'),
      if (value is BoolCellValue) XmlAttribute(XmlName('t'), 'b'),
    ];

    final cellStyle =
        _excel._sheetMap[sheet]?._sheetData[rowIndex]?[columnIndex]?.cellStyle;

    if (_excel._styleChanges && cellStyle != null) {
      int upperLevelPos = _checkPosition(_excel._cellStyleList, cellStyle);
      if (upperLevelPos == -1) {
        int lowerLevelPos = _checkPosition(_innerCellStyle, cellStyle);
        if (lowerLevelPos != -1) {
          upperLevelPos = lowerLevelPos + _excel._cellStyleList.length;
        } else {
          upperLevelPos = 0;
        }
      }
      attributes.insert(
        1,
        XmlAttribute(XmlName('s'), '$upperLevelPos'),
      );
    } else if (_excel._cellStyleReferenced.containsKey(sheet) &&
        _excel._cellStyleReferenced[sheet]!.containsKey(rC)) {
      attributes.insert(
        1,
        XmlAttribute(
            XmlName('s'), '${_excel._cellStyleReferenced[sheet]![rC]}'),
      );
    }

    // TODO track & write the numFmts/numFmt to styles.xml if used
    final List<XmlElement> children;
    switch (value) {
      case null:
        children = [];
      case FormulaCellValue():
        children = [
          XmlElement(XmlName('f'), [], [XmlText(value.formula)]),
          XmlElement(XmlName('v'), [], [XmlText('')]),
        ];
      case IntCellValue():
        final String v = switch (numberFormat) {
          NumericNumFormat() => numberFormat.writeInt(value),
          _ => throw Exception(
              '$numberFormat does not work for ${value.runtimeType}'),
        };
        children = [
          XmlElement(XmlName('v'), [], [XmlText(v)]),
        ];
      case DoubleCellValue():
        final String v = switch (numberFormat) {
          NumericNumFormat() => numberFormat.writeDouble(value),
          _ => throw Exception(
              '$numberFormat does not work for ${value.runtimeType}'),
        };
        children = [
          XmlElement(XmlName('v'), [], [XmlText(v)]),
        ];
      case DateTimeCellValue():
        final String v = switch (numberFormat) {
          DateTimeNumFormat() => numberFormat.writeDateTime(value),
          _ => throw Exception(
              '$numberFormat does not work for ${value.runtimeType}'),
        };
        children = [
          XmlElement(XmlName('v'), [], [XmlText(v)]),
        ];
      case DateCellValue():
        final String v = switch (numberFormat) {
          DateTimeNumFormat() => numberFormat.writeDate(value),
          _ => throw Exception(
              '$numberFormat does not work for ${value.runtimeType}'),
        };
        children = [
          XmlElement(XmlName('v'), [], [XmlText(v)]),
        ];
      case TimeCellValue():
        final String v = switch (numberFormat) {
          TimeNumFormat() => numberFormat.writeTime(value),
          _ => throw Exception(
              '$numberFormat does not work for ${value.runtimeType}'),
        };
        children = [
          XmlElement(XmlName('v'), [], [XmlText(v)]),
        ];
      case TextCellValue():
        children = [
          XmlElement(XmlName('v'), [], [
            XmlText(_excel._sharedStrings.indexOf(sharedString!).toString())
          ]),
        ];
      case BoolCellValue():
        children = [
          XmlElement(XmlName('v'), [], [XmlText(value.value ? '1' : '0')]),
        ];
    }

    return XmlElement(XmlName('c'), attributes, children);
  }

  /// Create a new row in the sheet.
  XmlElement _createNewRow(XmlElement table, int rowIndex, double? height) {
    var row = XmlElement(XmlName('row'), [
      XmlAttribute(XmlName('r'), (rowIndex + 1).toString()),
      if (height != null)
        XmlAttribute(XmlName('ht'), height.toStringAsFixed(2)),
      if (height != null) XmlAttribute(XmlName('customHeight'), '1'),
    ], []);
    table.children.add(row);
    return row;
  }

  /// Writing Font Color in [xl/styles.xml] from the Cells of the sheets.

  void _processStylesFile() {
    _innerCellStyle.clear();
    List<String> innerPatternFill = <String>[];
    List<_FontStyle> innerFontStyle = <_FontStyle>[];
    List<_BorderSet> innerBorderSet = <_BorderSet>[];

    _excel._sheetMap.forEach((sheetName, sheetObject) {
      sheetObject._sheetData.forEach((_, columnMap) {
        columnMap.forEach((_, dataObject) {
          if (dataObject.cellStyle != null) {
            int pos = _checkPosition(_innerCellStyle, dataObject.cellStyle!);
            if (pos == -1) {
              _innerCellStyle.add(dataObject.cellStyle!);
            }
          }
        });
      });
    });

    _innerCellStyle.forEach((cellStyle) {
      _FontStyle _fs = _FontStyle(
          bold: cellStyle.isBold,
          italic: cellStyle.isItalic,
          fontColorHex: cellStyle.fontColor,
          underline: cellStyle.underline,
          fontSize: cellStyle.fontSize,
          fontFamily: cellStyle.fontFamily,
          fontScheme: cellStyle.fontScheme);

      /// If `-1` is returned then it indicates that `_fontStyle` is not present in the `_fs`
      if (_fontStyleIndex(_excel._fontStyleList, _fs) == -1 &&
          _fontStyleIndex(innerFontStyle, _fs) == -1) {
        innerFontStyle.add(_fs);
      }

      /// Filling the inner usable extra list of background color
      String backgroundColor = cellStyle.backgroundColor.colorHex;
      if (!_excel._patternFill.contains(backgroundColor) &&
          !innerPatternFill.contains(backgroundColor)) {
        innerPatternFill.add(backgroundColor);
      }

      final _bs = _createBorderSetFromCellStyle(cellStyle);
      if (!_excel._borderSetList.contains(_bs) &&
          !innerBorderSet.contains(_bs)) {
        innerBorderSet.add(_bs);
      }
    });

    XmlElement fonts =
        _excel._xmlFiles['xl/styles.xml']!.findAllElements('fonts').first;

    var fontAttribute = fonts.getAttributeNode('count');
    if (fontAttribute != null) {
      fontAttribute.value =
          '${_excel._fontStyleList.length + innerFontStyle.length}';
    } else {
      fonts.attributes.add(XmlAttribute(XmlName('count'),
          '${_excel._fontStyleList.length + innerFontStyle.length}'));
    }

    innerFontStyle.forEach((fontStyleElement) {
      fonts.children.add(XmlElement(XmlName('font'), [], [
        /// putting color
        if (fontStyleElement._fontColorHex != null &&
            fontStyleElement._fontColorHex!.colorHex != "FF000000")
          XmlElement(XmlName('color'), [
            XmlAttribute(
                XmlName('rgb'), fontStyleElement._fontColorHex!.colorHex)
          ], []),

        /// putting bold
        if (fontStyleElement.isBold) XmlElement(XmlName('b'), [], []),

        /// putting italic
        if (fontStyleElement.isItalic) XmlElement(XmlName('i'), [], []),

        /// putting single underline
        if (fontStyleElement.underline != Underline.None &&
            fontStyleElement.underline == Underline.Single)
          XmlElement(XmlName('u'), [], []),

        /// putting double underline
        if (fontStyleElement.underline != Underline.None &&
            fontStyleElement.underline != Underline.Single &&
            fontStyleElement.underline == Underline.Double)
          XmlElement(
              XmlName('u'), [XmlAttribute(XmlName('val'), 'double')], []),

        /// putting fontFamily
        if (fontStyleElement.fontFamily != null &&
            fontStyleElement.fontFamily!.toLowerCase().toString() != 'null' &&
            fontStyleElement.fontFamily != '' &&
            fontStyleElement.fontFamily!.isNotEmpty)
          XmlElement(XmlName('name'), [
            XmlAttribute(XmlName('val'), fontStyleElement.fontFamily.toString())
          ], []),

        /// putting fontScheme
        if (fontStyleElement.fontScheme != FontScheme.Unset)
          XmlElement(XmlName('scheme'), [
            XmlAttribute(
                XmlName('val'),
                switch (fontStyleElement.fontScheme) {
                  FontScheme.Major => "major",
                  _ => "minor"
                })
          ], []),

        /// putting fontSize
        if (fontStyleElement.fontSize != null &&
            fontStyleElement.fontSize.toString().isNotEmpty)
          XmlElement(XmlName('sz'), [
            XmlAttribute(XmlName('val'), fontStyleElement.fontSize.toString())
          ], []),
      ]));
    });

    XmlElement fills =
        _excel._xmlFiles['xl/styles.xml']!.findAllElements('fills').first;

    var fillAttribute = fills.getAttributeNode('count');

    if (fillAttribute != null) {
      fillAttribute.value =
          '${_excel._patternFill.length + innerPatternFill.length}';
    } else {
      fills.attributes.add(XmlAttribute(XmlName('count'),
          '${_excel._patternFill.length + innerPatternFill.length}'));
    }

    innerPatternFill.forEach((color) {
      if (color.length >= 2) {
        if (color.substring(0, 2).toUpperCase() == 'FF') {
          fills.children.add(XmlElement(XmlName('fill'), [], [
            XmlElement(XmlName('patternFill'), [
              XmlAttribute(XmlName('patternType'), 'solid')
            ], [
              XmlElement(XmlName('fgColor'),
                  [XmlAttribute(XmlName('rgb'), color)], []),
              XmlElement(
                  XmlName('bgColor'), [XmlAttribute(XmlName('rgb'), color)], [])
            ])
          ]));
        } else if (color == "none" ||
            color == "gray125" ||
            color == "lightGray") {
          fills.children.add(XmlElement(XmlName('fill'), [], [
            XmlElement(XmlName('patternFill'),
                [XmlAttribute(XmlName('patternType'), color)], [])
          ]));
        }
      } else {
        _damagedExcel(
            text:
                "Corrupted Styles Found. Can't process further, Open up issue in github.");
      }
    });

    XmlElement borders =
        _excel._xmlFiles['xl/styles.xml']!.findAllElements('borders').first;
    var borderAttribute = borders.getAttributeNode('count');

    if (borderAttribute != null) {
      borderAttribute.value =
          '${_excel._borderSetList.length + innerBorderSet.length}';
    } else {
      borders.attributes.add(XmlAttribute(XmlName('count'),
          '${_excel._borderSetList.length + innerBorderSet.length}'));
    }

    innerBorderSet.forEach((border) {
      var borderElement = XmlElement(XmlName('border'));
      if (border.diagonalBorderDown) {
        borderElement.attributes
            .add(XmlAttribute(XmlName('diagonalDown'), '1'));
      }
      if (border.diagonalBorderUp) {
        borderElement.attributes.add(XmlAttribute(XmlName('diagonalUp'), '1'));
      }
      final Map<String, Border> borderMap = {
        'left': border.leftBorder,
        'right': border.rightBorder,
        'top': border.topBorder,
        'bottom': border.bottomBorder,
        'diagonal': border.diagonalBorder,
      };
      for (var key in borderMap.keys) {
        final borderValue = borderMap[key]!;

        final element = XmlElement(XmlName(key));
        final style = borderValue.borderStyle;
        if (style != null) {
          element.attributes.add(XmlAttribute(XmlName('style'), style.style));
        }
        final color = borderValue.borderColorHex;
        if (color != null) {
          element.children.add(XmlElement(
              XmlName('color'), [XmlAttribute(XmlName('rgb'), color)]));
        }
        borderElement.children.add(element);
      }

      borders.children.add(borderElement);
    });

    final styleSheet = _excel._xmlFiles['xl/styles.xml']!;

    XmlElement celx = styleSheet.findAllElements('cellXfs').first;
    var cellAttribute = celx.getAttributeNode('count');

    if (cellAttribute != null) {
      cellAttribute.value =
          '${_excel._cellStyleList.length + _innerCellStyle.length}';
    } else {
      celx.attributes.add(XmlAttribute(XmlName('count'),
          '${_excel._cellStyleList.length + _innerCellStyle.length}'));
    }

    _innerCellStyle.forEach((cellStyle) {
      String backgroundColor = cellStyle.backgroundColor.colorHex;

      _FontStyle _fs = _FontStyle(
          bold: cellStyle.isBold,
          italic: cellStyle.isItalic,
          fontColorHex: cellStyle.fontColor,
          underline: cellStyle.underline,
          fontSize: cellStyle.fontSize,
          fontFamily: cellStyle.fontFamily);

      HorizontalAlign horizontalAlign = cellStyle.horizontalAlignment;
      VerticalAlign verticalAlign = cellStyle.verticalAlignment;
      int rotation = cellStyle.rotation;
      TextWrapping? textWrapping = cellStyle.wrap;
      int backgroundIndex = innerPatternFill.indexOf(backgroundColor),
          fontIndex = _fontStyleIndex(innerFontStyle, _fs);
      _BorderSet _bs = _createBorderSetFromCellStyle(cellStyle);
      int borderIndex = innerBorderSet.indexOf(_bs);

      final numberFormat = cellStyle.numberFormat;
      final int numFmtId = switch (numberFormat) {
        StandardNumFormat() => numberFormat.numFmtId,
        CustomNumFormat() => _excel._numFormats.findOrAdd(numberFormat),
      };

      var attributes = <XmlAttribute>[
        XmlAttribute(XmlName('borderId'),
            '${borderIndex == -1 ? 0 : borderIndex + _excel._borderSetList.length}'),
        XmlAttribute(XmlName('fillId'),
            '${backgroundIndex == -1 ? 0 : backgroundIndex + _excel._patternFill.length}'),
        XmlAttribute(XmlName('fontId'),
            '${fontIndex == -1 ? 0 : fontIndex + _excel._fontStyleList.length}'),
        XmlAttribute(XmlName('numFmtId'), numFmtId.toString()),
        XmlAttribute(XmlName('xfId'), '0'),
      ];

      if ((_excel._patternFill.contains(backgroundColor) ||
              innerPatternFill.contains(backgroundColor)) &&
          backgroundColor != "none" &&
          backgroundColor != "gray125" &&
          backgroundColor.toLowerCase() != "lightgray") {
        attributes.add(XmlAttribute(XmlName('applyFill'), '1'));
      }

      if (_fontStyleIndex(_excel._fontStyleList, _fs) != -1 &&
          _fontStyleIndex(innerFontStyle, _fs) != -1) {
        attributes.add(XmlAttribute(XmlName('applyFont'), '1'));
      }

      var children = <XmlElement>[];

      if (horizontalAlign != HorizontalAlign.Left ||
          textWrapping != null ||
          verticalAlign != VerticalAlign.Bottom ||
          rotation != 0) {
        attributes.add(XmlAttribute(XmlName('applyAlignment'), '1'));
        var childAttributes = <XmlAttribute>[];

        if (textWrapping != null) {
          childAttributes.add(XmlAttribute(
              XmlName(textWrapping == TextWrapping.Clip
                  ? 'shrinkToFit'
                  : 'wrapText'),
              '1'));
        }

        if (verticalAlign != VerticalAlign.Bottom) {
          String ver = verticalAlign == VerticalAlign.Top ? 'top' : 'center';
          childAttributes.add(XmlAttribute(XmlName('vertical'), '$ver'));
        }

        if (horizontalAlign != HorizontalAlign.Left) {
          String hor =
              horizontalAlign == HorizontalAlign.Right ? 'right' : 'center';
          childAttributes.add(XmlAttribute(XmlName('horizontal'), '$hor'));
        }
        if (rotation != 0) {
          childAttributes
              .add(XmlAttribute(XmlName('textRotation'), '$rotation'));
        }

        children.add(XmlElement(XmlName('alignment'), childAttributes, []));
      }

      celx.children.add(XmlElement(XmlName('xf'), attributes, children));
    });

    final customNumberFormats = _excel._numFormats._map.entries
        .map<MapEntry<int, CustomNumFormat>?>((e) {
          final format = e.value;
          if (format is! CustomNumFormat) {
            return null;
          }
          return MapEntry<int, CustomNumFormat>(e.key, format);
        })
        .nonNulls
        .sorted((a, b) => a.key.compareTo(b.key));

    if (customNumberFormats.isNotEmpty) {
      var numFmtsElement = styleSheet
          .findAllElements('numFmts')
          .whereType<XmlElement>()
          .firstOrNull;
      int count;
      if (numFmtsElement == null) {
        numFmtsElement = XmlElement(XmlName('numFmts'));

        ///FIX: if no default numFormats were added in styles.xml - customNumFormats were added in wrong place,
        styleSheet
            .findElements('styleSheet')
            .first
            .children
            .insert(0, numFmtsElement);
        // styleSheet.children.insert(0, numFmtsElement);
      }
      count = int.parse(numFmtsElement.getAttribute('count') ?? '0');

      for (var numFormat in customNumberFormats) {
        final numFmtIdString = numFormat.key.toString();
        final formatCode = numFormat.value.formatCode;
        var numFmtElement = numFmtsElement.children
            .whereType<XmlElement>()
            .firstWhereOrNull((node) =>
                node.name.local == 'numFmt' &&
                node.getAttribute('numFmtId') == numFmtIdString);
        if (numFmtElement == null) {
          numFmtElement = XmlElement(
              XmlName('numFmt'),
              [
                XmlAttribute(XmlName('numFmtId'), numFmtIdString),
                XmlAttribute(XmlName('formatCode'), formatCode),
              ],
              [],
              true);
          numFmtsElement.children.add(numFmtElement);
          count++;
        } else if ((numFmtElement.getAttribute('formatCode') ?? '') !=
            formatCode) {
          numFmtElement.setAttribute('formatCode', formatCode);
        }
      }

      numFmtsElement.setAttribute('count', count.toString());
    }
  }

  List<int>? _save() {
    if (_excel._styleChanges) {
      _processStylesFile();
    }
    _setSheetElements();
    if (_excel._defaultSheet != null) {
      _setDefaultSheet(_excel._defaultSheet);
    }
    _setSharedStrings();

    if (_excel._mergeChanges) {
      _setMerge();
    }

    if (_excel._rtlChanges) {
      _setRTL();
    }

    for (var xmlFile in _excel._xmlFiles.keys) {
      var xml = _excel._xmlFiles[xmlFile].toString();
      var content = utf8.encode(xml);
      _archiveFiles[xmlFile] = ArchiveFile(xmlFile, content.length, content);
    }
    return ZipEncoder().encode(_cloneArchive(_excel._archive, _archiveFiles));
  }

  void _setColumns(Sheet sheetObject, XmlDocument xmlFile) {
    final columnElements = xmlFile.findAllElements('cols');

    if (sheetObject.getColumnWidths.isEmpty &&
        sheetObject.getColumnAutoFits.isEmpty) {
      if (columnElements.isEmpty) {
        return;
      }

      final columns = columnElements.first;
      final worksheet = xmlFile.findAllElements('worksheet').first;
      worksheet.children.remove(columns);
      return;
    }

    if (columnElements.isEmpty) {
      final worksheet = xmlFile.findAllElements('worksheet').first;
      final sheetData = xmlFile.findAllElements('sheetData').first;
      final index = worksheet.children.indexOf(sheetData);

      worksheet.children.insert(index, XmlElement(XmlName('cols'), [], []));
    }

    var columns = columnElements.first;

    if (columns.children.isNotEmpty) {
      columns.children.clear();
    }

    final autoFits = sheetObject.getColumnAutoFits;
    final customWidths = sheetObject.getColumnWidths;

    final columnCount = max(
        autoFits.isEmpty ? 0 : autoFits.keys.reduce(max) + 1,
        customWidths.isEmpty ? 0 : customWidths.keys.reduce(max) + 1);

    List<double> columnWidths = <double>[];

    double defaultColumnWidth =
        sheetObject.defaultColumnWidth ?? _excelDefaultColumnWidth;

    for (var index = 0; index < columnCount; index++) {
      double width = defaultColumnWidth;

      if (autoFits.containsKey(index) && (!customWidths.containsKey(index))) {
        width = _calcAutoFitColumnWidth(sheetObject, index);
      } else {
        if (customWidths.containsKey(index)) {
          width = customWidths[index]!;
        }
      }

      columnWidths.add(width);

      _addNewColumn(columns, index, index, width);
    }
  }

  void _setRows(String sheetName, Sheet sheetObject) {
    final customHeights = sheetObject.getRowHeights;

    for (var rowIndex = 0; rowIndex < sheetObject._maxRows; rowIndex++) {
      double? height;

      if (customHeights.containsKey(rowIndex)) {
        height = customHeights[rowIndex];
      }

      if (sheetObject._sheetData[rowIndex] == null) {
        continue;
      }
      var foundRow = _createNewRow(
          _excel._sheets[sheetName]! as XmlElement, rowIndex, height);
      for (var columnIndex = 0;
          columnIndex < sheetObject._maxColumns;
          columnIndex++) {
        var data = sheetObject._sheetData[rowIndex]![columnIndex];
        if (data == null) {
          continue;
        }
        _updateCell(sheetName, foundRow, columnIndex, rowIndex, data.value,
            data.cellStyle?.numberFormat);
      }
    }
  }

  bool _setDefaultSheet(String? sheetName) {
    if (sheetName == null || _excel._xmlFiles['xl/workbook.xml'] == null) {
      return false;
    }
    List<XmlElement> sheetList =
        _excel._xmlFiles['xl/workbook.xml']!.findAllElements('sheet').toList();
    XmlElement elementFound = XmlElement(XmlName(''));

    int position = -1;
    for (int i = 0; i < sheetList.length; i++) {
      var _sheetName = sheetList[i].getAttribute('name');
      if (_sheetName != null && _sheetName.toString() == sheetName) {
        elementFound = sheetList[i];
        position = i;
        break;
      }
    }

    if (position == -1) {
      return false;
    }
    if (position == 0) {
      return true;
    }

    _excel._xmlFiles['xl/workbook.xml']!
        .findAllElements('sheets')
        .first
        .children
      ..removeAt(position)
      ..insert(0, elementFound);

    String? expectedSheet = _excel._getDefaultSheet();

    return expectedSheet == sheetName;
  }

  void _setHeaderFooter(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    if (sheet == null) return;

    final xmlFile = _excel._xmlFiles[_excel._xmlSheetId[sheetName]];
    if (xmlFile == null) return;

    final sheetXmlElement = xmlFile.findAllElements("worksheet").first;

    final results = sheetXmlElement.findAllElements("headerFooter");
    if (results.isNotEmpty) {
      sheetXmlElement.children.remove(results.first);
    }

    if (sheet.headerFooter == null) return;

    sheetXmlElement.children.add(sheet.headerFooter!.toXmlElement());
  }

  /// Writing the merged cells information into the excel properties files.
  void _setMerge() {
    _selfCorrectSpanMap(_excel);
    _excel._mergeChangeLook.forEach((s) {
      if (_excel._sheetMap[s] != null &&
          _excel._sheetMap[s]!._spanList.isNotEmpty &&
          _excel._xmlSheetId.containsKey(s) &&
          _excel._xmlFiles.containsKey(_excel._xmlSheetId[s])) {
        Iterable<XmlElement>? iterMergeElement = _excel
            ._xmlFiles[_excel._xmlSheetId[s]]
            ?.findAllElements('mergeCells');
        late XmlElement mergeElement;
        if (iterMergeElement?.isNotEmpty ?? false) {
          mergeElement = iterMergeElement!.first;
        } else {
          if ((_excel._xmlFiles[_excel._xmlSheetId[s]]
                      ?.findAllElements('worksheet')
                      .length ??
                  0) >
              0) {
            int index = _excel._xmlFiles[_excel._xmlSheetId[s]]!
                .findAllElements('worksheet')
                .first
                .children
                .indexOf(_excel._xmlFiles[_excel._xmlSheetId[s]]!
                    .findAllElements("sheetData")
                    .first);
            if (index == -1) {
              _damagedExcel();
            }
            _excel._xmlFiles[_excel._xmlSheetId[s]]!
                .findAllElements('worksheet')
                .first
                .children
                .insert(
                    index + 1,
                    XmlElement(XmlName('mergeCells'),
                        [XmlAttribute(XmlName('count'), '0')]));

            mergeElement = _excel._xmlFiles[_excel._xmlSheetId[s]]!
                .findAllElements('mergeCells')
                .first;
          } else {
            _damagedExcel();
          }
        }

        List<String> _spannedItems =
            List<String>.from(_excel._sheetMap[s]!.spannedItems);

        [
          ['count', _spannedItems.length.toString()],
        ].forEach((value) {
          if (mergeElement.getAttributeNode(value[0]) == null) {
            mergeElement.attributes
                .add(XmlAttribute(XmlName(value[0]), value[1]));
          } else {
            mergeElement.getAttributeNode(value[0])!.value = value[1];
          }
        });

        mergeElement.children.clear();

        _spannedItems.forEach((value) {
          mergeElement.children.add(XmlElement(XmlName('mergeCell'),
              [XmlAttribute(XmlName('ref'), '$value')], []));
        });
      }
    });
  }

  // slow implementation
  /*XmlElement _findRowByIndex(XmlElement table, int rowIndex) {
    XmlElement row;
    var rows = _findRows(table);

    var currentIndex = 0;
    for (var currentRow in rows) {
      currentIndex = _getRowNumber(currentRow) - 1;
      if (currentIndex >= rowIndex) {
        row = currentRow;
        break;
      }
    }

    // Create row if required
    if (row == null || currentIndex != rowIndex) {
      row = __insertRow(table, row, rowIndex);
    }

    return row;
  }

  XmlElement _createRow(int rowIndex) {
    return XmlElement(XmlName('row'),
        [XmlAttribute(XmlName('r'), (rowIndex + 1).toString())], []);
  }

  XmlElement __insertRow(XmlElement table, XmlElement lastRow, int rowIndex) {
    var row = _createRow(rowIndex);
    if (lastRow == null) {
      table.children.add(row);
    } else {
      var index = table.children.indexOf(lastRow);
      table.children.insert(index, row);
    }
    return row;
  }*/

  void _setRTL() {
    _excel._rtlChangeLook.forEach((s) {
      var sheetObject = _excel._sheetMap[s];
      if (sheetObject != null &&
          _excel._xmlSheetId.containsKey(s) &&
          _excel._xmlFiles.containsKey(_excel._xmlSheetId[s])) {
        var itrSheetViewsRTLElement = _excel._xmlFiles[_excel._xmlSheetId[s]]
            ?.findAllElements('sheetViews');

        if (itrSheetViewsRTLElement?.isNotEmpty ?? false) {
          var itrSheetViewRTLElement = _excel._xmlFiles[_excel._xmlSheetId[s]]
              ?.findAllElements('sheetView');

          if (itrSheetViewRTLElement?.isNotEmpty ?? false) {
            /// clear all the children of the sheetViews here

            _excel._xmlFiles[_excel._xmlSheetId[s]]
                ?.findAllElements('sheetViews')
                .first
                .children
                .clear();
          }

          _excel._xmlFiles[_excel._xmlSheetId[s]]
              ?.findAllElements('sheetViews')
              .first
              .children
              .add(XmlElement(
                XmlName('sheetView'),
                [
                  if (sheetObject.isRTL)
                    XmlAttribute(XmlName('rightToLeft'), '1'),
                  XmlAttribute(XmlName('workbookViewId'), '0'),
                ],
              ));
        } else {
          _excel._xmlFiles[_excel._xmlSheetId[s]]
              ?.findAllElements('worksheet')
              .first
              .children
              .add(XmlElement(XmlName('sheetViews'), [], [
                XmlElement(
                  XmlName('sheetView'),
                  [
                    if (sheetObject.isRTL)
                      XmlAttribute(XmlName('rightToLeft'), '1'),
                    XmlAttribute(XmlName('workbookViewId'), '0'),
                  ],
                )
              ]));
        }
      }
    });
  }

  /// Writing the value of excel cells into the separate
  /// sharedStrings file so as to minimize the size of excel files.
  void _setSharedStrings() {
    var uniqueCount = 0;
    var count = 0;

    XmlElement shareString = _excel
        ._xmlFiles['xl/${_excel._sharedStringsTarget}']!
        .findAllElements('sst')
        .first;

    shareString.children.clear();

    _excel._sharedStrings._map.forEach((string, ss) {
      uniqueCount += 1;
      count += ss.count;

      shareString.children.add(string.node);
    });

    [
      ['count', '$count'],
      ['uniqueCount', '$uniqueCount']
    ].forEach((value) {
      if (shareString.getAttributeNode(value[0]) == null) {
        shareString.attributes.add(XmlAttribute(XmlName(value[0]), value[1]));
      } else {
        shareString.getAttributeNode(value[0])!.value = value[1];
      }
    });
  }

  /// Writing cell contained text into the excel sheet files.
  void _setSheetElements() {
    _excel._sharedStrings.clear();

    _excel._sheetMap.forEach((sheetName, sheetObject) {
      ///
      /// Create the sheet's xml file if it does not exist.
      if (_excel._sheets[sheetName] == null) {
        parser._createSheet(sheetName);
      }

      /// Clear the previous contents of the sheet if it exists,
      /// in order to reduce the time to find and compare with the sheet rows
      /// and hence just do the work of putting the data only i.e. creating new rows
      if (_excel._sheets[sheetName]?.children.isNotEmpty ?? false) {
        _excel._sheets[sheetName]!.children.clear();
      }

      /// `Above function is important in order to wipe out the old contents of the sheet.`

      XmlDocument? xmlFile = _excel._xmlFiles[_excel._xmlSheetId[sheetName]];
      if (xmlFile == null) return;

      // Set default column width and height for the sheet.
      double? defaultRowHeight = sheetObject.defaultRowHeight;
      double? defaultColumnWidth = sheetObject.defaultColumnWidth;

      XmlElement worksheetElement = xmlFile.findAllElements('worksheet').first;

      XmlElement? sheetFormatPrElement =
          worksheetElement.findElements('sheetFormatPr').isNotEmpty
              ? worksheetElement.findElements('sheetFormatPr').first
              : null;

      if (sheetFormatPrElement != null) {
        sheetFormatPrElement.attributes.clear();

        if (defaultRowHeight == null && defaultColumnWidth == null) {
          worksheetElement.children.remove(sheetFormatPrElement);
        }
      } else if (defaultRowHeight != null || defaultColumnWidth != null) {
        sheetFormatPrElement = XmlElement(XmlName('sheetFormatPr'), [], []);
        worksheetElement.children.insert(0, sheetFormatPrElement);
      }

      if (defaultRowHeight != null) {
        sheetFormatPrElement!.attributes.add(XmlAttribute(
            XmlName('defaultRowHeight'), defaultRowHeight.toStringAsFixed(2)));
      }
      if (defaultColumnWidth != null) {
        sheetFormatPrElement!.attributes.add(XmlAttribute(
            XmlName('defaultColWidth'), defaultColumnWidth.toStringAsFixed(2)));
      }

      _setColumns(sheetObject, xmlFile);

      _setRows(sheetName, sheetObject);

      _setHeaderFooter(sheetName);
    });
  }

  // slow implementation
/*   XmlElement _updateCell(String sheet, XmlElement node, int columnIndex,
      int rowIndex, CellValue? value) {
    XmlElement cell;
    var cells = _findCells(node);

    var currentIndex = 0; // cells could be empty
    for (var currentCell in cells) {
      currentIndex = _getCellNumber(currentCell);
      if (currentIndex >= columnIndex) {
        cell = currentCell;
        break;
      }
    }

    if (cell == null || currentIndex != columnIndex) {
      cell = _insertCell(sheet, node, cell, columnIndex, rowIndex, value);
    } else {
      cell = _replaceCell(sheet, node, cell, columnIndex, rowIndex, value);
    }

    return cell;
  } */
  XmlElement _updateCell(String sheet, XmlElement row, int columnIndex,
      int rowIndex, CellValue? value, NumFormat? numberFormat) {
    var cell = _createCell(sheet, columnIndex, rowIndex, value, numberFormat);
    row.children.add(cell);
    return cell;
  }

  _BorderSet _createBorderSetFromCellStyle(CellStyle cellStyle) => _BorderSet(
        leftBorder: cellStyle.leftBorder,
        rightBorder: cellStyle.rightBorder,
        topBorder: cellStyle.topBorder,
        bottomBorder: cellStyle.bottomBorder,
        diagonalBorder: cellStyle.diagonalBorder,
        diagonalBorderUp: cellStyle.diagonalBorderUp,
        diagonalBorderDown: cellStyle.diagonalBorderDown,
      );
}



================================================
FILE: lib/src/save/self_correct_span.dart
================================================
part of excel;

///Self correct the spanning of rows and columns by checking their cross-sectional relationship between if exists.
_selfCorrectSpanMap(Excel _excel) {
  _excel._mergeChangeLook.forEach((String key) {
    if (_excel._sheetMap[key] != null &&
        _excel._sheetMap[key]!._spanList.isNotEmpty) {
      List<_Span?> spanList =
          List<_Span?>.from(_excel._sheetMap[key]!._spanList);

      for (int i = 0; i < spanList.length; i++) {
        _Span? checkerPos = spanList[i];
        if (checkerPos == null) {
          continue;
        }
        int startRow = checkerPos.rowSpanStart,
            startColumn = checkerPos.columnSpanStart,
            endRow = checkerPos.rowSpanEnd,
            endColumn = checkerPos.columnSpanEnd;

        for (int j = i + 1; j < spanList.length; j++) {
          _Span? spanObj = spanList[j];
          if (spanObj == null) {
            continue;
          }

          final locationChange = _isLocationChangeRequired(
              startColumn, startRow, endColumn, endRow, spanObj);
          if (locationChange.$1) {
            startColumn = locationChange.$2.$1;
            startRow = locationChange.$2.$2;
            endColumn = locationChange.$2.$3;
            endRow = locationChange.$2.$4;
            spanList[j] = null;
          } else {
            final locationChange2 = _isLocationChangeRequired(
                spanObj.columnSpanStart,
                spanObj.rowSpanStart,
                spanObj.columnSpanEnd,
                spanObj.rowSpanEnd,
                checkerPos);

            if (locationChange2.$1) {
              startColumn = locationChange2.$2.$1;
              startRow = locationChange2.$2.$2;
              endColumn = locationChange2.$2.$3;
              endRow = locationChange2.$2.$4;
              spanList[j] = null;
            }
          }
        }
        _Span spanObj1 = _Span(
          rowSpanStart: startRow,
          columnSpanStart: startColumn,
          rowSpanEnd: endRow,
          columnSpanEnd: endColumn,
        );
        spanList[i] = spanObj1;
      }
      _excel._sheetMap[key]!._spanList = List<_Span?>.from(spanList);
      _excel._sheetMap[key]!._cleanUpSpanMap();
    }
  });
}



================================================
FILE: lib/src/sharedStrings/shared_strings.dart
================================================
part of excel;

class _SharedStringsMaintainer {
  final Map<SharedString, _IndexingHolder> _map =
      <SharedString, _IndexingHolder>{};
  final Map<String, SharedString> _mapString = <String, SharedString>{};
  final List<SharedString> _list = <SharedString>[];
  int _index = 0;

  _SharedStringsMaintainer._();

  SharedString? tryFind(String val) {
    return _mapString[val];
  }

  SharedString addFromString(String val) {
    final newSharedString = SharedString(
        node: XmlElement(XmlName('si'), [], [
      XmlElement(XmlName('t'),
          [XmlAttribute(XmlName("space", "xml"), "preserve")], [XmlText(val)]),
    ]));

    add(newSharedString, val);
    return newSharedString;
  }

  void add(SharedString val, String key) {
    _map[val]?.increaseCount();
    _map.putIfAbsent(val, () {
      _mapString[key] = val;
      _list.add(val);
      return _IndexingHolder(_index++);
    });
  }

  int indexOf(SharedString val) {
    return _map[val] != null ? _map[val]!.index : -1;
  }

  SharedString? value(int i) {
    if (i < _list.length) {
      return _list[i];
    } else {
      return null;
    }
  }

  void clear() {
    _index = 0;
    _list.clear();
    _map.clear();
    _mapString.clear();
  }
}

class _IndexingHolder {
  final int index;
  int count;

  _IndexingHolder(this.index, [int _count = 1]) : count = _count;

  void increaseCount() {
    this.count += 1;
  }
}

class SharedString {
  final XmlElement node;
  final int _hashCode;

  SharedString({required this.node}) : _hashCode = node.toString().hashCode;

  @override
  String toString() {
    assert(false,
        'prefer stringValue over SharedString.toString() in development');
    return stringValue;
  }

  TextSpan get textSpan {
    bool getBool(XmlElement element) {
      return bool.tryParse(element.getAttribute('val') ?? '') ?? true;
    }

    int getDouble(XmlElement element) {
      // Should be double
      return double.parse(element.getAttribute('val')!).toInt();
    }

    String? text;
    List<TextSpan>? children;

    /// SharedStringItem
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sharedstringitem?view=openxml-3.0.1
    assert(node.localName == 'si'); //18.4.8 si (String Item)

    for (final child in node.childElements) {
      switch (child.localName) {
        /// Text
        /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.text?view=openxml-3.0.1
        case 't': //18.4.12 t (Text)
          text = (text ?? '') + child.innerText;
          break;

        /// Rich Text Run
        /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.run?view=openxml-3.0.1
        case 'r': //18.4.4 r (Rich Text Run)
          var style = CellStyle();
          for (final runChild in child.childElements) {
            switch (runChild.localName) {
              /// RunProperties
              /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.runproperties?view=openxml-3.0.1
              case 'rPr':
                for (final runProperty in runChild.childElements) {
                  switch (runProperty.localName) {
                    case 'b': //18.8.2 b (Bold)
                      style = style.copyWith(boldVal: getBool(runProperty));
                      break;
                    case 'i': //18.8.26 i (Italic)
                      style = style.copyWith(italicVal: getBool(runProperty));
                      break;
                    case 'u': //18.4.13 u (Underline)
                      style = style.copyWith(
                          underlineVal:
                              runProperty.getAttribute('val') == 'double'
                                  ? Underline.Double
                                  : Underline.Single);
                      break;
                    case 'sz': //18.4.11 sz (Font Size)
                      style =
                          style.copyWith(fontSizeVal: getDouble(runProperty));
                      break;
                    case 'rFont': //18.4.5 rFont (Font)
                      style = style.copyWith(
                          fontFamilyVal: runProperty.getAttribute('val'));
                      break;
                    case 'color': //18.3.1.15 color (Data Bar Color)
                      style = style.copyWith(
                          fontColorHexVal:
                              runProperty.getAttribute('rgb')?.excelColor);
                      break;
                  }
                }
                break;

              /// Text
              case 't': //18.4.12 t (Text)
                if (children == null) children = [];
                children.add(TextSpan(text: runChild.innerText, style: style));
                break;
            }
          }
          break;

        /// Phonetic Run
        /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.phoneticrun?view=openxml-3.0.1
        case 'rPh': //18.4.6 rPh (Phonetic Run)
          break;
      }
    }

    return TextSpan(text: text, children: children);
  }

  String get stringValue {
    var buffer = StringBuffer();
    node.findAllElements('t').forEach((child) {
      if (child.parentElement == null ||
          child.parentElement!.name.local != 'rPh') {
        buffer.write(Parser._parseValue(child));
      }
    });
    return buffer.toString();
  }

  @override
  int get hashCode => _hashCode;

  @override
  operator ==(Object other) {
    return other is SharedString &&
        other.hashCode == _hashCode &&
        other.stringValue == stringValue;
  }

  bool matches(String value) {
    return value.isNotEmpty && value == stringValue;
  }
}

class TextSpan {
  final String? text;
  final List<TextSpan>? children;
  final CellStyle? style;

  const TextSpan({this.children, this.text, this.style});

  @override
  String toString() {
    String r = '';
    if (text != null) r += text!;
    if (children != null) r += children!.join();
    return r;
  }

  @override
  operator ==(Object other) {
    if (identical(this, other)) return true;
    if (other.runtimeType != runtimeType) return false;
    return other is TextSpan &&
        other.text == text &&
        other.style == style &&
        ListEquality().equals(other.children, children);
  }

  @override
  int get hashCode =>
      Object.hash(text, style, Object.hashAll(children ?? const []));
}



================================================
FILE: lib/src/sheet/border_style.dart
================================================
part of excel;

class Border extends Equatable {
  final BorderStyle? borderStyle;
  final String? borderColorHex;

  Border({BorderStyle? borderStyle, ExcelColor? borderColorHex})
      : borderStyle = borderStyle == BorderStyle.None ? null : borderStyle,
        borderColorHex = borderColorHex != null
            ? _isColorAppropriate(borderColorHex.colorHex)
            : null;

  @override
  String toString() {
    return 'Border(borderStyle: $borderStyle, borderColorHex: $borderColorHex)';
  }

  @override
  List<Object?> get props => [
        borderStyle,
        borderColorHex,
      ];
}

class _BorderSet extends Equatable {
  final Border leftBorder;
  final Border rightBorder;
  final Border topBorder;
  final Border bottomBorder;
  final Border diagonalBorder;
  final bool diagonalBorderUp;
  final bool diagonalBorderDown;

  _BorderSet({
    required this.leftBorder,
    required this.rightBorder,
    required this.topBorder,
    required this.bottomBorder,
    required this.diagonalBorder,
    required this.diagonalBorderUp,
    required this.diagonalBorderDown,
  });

  _BorderSet copyWith({
    Border? leftBorder,
    Border? rightBorder,
    Border? topBorder,
    Border? bottomBorder,
    Border? diagonalBorder,
    bool? diagonalBorderUp,
    bool? diagonalBorderDown,
  }) {
    return _BorderSet(
      leftBorder: leftBorder ?? this.leftBorder,
      rightBorder: rightBorder ?? this.rightBorder,
      topBorder: topBorder ?? this.topBorder,
      bottomBorder: bottomBorder ?? this.bottomBorder,
      diagonalBorder: diagonalBorder ?? this.diagonalBorder,
      diagonalBorderUp: diagonalBorderUp ?? this.diagonalBorderUp,
      diagonalBorderDown: diagonalBorderDown ?? this.diagonalBorderDown,
    );
  }

  @override
  List<Object?> get props => [
        leftBorder,
        rightBorder,
        topBorder,
        bottomBorder,
        diagonalBorder,
        diagonalBorderUp,
        diagonalBorderDown,
      ];
}

enum BorderStyle {
  None('none'),
  DashDot('dashDot'),
  DashDotDot('dashDotDot'),
  Dashed('dashed'),
  Dotted('dotted'),
  Double('double'),
  Hair('hair'),
  Medium('medium'),
  MediumDashDot('mediumDashDot'),
  MediumDashDotDot('mediumDashDotDot'),
  MediumDashed('mediumDashed'),
  SlantDashDot('slantDashDot'),
  Thick('thick'),
  Thin('thin');

  final String style;
  const BorderStyle(this.style);
}

BorderStyle? getBorderStyleByName(String name) =>
    BorderStyle.values.firstWhereOrNull((e) =>
        e.toString().toLowerCase() == 'borderstyle.' + name.toLowerCase());



================================================
FILE: lib/src/sheet/cell_index.dart
================================================
part of excel;

class CellIndex extends Equatable {
  CellIndex._({required this.columnIndex, required this.rowIndex});

  ///
  ///```
  ///CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0 ); // A1
  ///CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 1 ); // A2
  ///```
  factory CellIndex.indexByColumnRow(
      {required int columnIndex, required int rowIndex}) {
    return CellIndex._(columnIndex: columnIndex, rowIndex: rowIndex);
  }

  ///
  ///```
  /// CellIndex.indexByString('A1'); // columnIndex: 0, rowIndex: 0
  /// CellIndex.indexByString('A2'); // columnIndex: 0, rowIndex: 1
  ///```
  factory CellIndex.indexByString(String cellIndex) {
    final coords = _cellCoordsFromCellId(cellIndex);
    return CellIndex._(rowIndex: coords.$1, columnIndex: coords.$2);
  }

  /// Avoid using it as it is very process expensive function.
  ///
  /// ```
  /// var cellIndex = CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0 );
  /// var cell = cellIndex.cellId; // A1
  String get cellId {
    return getCellId(this.columnIndex, this.rowIndex);
  }

  final int rowIndex;
  final int columnIndex;

  @override
  List<Object?> get props => [rowIndex, columnIndex];
}



================================================
FILE: lib/src/sheet/cell_style.dart
================================================
part of excel;

/// Styling class for cells
// ignore: must_be_immutable
class CellStyle extends Equatable {
  String _fontColorHex = ExcelColor.black.colorHex;
  String _backgroundColorHex = ExcelColor.none.colorHex;
  String? _fontFamily;
  FontScheme _fontScheme;
  HorizontalAlign _horizontalAlign = HorizontalAlign.Left;
  VerticalAlign _verticalAlign = VerticalAlign.Bottom;
  TextWrapping? _textWrapping;
  bool _bold = false, _italic = false;
  Underline _underline = Underline.None;
  int? _fontSize;
  int _rotation = 0;
  Border _leftBorder;
  Border _rightBorder;
  Border _topBorder;
  Border _bottomBorder;
  Border _diagonalBorder;
  bool _diagonalBorderUp = false;
  bool _diagonalBorderDown = false;
  NumFormat numberFormat;

  CellStyle({
    ExcelColor fontColorHex = ExcelColor.black,
    ExcelColor backgroundColorHex = ExcelColor.none,
    int? fontSize,
    String? fontFamily,
    FontScheme? fontScheme,
    HorizontalAlign horizontalAlign = HorizontalAlign.Left,
    VerticalAlign verticalAlign = VerticalAlign.Bottom,
    TextWrapping? textWrapping,
    bool bold = false,
    Underline underline = Underline.None,
    bool italic = false,
    int rotation = 0,
    Border? leftBorder,
    Border? rightBorder,
    Border? topBorder,
    Border? bottomBorder,
    Border? diagonalBorder,
    bool diagonalBorderUp = false,
    bool diagonalBorderDown = false,
    this.numberFormat = NumFormat.standard_0,
  })  : _textWrapping = textWrapping,
        _bold = bold,
        _fontSize = fontSize,
        _italic = italic,
        _fontFamily = fontFamily,
        _fontScheme = fontScheme ?? FontScheme.Unset,
        _rotation = rotation,
        _fontColorHex = _isColorAppropriate(fontColorHex.colorHex),
        _backgroundColorHex = _isColorAppropriate(backgroundColorHex.colorHex),
        _verticalAlign = verticalAlign,
        _horizontalAlign = horizontalAlign,
        _leftBorder = leftBorder ?? Border(),
        _rightBorder = rightBorder ?? Border(),
        _topBorder = topBorder ?? Border(),
        _bottomBorder = bottomBorder ?? Border(),
        _diagonalBorder = diagonalBorder ?? Border(),
        _diagonalBorderUp = diagonalBorderUp,
        _diagonalBorderDown = diagonalBorderDown;

  CellStyle copyWith({
    ExcelColor? fontColorHexVal,
    ExcelColor? backgroundColorHexVal,
    String? fontFamilyVal,
    FontScheme? fontSchemeVal,
    HorizontalAlign? horizontalAlignVal,
    VerticalAlign? verticalAlignVal,
    TextWrapping? textWrappingVal,
    bool? boldVal,
    bool? italicVal,
    Underline? underlineVal,
    int? fontSizeVal,
    int? rotationVal,
    Border? leftBorderVal,
    Border? rightBorderVal,
    Border? topBorderVal,
    Border? bottomBorderVal,
    Border? diagonalBorderVal,
    bool? diagonalBorderUpVal,
    bool? diagonalBorderDownVal,
    NumFormat? numberFormat,
  }) {
    return CellStyle(
      fontColorHex: fontColorHexVal ?? this._fontColorHex.excelColor,
      backgroundColorHex:
          backgroundColorHexVal ?? this._backgroundColorHex.excelColor,
      fontFamily: fontFamilyVal ?? this._fontFamily,
      fontScheme: fontSchemeVal ?? this._fontScheme,
      horizontalAlign: horizontalAlignVal ?? this._horizontalAlign,
      verticalAlign: verticalAlignVal ?? this._verticalAlign,
      textWrapping: textWrappingVal ?? this._textWrapping,
      bold: boldVal ?? this._bold,
      italic: italicVal ?? this._italic,
      underline: underlineVal ?? this._underline,
      fontSize: fontSizeVal ?? this._fontSize,
      rotation: rotationVal ?? this._rotation,
      leftBorder: leftBorderVal ?? this._leftBorder,
      rightBorder: rightBorderVal ?? this._rightBorder,
      topBorder: topBorderVal ?? this._topBorder,
      bottomBorder: bottomBorderVal ?? this._bottomBorder,
      diagonalBorder: diagonalBorderVal ?? this._diagonalBorder,
      diagonalBorderUp: diagonalBorderUpVal ?? this._diagonalBorderUp,
      diagonalBorderDown: diagonalBorderDownVal ?? this._diagonalBorderDown,
      numberFormat: numberFormat ?? this.numberFormat,
    );
  }

  ///Get Font Color
  ///
  ExcelColor get fontColor {
    return _fontColorHex.excelColor;
  }

  ///Set Font Color
  ///
  set fontColor(ExcelColor fontColorHex) {
    _fontColorHex = _isColorAppropriate(fontColorHex.colorHex);
  }

  ///Get Background Color
  ///
  ExcelColor get backgroundColor {
    return _backgroundColorHex.excelColor;
  }

  ///Set Background Color
  ///
  set backgroundColor(ExcelColor backgroundColorHex) {
    _backgroundColorHex = _isColorAppropriate(backgroundColorHex.colorHex);
  }

  ///Get Horizontal Alignment
  ///
  HorizontalAlign get horizontalAlignment {
    return _horizontalAlign;
  }

  ///Set Horizontal Alignment
  ///
  set horizontalAlignment(HorizontalAlign horizontalAlign) {
    _horizontalAlign = horizontalAlign;
  }

  ///Get Vertical Alignment
  ///
  VerticalAlign get verticalAlignment {
    return _verticalAlign;
  }

  ///Set Vertical Alignment
  ///
  set verticalAlignment(VerticalAlign verticalAlign) {
    _verticalAlign = verticalAlign;
  }

  ///`Get Wrapping`
  ///
  TextWrapping? get wrap {
    return _textWrapping;
  }

  ///`Set Wrapping`
  ///
  set wrap(TextWrapping? textWrapping) {
    _textWrapping = textWrapping;
  }

  ///`Get FontFamily`
  ///
  String? get fontFamily {
    return _fontFamily;
  }

  ///`Set FontFamily`
  ///
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

  ///Get Font Size
  ///
  int? get fontSize {
    return _fontSize;
  }

  ///Set Font Size
  ///
  set fontSize(int? _fs) {
    _fontSize = _fs;
  }

  ///Get Rotation
  ///
  int get rotation {
    return _rotation;
  }

  ///Rotation varies from [90 to -90]
  ///
  set rotation(int _rotate) {
    if (_rotate > 90 || _rotate < -90) {
      _rotate = 0;
    }
    if (_rotate < 0) {
      /// The value is from 0 to -90 so now make it absolute and add it to 90
      ///
      /// -(_rotate) + 90
      _rotate = -(_rotate) + 90;
    }
    _rotation = _rotate;
  }

  ///Get `Underline`
  ///
  Underline get underline {
    return _underline;
  }

  ///set `Underline`
  ///
  set underline(Underline _) {
    _underline = _;
  }

  ///Get `Bold`
  ///
  bool get isBold {
    return _bold;
  }

  ///Set `Bold`
  set isBold(bool bold) {
    _bold = bold;
  }

  ///Get `Italic`
  ///
  bool get isItalic {
    return _italic;
  }

  ///Set `Italic`
  ///
  set isItalic(bool italic) {
    _italic = italic;
  }

  ///Get `LeftBorder`
  ///
  Border get leftBorder {
    return _leftBorder;
  }

  ///Set `LeftBorder`
  ///
  set leftBorder(Border? leftBorder) {
    _leftBorder = leftBorder ?? Border();
  }

  ///Get `RightBorder`
  ///
  Border get rightBorder {
    return _rightBorder;
  }

  ///Set `RightBorder`
  ///
  set rightBorder(Border? rightBorder) {
    _rightBorder = rightBorder ?? Border();
  }

  ///Get `TopBorder`
  ///
  Border get topBorder {
    return _topBorder;
  }

  ///Set `TopBorder`
  ///
  set topBorder(Border? topBorder) {
    _topBorder = topBorder ?? Border();
  }

  ///Get `BottomBorder`
  ///
  Border get bottomBorder {
    return _bottomBorder;
  }

  ///Set `BottomBorder`
  ///
  set bottomBorder(Border? bottomBorder) {
    _bottomBorder = bottomBorder ?? Border();
  }

  ///Get `DiagonalBorder`
  ///
  Border get diagonalBorder {
    return _diagonalBorder;
  }

  ///Set `DiagonalBorder`
  ///
  set diagonalBorder(Border? diagonalBorder) {
    _diagonalBorder = diagonalBorder ?? Border();
  }

  ///Get `DiagonalBorderUp`
  ///
  bool get diagonalBorderUp {
    return _diagonalBorderUp;
  }

  ///Set `DiagonalBorderUp`
  ///
  set diagonalBorderUp(bool diagonalBorderUp) {
    _diagonalBorderUp = diagonalBorderUp;
  }

  ///Get `DiagonalBorderDown`
  ///
  bool get diagonalBorderDown {
    return _diagonalBorderDown;
  }

  ///Set `DiagonalBorderDown`
  ///
  set diagonalBorderDown(bool diagonalBorderDown) {
    _diagonalBorderDown = diagonalBorderDown;
  }

  @override
  List<Object?> get props => [
        _bold,
        _rotation,
        _italic,
        _underline,
        _fontSize,
        _fontFamily,
        _fontScheme,
        _textWrapping,
        _verticalAlign,
        _horizontalAlign,
        _fontColorHex,
        _backgroundColorHex,
        _leftBorder,
        _rightBorder,
        _topBorder,
        _bottomBorder,
        _diagonalBorder,
        _diagonalBorderUp,
        _diagonalBorderDown,
        numberFormat,
      ];
}



================================================
FILE: lib/src/sheet/data_model.dart
================================================
part of excel;

// ignore: must_be_immutable
class Data extends Equatable {
  CellStyle? _cellStyle;
  CellValue? _value;
  Sheet _sheet;
  String _sheetName;
  int _rowIndex;
  int _columnIndex;

  ///
  ///It will clone the object by changing the `this` reference of previous DataObject and putting `new this` reference, with copying the values too
  ///
  Data._clone(Sheet sheet, Data dataObject)
      : this._(
          sheet,
          dataObject._rowIndex,
          dataObject.columnIndex,
          value: dataObject._value,
          cellStyleVal: dataObject._cellStyle,
        );

  ///
  ///Initializes the new `Data Object`
  ///
  Data._(
    Sheet sheet,
    int row,
    int column, {
    CellValue? value,
    NumFormat? numberFormat,
    CellStyle? cellStyleVal,
    bool isFormulaVal = false,
  })  : _sheet = sheet,
        _value = value,
        _cellStyle = cellStyleVal,
        _sheetName = sheet.sheetName,
        _rowIndex = row,
        _columnIndex = column;

  /// returns the newData object when called from Sheet Class
  static Data newData(Sheet sheet, int row, int column) {
    return Data._(sheet, row, column);
  }

  /// returns the row Index
  int get rowIndex {
    return _rowIndex;
  }

  /// returns the column Index
  int get columnIndex {
    return _columnIndex;
  }

  /// returns the sheet-name
  String get sheetName {
    return _sheetName;
  }

  /// returns the string based cellId as A1, A2 or Z5
  CellIndex get cellIndex {
    return CellIndex.indexByColumnRow(
        columnIndex: _columnIndex, rowIndex: _rowIndex);
  }

  /// Helps to set the formula
  ///```
  ///var sheet = excel['Sheet1'];
  ///var cell = sheet.cell(CellIndex.indexByString("E5"));
  ///cell.setFormula('=SUM(1,2)');
  ///```
  void setFormula(String formula) {
    _sheet.updateCell(cellIndex, FormulaCellValue(formula));
  }

  set value(CellValue? val) {
    _sheet.updateCell(cellIndex, val);
  }

  /// returns the value stored in this cell;
  ///
  /// It will return `null` if no value is stored in this cell.
  CellValue? get value => _value;

  /// returns the user-defined CellStyle
  ///
  /// if `no` cellStyle is set then it returns `null`
  CellStyle? get cellStyle {
    return _cellStyle;
  }

  /// sets the user defined CellStyle in this current cell
  set cellStyle(CellStyle? _) {
    _sheet._excel._styleChanges = true;
    _cellStyle = _;
  }

  @override
  List<Object?> get props => [
        _value,
        _columnIndex,
        _rowIndex,
        _cellStyle,
        _sheetName,
      ];
}

sealed class CellValue {
  const CellValue();
}

class FormulaCellValue extends CellValue {
  final String formula;

  const FormulaCellValue(this.formula);

  @override
  String toString() {
    return formula;
  }

  @override
  int get hashCode => Object.hash(runtimeType, formula);

  @override
  operator ==(Object other) {
    return other is FormulaCellValue && other.formula == formula;
  }
}

class IntCellValue extends CellValue {
  final int value;

  const IntCellValue(this.value);

  @override
  String toString() {
    return value.toString();
  }

  @override
  int get hashCode => Object.hash(runtimeType, value);

  @override
  operator ==(Object other) {
    return other is IntCellValue && other.value == value;
  }
}

class DoubleCellValue extends CellValue {
  final double value;

  const DoubleCellValue(this.value);

  @override
  String toString() {
    return value.toString();
  }

  @override
  int get hashCode => Object.hash(runtimeType, value);

  @override
  operator ==(Object other) {
    return other is DoubleCellValue && other.value == value;
  }
}

class DateCellValue extends CellValue {
  final int year;
  final int month;
  final int day;

  const DateCellValue({
    required this.year,
    required this.month,
    required this.day,
  })  : assert(month <= 12 && month >= 1),
        assert(day <= 31 && day >= 1);

  DateCellValue.fromDateTime(DateTime dt)
      : year = dt.year,
        month = dt.month,
        day = dt.day;

  DateTime asDateTimeLocal() {
    return DateTime(year, month, day);
  }

  DateTime asDateTimeUtc() {
    return DateTime.utc(year, month, day);
  }

  @override
  String toString() {
    return asDateTimeUtc().toIso8601String();
  }

  @override
  int get hashCode => Object.hash(runtimeType, year, month, day);

  @override
  operator ==(Object other) {
    return other is DateCellValue &&
        other.year == year &&
        other.month == month &&
        other.day == day;
  }
}

class TextCellValue extends CellValue {
  final TextSpan value;

  TextCellValue(String text) : value = TextSpan(text: text);
  TextCellValue.span(this.value);

  @override
  String toString() {
    return value.toString();
  }

  @override
  int get hashCode => Object.hash(runtimeType, value);

  @override
  operator ==(Object other) {
    return other is TextCellValue && other.value == value;
  }
}

class BoolCellValue extends CellValue {
  final bool value;

  const BoolCellValue(this.value);

  @override
  String toString() {
    return value.toString();
  }

  @override
  int get hashCode => Object.hash(runtimeType, value);

  @override
  operator ==(Object other) {
    return other is BoolCellValue && other.value == value;
  }
}

class TimeCellValue extends CellValue {
  final int hour;
  final int minute;
  final int second;
  final int millisecond;
  final int microsecond;

  const TimeCellValue({
    this.hour = 0,
    this.minute = 0,
    this.second = 0,
    this.millisecond = 0,
    this.microsecond = 0,
  })  : assert(hour >= 0),
        assert(minute <= 60 && minute >= 0),
        assert(second <= 60 && second >= 0),
        assert(millisecond <= 1000 && millisecond >= 0),
        assert(microsecond <= 1000 && microsecond >= 0);

  /// [fractionOfDay]=1.0 is 24 hours, 0.5 is 12 hours and so on.
  factory TimeCellValue.fromFractionOfDay(num fractionOfDay) {
    var duration =
        Duration(milliseconds: (fractionOfDay * 24 * 3600 * 1000).round());
    return TimeCellValue.fromDuration(duration);
  }

  factory TimeCellValue.fromDuration(Duration duration) {
    final someUtcDate = DateTime.utc(0).add(duration);
    return TimeCellValue(
      hour: someUtcDate.hour,
      minute: someUtcDate.minute,
      second: someUtcDate.second,
      millisecond: someUtcDate.millisecond,
      microsecond: someUtcDate.microsecond,
    );
  }

  TimeCellValue.fromTimeOfDateTime(DateTime dt)
      : hour = dt.hour,
        minute = dt.minute,
        second = dt.second,
        millisecond = dt.millisecond,
        microsecond = dt.microsecond;

  Duration asDuration() {
    return Duration(
      hours: hour,
      minutes: minute,
      seconds: second,
      milliseconds: millisecond,
      microseconds: microsecond,
    );
  }

  @override
  String toString() {
    return '${_twoDigits(hour)}:${_twoDigits(minute)}:${_twoDigits(second)}';
  }

  @override
  int get hashCode => Object.hash(
        runtimeType,
        hour,
        minute,
        second,
        millisecond,
        microsecond,
      );

  @override
  operator ==(Object other) {
    return other is TimeCellValue &&
        other.hour == hour &&
        other.minute == minute &&
        other.second == second &&
        other.millisecond == millisecond &&
        other.microsecond == microsecond;
  }
}

/// Excel does not know if this is UTC or not. Use methods [asDateTimeLocal]
/// or [asDateTimeUtc] to get the DateTime object you prefer.
class DateTimeCellValue extends CellValue {
  final int year;
  final int month;
  final int day;
  final int hour;
  final int minute;
  final int second;
  final int millisecond;
  final int microsecond;

  const DateTimeCellValue({
    required this.year,
    required this.month,
    required this.day,
    required this.hour,
    required this.minute,
    this.second = 0,
    this.millisecond = 0,
    this.microsecond = 0,
  })  : assert(month <= 12 && month >= 1),
        assert(day <= 31 && day >= 1),
        assert(hour <= 24 && hour >= 0),
        assert(minute <= 60 && minute >= 0),
        assert(second <= 60 && second >= 0),
        assert(millisecond <= 1000 && millisecond >= 0),
        assert(microsecond <= 1000 && microsecond >= 0);

  DateTimeCellValue.fromDateTime(DateTime date)
      : year = date.year,
        month = date.month,
        day = date.day,
        hour = date.hour,
        minute = date.minute,
        second = date.second,
        millisecond = date.millisecond,
        microsecond = date.microsecond;

  DateTime asDateTimeLocal() {
    return DateTime(
        year, month, day, hour, minute, second, millisecond, microsecond);
  }

  DateTime asDateTimeUtc() {
    return DateTime.utc(
        year, month, day, hour, minute, second, millisecond, microsecond);
  }

  @override
  String toString() {
    return asDateTimeUtc().toIso8601String();
  }

  @override
  int get hashCode => Object.hash(
        runtimeType,
        year,
        month,
        day,
        hour,
        minute,
        second,
        millisecond,
        microsecond,
      );

  @override
  operator ==(Object other) {
    return other is DateTimeCellValue &&
        other.year == year &&
        other.month == month &&
        other.day == day &&
        other.hour == hour &&
        other.minute == minute &&
        other.second == second &&
        other.millisecond == millisecond &&
        other.microsecond == microsecond;
  }
}



================================================
FILE: lib/src/sheet/font_family.dart
================================================
part of excel;

enum FontFamily {
  Al_Bayan_Plain,
  Abadi_MT_Condensed_Light,
  Abadi_MT_Condensed_Extra_Bold,
  Al_Nile,
  Al_Tarikh_Regular,
  American_Typewriter,
  Andale_Mono,
  Angsana_New,
  Apple_Braille_Outline_8_Dot,
  Apple_Chancery,
  Apple_Color_Emoji,
  Apple_Symbols,
  Arial,
  Arial_Hebrew,
  Arial_Hebrew_Scholar,
  Arial_Narrow,
  Arial_Rounded_MT_Bold,
  Arial_Unicode_MS,
  Athelas_Regular,
  Avenir_Book,
  Avenir_Next_Regular,
  Avenir_Next_Condensed_Regular,
  Ayuthaya,
  Baghdad,
  Bangla_MN,
  Bangla_Sangam_MN,
  Baskerville,
  Baskerville_Old_Face,
  Bauhaus_93,
  Beirut,
  Bell_MT,
  Bernard_MT_Condensed,
  Big_Caslon,
  Bodoni_72,
  Bodoni_72_Oldstyle,
  Bodoni_72_Smallcaps,
  Bodoni_Ornaments,
  Book_Antiqua,
  Bookman_Old_Style,
  Bookshelf_Symbol_7,
  Bradley_Hand,
  Braggadocio,
  Britannic_Bold,
  Brush_Script_MT,
  Calibri,
  Calisto_MT,
  Cambria,
  Cambria_Math,
  Candara,
  Century,
  Century_Gothic,
  Century_Schoolbook,
  Chalkboard,
  Chalkboard_SE,
  Chalkduster,
  Charter,
  Cochin,
  Colonna_MT,
  Comic_Sans_MS,
  Consolas,
  Constantia,
  Cooper_Black,
  Copperplate,
  Copperplate_Gothic_Bold,
  Corbel,
  Cordia_New,
  CordiaUPC,
  Corsiva_Hebrew,
  Courier,
  Courier_New,
  Curlz_MT,
  Damascus,
  David,
  DecoType_Naskh,
  Desdemona,
  Devanagari_MT,
  Devanagari_Sangam_MN,
  Didot,
  DIN_Alternate,
  DIN_Condensed,
  Diwan_Kufi,
  Diwan_Thuluth,
  Dubai,
  Edwardian_Script_ITC,
  Engravers_MT,
  Euphemia_UCAS,
  Eurostile,
  Farah,
  Farisi,
  Footlight_MT_Light,
  Franklin_Gothic_Book,
  Franklin_Gothic_Demi,
  Franklin_Gothic_Demi_Cond,
  Franklin_Gothic_Heavy,
  Franklin_Gothic_Medium,
  Franklin_Gothic_Medium_Cond,
  Futura,
  Gabriola,
  Galvji,
  Garamond,
  Gautami,
  Geeza_Pro,
  Geneva,
  Georgia,
  Gill_Sans,
  Gill_Sans_MT,
  Gill_Sans_MT_Condensed,
  Gill_Sans_MT_Ext_Condensed_Bold,
  Gill_Sans_Ultra_Bold,
  Gloucester_MT_Extra_Condensed,
  Goudy_Old_Style,
  Gujarati_MT,
  Gujarati_Sangam_MN,
  Gurmukhi_MN,
  Gurmukhi_MT,
  Gurmukhi_Sangam_MN,
  Haettenschweiler,
  Harrington,
  Helvetica,
  Helvetica_Neue,
  Herculanum,
  Hoefler_Text,
  Impact,
  Imprint_MT_Shadow,
  InaiMathi,
  Iowan_Old_Style,
  ITF_Devanagari,
  ITF_Devanagari_Marathi,
  Kailasa,
  Kannada_MN,
  Kannada_Sangam_MN,
  Kartika,
  Kefa,
  Khmer_MN,
  Khmer_Sangam_MN,
  Kino_MT,
  Kohinoor_Bangla,
  Kohinoor_Devanagari,
  Kohinoor_Gujarati,
  Kohinoor_Telugu,
  Kokonor,
  Lao_MN,
  Lao_Sangam_MN,
  Latha,
  Lucida_Blackletter,
  Lucida_Bright,
  Lucida_Calligraphy,
  Lucida_Console,
  Lucida_Fax,
  Lucida_Grande,
  Lucida_Handwriting,
  Lucida_Sans,
  Lucida_Sans_Typewriter,
  Lucida_Sans_Unicode,
  Luminari,
  Malayalam_MN,
  Malayalam_Sangam_MN,
  Mangal,
  Marion,
  Marker_Felt,
  Marlett,
  Matura_MT_Script_Capitals,
  Menlo,
  Microsoft_New_Tai_Lue,
  Microsoft_Sans_Serif,
  Microsoft_Tai_Le,
  Microsoft_Yi_Baiti,
  Mishafi,
  Mishafi_Gold,
  Mistral,
  Monaco,
  Monotype_Corsiva,
  Monotype_Sorts,
  MS_Reference_Sans_Serif,
  MS_Reference_Specialty,
  Mshtakan,
  MT_Extra,
  Mukta_Mahee,
  Muna,
  Myanmar_MN,
  Myanmar_Sangam_MN,
  Myanmar_Text,
}

///
///
///returns the `Font Family Name`
///
///
String getFontFamily(FontFamily fontFamily) {
  return (fontFamily.toString().replaceAll('FontFamily.', ''))
      .replaceAll('_', ' ');
}



================================================
FILE: lib/src/sheet/font_style.dart
================================================
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

  @override
  List<Object?> get props => [
        _bold,
        _italic,
        _fontSize,
        _underline,
        _fontFamily,
        _fontColorHex,
      ];
}



================================================
FILE: lib/src/sheet/header_footer.dart
================================================
part of excel;

class HeaderFooter {
  bool? alignWithMargins;
  bool? differentFirst;
  bool? differentOddEven;
  bool? scaleWithDoc;

  String? evenFooter;
  String? evenHeader;
  String? firstFooter;
  String? firstHeader;
  String? oddFooter;
  String? oddHeader;

  HeaderFooter({
    this.alignWithMargins,
    this.differentFirst,
    this.differentOddEven,
    this.scaleWithDoc,
    this.evenFooter,
    this.evenHeader,
    this.firstFooter,
    this.firstHeader,
    this.oddFooter,
    this.oddHeader,
  });

  XmlNode toXmlElement() {
    final attributes = <XmlAttribute>[];
    if (alignWithMargins != null) {
      attributes.add(XmlAttribute(
          XmlName("alignWithMargins"), alignWithMargins.toString()));
    }
    if (differentFirst != null) {
      attributes.add(
          XmlAttribute(XmlName("differentFirst"), differentFirst.toString()));
    }
    if (differentOddEven != null) {
      attributes.add(XmlAttribute(
          XmlName("differentOddEven"), differentOddEven.toString()));
    }
    if (scaleWithDoc != null) {
      attributes
          .add(XmlAttribute(XmlName("scaleWithDoc"), scaleWithDoc.toString()));
    }

    final children = <XmlNode>[];
    if (evenHeader != null) {
      children.add(XmlElement(
          XmlName("evenHeader"), [], [XmlText(evenHeader!.simplifyText())]));
    }
    if (evenFooter != null) {
      children.add(XmlElement(
          XmlName("evenFooter"), [], [XmlText(evenFooter!.simplifyText())]));
    }
    if (firstHeader != null) {
      children.add(XmlElement(
          XmlName("firstHeader"), [], [XmlText(firstHeader!.simplifyText())]));
    }
    if (firstFooter != null) {
      children.add(XmlElement(
          XmlName("firstFooter"), [], [XmlText(firstFooter!.simplifyText())]));
    }
    if (oddHeader != null) {
      children.add(XmlElement(
          XmlName("oddHeader"), [], [XmlText(oddHeader!.simplifyText())]));
    }
    if (oddFooter != null) {
      children.add(XmlElement(
          XmlName("oddFooter"), [], [XmlText(oddFooter!.simplifyText())]));
    }

    return XmlElement(XmlName("headerFooter"), attributes, children);
  }

  static HeaderFooter fromXmlElement(XmlElement headerFooterElement) {
    return HeaderFooter(
        alignWithMargins:
            headerFooterElement.getAttribute("alignWithMargins")?.parseBool(),
        differentFirst:
            headerFooterElement.getAttribute("differentFirst")?.parseBool(),
        differentOddEven:
            headerFooterElement.getAttribute("differentOddEven")?.parseBool(),
        scaleWithDoc:
            headerFooterElement.getAttribute("scaleWithDoc")?.parseBool(),
        evenHeader: headerFooterElement.getElement("evenHeader")?.innerText,
        evenFooter: headerFooterElement.getElement("evenFooter")?.innerText,
        firstHeader: headerFooterElement.getElement("firstHeader")?.innerText,
        firstFooter: headerFooterElement.getElement("firstFooter")?.innerText,
        oddFooter: headerFooterElement.getElement("oddFooter")?.innerText,
        oddHeader: headerFooterElement.getElement("oddHeader")?.innerText);
  }
}

extension BoolParsing on String {
  bool parseBool() {
    var value = this.toLowerCase();
    if (value == 'true' || value == '1') {
      return true;
    } else if (value == 'false' || value == '0') {
      return false;
    }

    throw '"$this" can not be parsed to boolean.';
  }

  String simplifyText() {
    String value = this.replaceAll('&amp', '&');
    value = value.replaceAll('amp', '&');
    value = value.replaceAll('&', '&amp;');
    value = value.replaceAll('"', '&quot;');
    return value;
  }
}



================================================
FILE: lib/src/sheet/sheet.dart
================================================
part of excel;

class Sheet {
  final Excel _excel;
  final String _sheet;
  bool _isRTL = false;
  int _maxRows = 0;
  int _maxColumns = 0;
  double? _defaultColumnWidth;
  double? _defaultRowHeight;
  Map<int, double> _columnWidths = {};
  Map<int, double> _rowHeights = {};
  Map<int, bool> _columnAutoFit = {};
  FastList<String> _spannedItems = FastList<String>();
  List<_Span?> _spanList = [];
  Map<int, Map<int, Data>> _sheetData = {};
  HeaderFooter? _headerFooter;

  ///
  /// It will clone the object by changing the `this` reference of previous oldSheetObject and putting `new this` reference, with copying the values too
  ///
  Sheet._clone(Excel excel, String sheetName, Sheet oldSheetObject)
      : this._(excel, sheetName,
            sh: oldSheetObject._sheetData,
            spanL_: oldSheetObject._spanList,
            spanI_: oldSheetObject._spannedItems,
            maxRowsVal: oldSheetObject._maxRows,
            maxColumnsVal: oldSheetObject._maxColumns,
            columnWidthsVal: oldSheetObject._columnWidths,
            rowHeightsVal: oldSheetObject._rowHeights,
            columnAutoFitVal: oldSheetObject._columnAutoFit,
            isRTLVal: oldSheetObject._isRTL,
            headerFooter: oldSheetObject._headerFooter);

  Sheet._(this._excel, this._sheet,
      {Map<int, Map<int, Data>>? sh,
      List<_Span?>? spanL_,
      FastList<String>? spanI_,
      int? maxRowsVal,
      int? maxColumnsVal,
      bool? isRTLVal,
      Map<int, double>? columnWidthsVal,
      Map<int, double>? rowHeightsVal,
      Map<int, bool>? columnAutoFitVal,
      HeaderFooter? headerFooter}) {
    _headerFooter = headerFooter;

    if (spanL_ != null) {
      _spanList = List<_Span?>.from(spanL_);
      _excel._mergeChangeLookup = sheetName;
    }
    if (spanI_ != null) {
      _spannedItems = FastList<String>.from(spanI_);
    }
    if (maxColumnsVal != null) {
      _maxColumns = maxColumnsVal;
    }
    if (maxRowsVal != null) {
      _maxRows = maxRowsVal;
    }
    if (isRTLVal != null) {
      _isRTL = isRTLVal;
      _excel._rtlChangeLookup = sheetName;
    }
    if (columnWidthsVal != null) {
      _columnWidths = Map<int, double>.from(columnWidthsVal);
    }
    if (rowHeightsVal != null) {
      _rowHeights = Map<int, double>.from(rowHeightsVal);
    }
    if (columnAutoFitVal != null) {
      _columnAutoFit = Map<int, bool>.from(columnAutoFitVal);
    }

    /// copy the data objects into a temp folder and then while putting it into `_sheetData` change the data objects references.
    if (sh != null) {
      _sheetData = <int, Map<int, Data>>{};
      Map<int, Map<int, Data>> temp = Map<int, Map<int, Data>>.from(sh);
      temp.forEach((key, value) {
        if (_sheetData[key] == null) {
          _sheetData[key] = <int, Data>{};
        }
        temp[key]!.forEach((key1, oldDataObject) {
          _sheetData[key]![key1] = Data._clone(this, oldDataObject);
        });
      });
    }
    _countRowsAndColumns();
  }

  /// Removes a cell from the specified [rowIndex] and [columnIndex].
  ///
  /// If the specified [rowIndex] or [columnIndex] does not exist,
  /// no action is taken.
  ///
  /// If the removal of the cell results in an empty row, the entire row is removed.
  ///
  /// Parameters:
  ///   - [rowIndex]: The index of the row from which to remove the cell.
  ///   - [columnIndex]: The index of the column from which to remove the cell.
  ///
  /// Example:
  /// ```dart
  /// final sheet = Spreadsheet();
  /// sheet.removeCell(1, 2);
  /// ```
  void _removeCell(int rowIndex, int columnIndex) {
    _sheetData[rowIndex]?.remove(columnIndex);
    final rowIsEmptyAfterRemovalOfCell = _sheetData[rowIndex]?.isEmpty == true;
    if (rowIsEmptyAfterRemovalOfCell) {
      _sheetData.remove(rowIndex);
    }
  }

  ///
  /// returns `true` is this sheet is `right-to-left` other-wise `false`
  ///
  bool get isRTL {
    return _isRTL;
  }

  ///
  /// set sheet-object to `true` for making it `right-to-left` otherwise `false`
  ///
  set isRTL(bool _u) {
    _isRTL = _u;
    _excel._rtlChangeLookup = sheetName;
  }

  ///
  /// returns the `DataObject` at position of `cellIndex`
  ///
  Data cell(CellIndex cellIndex) {
    _checkMaxColumn(cellIndex.columnIndex);
    _checkMaxRow(cellIndex.rowIndex);
    if (cellIndex.columnIndex < 0 || cellIndex.rowIndex < 0) {
      _damagedExcel(
          text:
              '${cellIndex.columnIndex < 0 ? "Column" : "Row"} Index: ${cellIndex.columnIndex < 0 ? cellIndex.columnIndex : cellIndex.rowIndex} Negative index does not exist.');
    }

    /// increasing the row count
    if (_maxRows < (cellIndex.rowIndex + 1)) {
      _maxRows = cellIndex.rowIndex + 1;
    }

    /// increasing the column count
    if (_maxColumns < (cellIndex.columnIndex + 1)) {
      _maxColumns = cellIndex.columnIndex + 1;
    }

    /// checking if the map has been already initialized or not?
    /// if the user has called this class by its own
    /* if (_sheetData == null) {
      _sheetData = Map<int, Map<int, Data>>();
    } */

    /// if the sheetData contains the row then start putting the column
    if (_sheetData[cellIndex.rowIndex] != null) {
      if (_sheetData[cellIndex.rowIndex]![cellIndex.columnIndex] == null) {
        _sheetData[cellIndex.rowIndex]![cellIndex.columnIndex] =
            Data.newData(this, cellIndex.rowIndex, cellIndex.columnIndex);
      }
    } else {
      /// else put the column with map showing.
      _sheetData[cellIndex.rowIndex] = {
        cellIndex.columnIndex:
            Data.newData(this, cellIndex.rowIndex, cellIndex.columnIndex)
      };
    }

    return _sheetData[cellIndex.rowIndex]![cellIndex.columnIndex]!;
  }

  ///
  /// returns `2-D dynamic List` of the sheet elements
  ///
  List<List<Data?>> get rows {
    var _data = <List<Data?>>[];

    if (_sheetData.isEmpty) {
      return _data;
    }

    if (_maxRows > 0 && maxColumns > 0) {
      _data = List.generate(_maxRows, (rowIndex) {
        return List.generate(_maxColumns, (columnIndex) {
          if (_sheetData[rowIndex] != null &&
              _sheetData[rowIndex]![columnIndex] != null) {
            return _sheetData[rowIndex]![columnIndex];
          }
          return null;
        });
      });
    }

    return _data;
  }

  ///
  /// returns `2-D dynamic List` of the sheet cell data in that range.
  ///
  /// Ex. selectRange('D8:H12'); or selectRange('D8');
  ///
  List<List<Data?>?> selectRangeWithString(String range) {
    List<List<Data?>?> _selectedRange = <List<Data?>?>[];
    if (!range.contains(':')) {
      var start = CellIndex.indexByString(range);
      _selectedRange = selectRange(start);
    } else {
      var rangeVars = range.split(':');
      var start = CellIndex.indexByString(rangeVars[0]);
      var end = CellIndex.indexByString(rangeVars[1]);
      _selectedRange = selectRange(start, end: end);
    }
    return _selectedRange;
  }

  ///
  /// returns `2-D dynamic List` of the sheet cell data in that range.
  ///
  List<List<Data?>?> selectRange(CellIndex start, {CellIndex? end}) {
    _checkMaxColumn(start.columnIndex);
    _checkMaxRow(start.rowIndex);
    if (end != null) {
      _checkMaxColumn(end.columnIndex);
      _checkMaxRow(end.rowIndex);
    }

    int _startColumn = start.columnIndex, _startRow = start.rowIndex;
    int? _endColumn = end?.columnIndex, _endRow = end?.rowIndex;

    if (_endColumn != null && _endRow != null) {
      if (_startRow > _endRow) {
        _startRow = end!.rowIndex;
        _endRow = start.rowIndex;
      }
      if (_endColumn < _startColumn) {
        _endColumn = start.columnIndex;
        _startColumn = end!.columnIndex;
      }
    }

    List<List<Data?>?> _selectedRange = <List<Data?>?>[];
    if (_sheetData.isEmpty) {
      return _selectedRange;
    }

    for (var i = _startRow; i <= (_endRow ?? maxRows); i++) {
      var mapData = _sheetData[i];
      if (mapData != null) {
        List<Data?> row = <Data?>[];
        for (var j = _startColumn; j <= (_endColumn ?? maxColumns); j++) {
          row.add(mapData[j]);
        }
        _selectedRange.add(row);
      } else {
        _selectedRange.add(null);
      }
    }

    return _selectedRange;
  }

  ///
  /// returns `2-D dynamic List` of the sheet elements in that range.
  ///
  /// Ex. selectRange('D8:H12'); or selectRange('D8');
  ///
  List<List<dynamic>?> selectRangeValuesWithString(String range) {
    List<List<dynamic>?> _selectedRange = <List<dynamic>?>[];
    if (!range.contains(':')) {
      var start = CellIndex.indexByString(range);
      _selectedRange = selectRangeValues(start);
    } else {
      var rangeVars = range.split(':');
      var start = CellIndex.indexByString(rangeVars[0]);
      var end = CellIndex.indexByString(rangeVars[1]);
      _selectedRange = selectRangeValues(start, end: end);
    }
    return _selectedRange;
  }

  ///
  /// returns `2-D dynamic List` of the sheet elements in that range.
  ///
  List<List<dynamic>?> selectRangeValues(CellIndex start, {CellIndex? end}) {
    var _list =
        (end == null ? selectRange(start) : selectRange(start, end: end));
    return _list
        .map((List<Data?>? e) =>
            e?.map((e1) => e1 != null ? e1.value : null).toList())
        .toList();
  }

  ///
  /// updates count of rows and columns
  ///
  _countRowsAndColumns() {
    int maximumColumnIndex = -1, maximumRowIndex = -1;
    List<int> sortedKeys = _sheetData.keys.toList()..sort();
    sortedKeys.forEach((rowKey) {
      if (_sheetData[rowKey] != null && _sheetData[rowKey]!.isNotEmpty) {
        List<int> keys = _sheetData[rowKey]!.keys.toList()..sort();
        if (keys.isNotEmpty && keys.last > maximumColumnIndex) {
          maximumColumnIndex = keys.last;
        }
      }
    });

    if (sortedKeys.isNotEmpty) {
      maximumRowIndex = sortedKeys.last;
    }

    _maxColumns = maximumColumnIndex + 1;
    _maxRows = maximumRowIndex + 1;
  }

  ///
  /// If `sheet` exists and `columnIndex < maxColumns` then it removes column at index = `columnIndex`
  ///
  void removeColumn(int columnIndex) {
    _checkMaxColumn(columnIndex);
    if (columnIndex < 0 || columnIndex >= maxColumns) {
      return;
    }

    bool updateSpanCell = false;

    /// Do the shifting of the cell Id of span Object

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }
      int startColumn = spanObj.columnSpanStart,
          startRow = spanObj.rowSpanStart,
          endColumn = spanObj.columnSpanEnd,
          endRow = spanObj.rowSpanEnd;

      if (columnIndex <= endColumn) {
        if (columnIndex < startColumn) {
          startColumn -= 1;
        }
        endColumn -= 1;
        if (/* startColumn >= endColumn */
            (columnIndex == (endColumn + 1)) &&
                (columnIndex ==
                    (columnIndex < startColumn
                        ? startColumn + 1
                        : startColumn))) {
          _spanList[i] = null;
        } else {
          _Span newSpanObj = _Span(
            rowSpanStart: startRow,
            columnSpanStart: startColumn,
            rowSpanEnd: endRow,
            columnSpanEnd: endColumn,
          );
          _spanList[i] = newSpanObj;
        }
        updateSpanCell = true;
        _excel._mergeChanges = true;
      }

      if (_spanList[i] != null) {
        String rc = getSpanCellId(startColumn, startRow, endColumn, endRow);
        if (!_spannedItems.contains(rc)) {
          _spannedItems.add(rc);
        }
      }
    }
    _cleanUpSpanMap();

    if (updateSpanCell) {
      _excel._mergeChangeLookup = sheetName;
    }

    Map<int, Map<int, Data>> _data = Map<int, Map<int, Data>>();
    if (columnIndex <= maxColumns - 1) {
      /// do the shifting task
      List<int> sortedKeys = _sheetData.keys.toList()..sort();
      sortedKeys.forEach((rowKey) {
        Map<int, Data> columnMap = Map<int, Data>();
        List<int> sortedColumnKeys = _sheetData[rowKey]!.keys.toList()..sort();
        sortedColumnKeys.forEach((columnKey) {
          if (_sheetData[rowKey] != null &&
              _sheetData[rowKey]![columnKey] != null) {
            if (columnKey < columnIndex) {
              columnMap[columnKey] = _sheetData[rowKey]![columnKey]!;
            }
            if (columnIndex == columnKey) {
              _sheetData[rowKey]!.remove(columnKey);
            }
            if (columnIndex < columnKey) {
              columnMap[columnKey - 1] = _sheetData[rowKey]![columnKey]!;
              _sheetData[rowKey]!.remove(columnKey);
            }
          }
        });
        _data[rowKey] = Map<int, Data>.from(columnMap);
      });
      _sheetData = Map<int, Map<int, Data>>.from(_data);
    }

    if (_maxColumns - 1 <= columnIndex) {
      _maxColumns -= 1;
    }
  }

  ///
  /// Inserts an empty `column` in sheet at position = `columnIndex`.
  ///
  /// If `columnIndex == null` or `columnIndex < 0` if will not execute
  ///
  /// If the `sheet` does not exists then it will be created automatically.
  ///
  void insertColumn(int columnIndex) {
    if (columnIndex < 0) {
      return;
    }
    _checkMaxColumn(columnIndex);

    bool updateSpanCell = false;

    _spannedItems = FastList<String>();
    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }
      int startColumn = spanObj.columnSpanStart,
          startRow = spanObj.rowSpanStart,
          endColumn = spanObj.columnSpanEnd,
          endRow = spanObj.rowSpanEnd;

      if (columnIndex <= endColumn) {
        if (columnIndex <= startColumn) {
          startColumn += 1;
        }
        endColumn += 1;
        _Span newSpanObj = _Span(
          rowSpanStart: startRow,
          columnSpanStart: startColumn,
          rowSpanEnd: endRow,
          columnSpanEnd: endColumn,
        );
        _spanList[i] = newSpanObj;
        updateSpanCell = true;
        _excel._mergeChanges = true;
      }
      String rc = getSpanCellId(startColumn, startRow, endColumn, endRow);
      if (!_spannedItems.contains(rc)) {
        _spannedItems.add(rc);
      }
    }

    if (updateSpanCell) {
      _excel._mergeChangeLookup = sheetName;
    }

    if (_sheetData.isNotEmpty) {
      final Map<int, Map<int, Data>> _data = Map<int, Map<int, Data>>();
      final List<int> sortedKeys = _sheetData.keys.toList()..sort();
      if (columnIndex <= maxColumns - 1) {
        /// do the shifting task
        sortedKeys.forEach((rowKey) {
          final Map<int, Data> columnMap = Map<int, Data>();

          /// getting the column keys in descending order so as to shifting becomes easy
          final List<int> sortedColumnKeys = _sheetData[rowKey]!.keys.toList()
            ..sort((a, b) {
              return b.compareTo(a);
            });
          sortedColumnKeys.forEach((columnKey) {
            if (_sheetData[rowKey] != null &&
                _sheetData[rowKey]![columnKey] != null) {
              if (columnKey < columnIndex) {
                columnMap[columnKey] = _sheetData[rowKey]![columnKey]!;
              }
              if (columnIndex <= columnKey) {
                columnMap[columnKey + 1] = _sheetData[rowKey]![columnKey]!;
              }
            }
          });
          columnMap[columnIndex] = Data.newData(this, rowKey, columnIndex);
          _data[rowKey] = Map<int, Data>.from(columnMap);
        });
        _sheetData = Map<int, Map<int, Data>>.from(_data);
      } else {
        /// just put the data in the very first available row and
        /// in the desired Column index only one time as we will be using less space on internal implementatoin
        /// and mock the user as if the 2-D list is being saved
        ///
        /// As when user calls DataObject.cells then we will output 2-D list - pretending.
        _sheetData[sortedKeys.first]![columnIndex] =
            Data.newData(this, sortedKeys.first, columnIndex);
      }
    } else {
      /// here simply just take the first row and put the columnIndex as the _sheetData was previously null
      _sheetData = Map<int, Map<int, Data>>();
      _sheetData[0] = {columnIndex: Data.newData(this, 0, columnIndex)};
    }
    if (_maxColumns - 1 <= columnIndex) {
      _maxColumns += 1;
    } else {
      _maxColumns = columnIndex + 1;
    }

    //_countRowsAndColumns();
  }

  ///
  /// If `sheet` exists and `rowIndex < maxRows` then it removes row at index = `rowIndex`
  ///
  void removeRow(int rowIndex) {
    if (rowIndex < 0 || rowIndex >= _maxRows) {
      return;
    }
    _checkMaxRow(rowIndex);

    bool updateSpanCell = false;

    for (int i = 0; i < _spanList.length; i++) {
      final _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }
      int startColumn = spanObj.columnSpanStart,
          startRow = spanObj.rowSpanStart,
          endColumn = spanObj.columnSpanEnd,
          endRow = spanObj.rowSpanEnd;

      if (rowIndex <= endRow) {
        if (rowIndex < startRow) {
          startRow -= 1;
        }
        endRow -= 1;
        if (/* startRow >= endRow */
            (rowIndex == (endRow + 1)) &&
                (rowIndex == (rowIndex < startRow ? startRow + 1 : startRow))) {
          _spanList[i] = null;
        } else {
          final _Span newSpanObj = _Span(
            rowSpanStart: startRow,
            columnSpanStart: startColumn,
            rowSpanEnd: endRow,
            columnSpanEnd: endColumn,
          );
          _spanList[i] = newSpanObj;
        }
        updateSpanCell = true;
        _excel._mergeChanges = true;
      }
      if (_spanList[i] != null) {
        final String rc =
            getSpanCellId(startColumn, startRow, endColumn, endRow);
        if (!_spannedItems.contains(rc)) {
          _spannedItems.add(rc);
        }
      }
    }
    _cleanUpSpanMap();

    if (updateSpanCell) {
      _excel._mergeChangeLookup = sheetName;
    }

    if (_sheetData.isNotEmpty) {
      final Map<int, Map<int, Data>> _data = Map<int, Map<int, Data>>();
      if (rowIndex <= maxRows - 1) {
        /// do the shifting task
        final List<int> sortedKeys = _sheetData.keys.toList()..sort();
        sortedKeys.forEach((rowKey) {
          if (rowKey < rowIndex && _sheetData[rowKey] != null) {
            _data[rowKey] = Map<int, Data>.from(_sheetData[rowKey]!);
          }
          if (rowIndex < rowKey && _sheetData[rowKey] != null) {
            _data[rowKey - 1] = Map<int, Data>.from(_sheetData[rowKey]!);
          }
        });
        _sheetData = Map<int, Map<int, Data>>.from(_data);
      }
      //_countRowsAndColumns();
    } else {
      _maxRows = 0;
      _maxColumns = 0;
    }

    if (_maxRows - 1 <= rowIndex) {
      _maxRows -= 1;
    }
  }

  ///
  /// Inserts an empty row in `sheet` at position = `rowIndex`.
  ///
  /// If `rowIndex == null` or `rowIndex < 0` if will not execute
  ///
  /// If the `sheet` does not exists then it will be created automatically.
  ///
  void insertRow(int rowIndex) {
    if (rowIndex < 0) {
      return;
    }

    _checkMaxRow(rowIndex);

    bool updateSpanCell = false;

    _spannedItems = FastList<String>();
    for (int i = 0; i < _spanList.length; i++) {
      final _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }
      int startColumn = spanObj.columnSpanStart,
          startRow = spanObj.rowSpanStart,
          endColumn = spanObj.columnSpanEnd,
          endRow = spanObj.rowSpanEnd;

      if (rowIndex <= endRow) {
        if (rowIndex <= startRow) {
          startRow += 1;
        }
        endRow += 1;
        final _Span newSpanObj = _Span(
          rowSpanStart: startRow,
          columnSpanStart: startColumn,
          rowSpanEnd: endRow,
          columnSpanEnd: endColumn,
        );
        _spanList[i] = newSpanObj;
        updateSpanCell = true;
        _excel._mergeChanges = true;
      }
      String rc = getSpanCellId(startColumn, startRow, endColumn, endRow);
      if (!_spannedItems.contains(rc)) {
        _spannedItems.add(rc);
      }
    }

    if (updateSpanCell) {
      _excel._mergeChangeLookup = sheetName;
    }

    Map<int, Map<int, Data>> _data = Map<int, Map<int, Data>>();
    if (_sheetData.isNotEmpty) {
      List<int> sortedKeys = _sheetData.keys.toList()
        ..sort((a, b) {
          return b.compareTo(a);
        });
      if (rowIndex <= maxRows - 1) {
        /// do the shifting task
        sortedKeys.forEach((rowKey) {
          if (rowKey < rowIndex) {
            _data[rowKey] = _sheetData[rowKey]!;
          }
          if (rowIndex <= rowKey) {
            _data[rowKey + 1] = _sheetData[rowKey]!;
            _data[rowKey + 1]!.forEach((key, value) {
              value._rowIndex++;
            });
          }
        });
      }
    }
    _data[rowIndex] = {0: Data.newData(this, rowIndex, 0)};
    _sheetData = Map<int, Map<int, Data>>.from(_data);

    if (_maxRows - 1 <= rowIndex) {
      _maxRows = rowIndex + 1;
    } else {
      _maxRows += 1;
    }

    //_countRowsAndColumns();
  }

  ///
  /// Updates the contents of `sheet` of the `cellIndex: CellIndex.indexByColumnRow(0, 0);` where indexing starts from 0
  ///
  /// ----or---- by `cellIndex: CellIndex.indexByString("A3");`.
  ///
  /// Styling of cell can be done by passing the CellStyle object to `cellStyle`.
  ///
  /// If `sheet` does not exist then it will be automatically created.
  ///
  void updateCell(CellIndex cellIndex, CellValue? value,
      {CellStyle? cellStyle}) {
    int columnIndex = cellIndex.columnIndex;
    int rowIndex = cellIndex.rowIndex;
    if (columnIndex < 0 || rowIndex < 0) {
      return;
    }
    _checkMaxColumn(columnIndex);
    _checkMaxRow(rowIndex);

    int newRowIndex = rowIndex, newColumnIndex = columnIndex;

    /// Check if this is lying in merged-cell cross-section
    /// If yes then get the starting position of merged cells
    if (_spanList.isNotEmpty) {
      (newRowIndex, newColumnIndex) = _isInsideSpanning(rowIndex, columnIndex);
    }

    /// Puts Data
    _putData(newRowIndex, newColumnIndex, value);

    // check if the numberFormat works with the value provided
    // otherwise fall back to the default for this value type
    if (cellStyle != null) {
      final numberFormat = cellStyle.numberFormat;
      if (!numberFormat.accepts(value)) {
        cellStyle =
            cellStyle.copyWith(numberFormat: NumFormat.defaultFor(value));
      }
    } else {
      final cellStyleBefore =
          _sheetData[cellIndex.rowIndex]?[cellIndex.columnIndex]?.cellStyle;
      if (cellStyleBefore != null &&
          !cellStyleBefore.numberFormat.accepts(value)) {
        cellStyle =
            cellStyleBefore.copyWith(numberFormat: NumFormat.defaultFor(value));
      }
    }

    /// Puts the cellStyle
    if (cellStyle != null) {
      _sheetData[newRowIndex]![newColumnIndex]!._cellStyle = cellStyle;
      _excel._styleChanges = true;
    }
  }

  ///
  /// Merges the cells starting from `start` to `end`.
  ///
  /// If `custom value` is not defined then it will look for the very first available value in range `start` to `end` by searching row-wise from left to right.
  ///
  void merge(CellIndex start, CellIndex end, {CellValue? customValue}) {
    int startColumn = start.columnIndex,
        startRow = start.rowIndex,
        endColumn = end.columnIndex,
        endRow = end.rowIndex;

    _checkMaxColumn(startColumn);
    _checkMaxColumn(endColumn);
    _checkMaxRow(startRow);
    _checkMaxRow(endRow);

    if ((startColumn == endColumn && startRow == endRow) ||
        (startColumn < 0 || startRow < 0 || endColumn < 0 || endRow < 0) ||
        (_spannedItems.contains(
            getSpanCellId(startColumn, startRow, endColumn, endRow)))) {
      return;
    }

    List<int> gotPosition = _getSpanPosition(start, end);

    _excel._mergeChanges = true;

    startColumn = gotPosition[0];
    startRow = gotPosition[1];
    endColumn = gotPosition[2];
    endRow = gotPosition[3];

    // Update maxColumns maxRows
    _maxColumns = _maxColumns > endColumn ? _maxColumns : endColumn + 1;
    _maxRows = _maxRows > endRow ? _maxRows : endRow + 1;

    bool getValue = true;

    Data value = Data.newData(this, startRow, startColumn);
    if (customValue != null) {
      value._value = customValue;
      getValue = false;
    }

    for (int j = startRow; j <= endRow; j++) {
      for (int k = startColumn; k <= endColumn; k++) {
        if (_sheetData[j] != null) {
          if (getValue && _sheetData[j]![k]?.value != null) {
            value = _sheetData[j]![k]!;
            getValue = false;
          }
          _sheetData[j]!.remove(k);
        }
      }
    }

    if (_sheetData[startRow] != null) {
      _sheetData[startRow]![startColumn] = value;
    } else {
      _sheetData[startRow] = {startColumn: value};
    }

    String sp = getSpanCellId(startColumn, startRow, endColumn, endRow);

    if (!_spannedItems.contains(sp)) {
      _spannedItems.add(sp);
    }

    _Span s = _Span(
      rowSpanStart: startRow,
      columnSpanStart: startColumn,
      rowSpanEnd: endRow,
      columnSpanEnd: endColumn,
    );

    _spanList.add(s);
    _excel._mergeChangeLookup = sheetName;
  }

  ///
  /// unMerge the merged cells.
  ///
  ///        var sheet = 'DesiredSheet';
  ///        List<String> spannedCells = excel.getMergedCells(sheet);
  ///        var cellToUnMerge = "A1:A2";
  ///        excel.unMerge(sheet, cellToUnMerge);
  ///
  void unMerge(String unmergeCells) {
    if (_spannedItems.isNotEmpty &&
        _spanList.isNotEmpty &&
        _spannedItems.contains(unmergeCells)) {
      List<String> lis = unmergeCells.split(RegExp(r":"));
      if (lis.length == 2) {
        bool remove = false;
        CellIndex start = CellIndex.indexByString(lis[0]),
            end = CellIndex.indexByString(lis[1]);
        for (int i = 0; i < _spanList.length; i++) {
          _Span? spanObject = _spanList[i];
          if (spanObject == null) {
            continue;
          }

          if (spanObject.columnSpanStart == start.columnIndex &&
              spanObject.rowSpanStart == start.rowIndex &&
              spanObject.columnSpanEnd == end.columnIndex &&
              spanObject.rowSpanEnd == end.rowIndex) {
            _spanList[i] = null;
            remove = true;
          }
        }
        if (remove) {
          _cleanUpSpanMap();
        }
      }
      _spannedItems.remove(unmergeCells);
      _excel._mergeChangeLookup = sheetName;
    }
  }

  ///
  /// Sets the cellStyle of the merged cells.
  ///
  /// It will get the merged cells only by giving the starting position of merged cells.
  ///
  void setMergedCellStyle(CellIndex start, CellStyle mergedCellStyle) {
    List<List<CellIndex>> _mergedCells = spannedItems
        .map(
          (e) => e.split(":").map((e) => CellIndex.indexByString(e)).toList(),
        )
        .toList();

    List<CellIndex> _startIndices = _mergedCells.map((e) => e[0]).toList();
    List<CellIndex> _endIndices = _mergedCells.map((e) => e[1]).toList();

    if (_mergedCells.isEmpty ||
        start.columnIndex < 0 ||
        start.rowIndex < 0 ||
        !_startIndices.contains(start)) {
      return;
    }

    CellIndex end = _endIndices[_startIndices.indexOf(start)];

    bool hasBorder = mergedCellStyle.topBorder != Border() ||
        mergedCellStyle.bottomBorder != Border() ||
        mergedCellStyle.leftBorder != Border() ||
        mergedCellStyle.rightBorder != Border() ||
        mergedCellStyle.diagonalBorderUp ||
        mergedCellStyle.diagonalBorderDown;
    if (hasBorder) {
      for (var i = start.rowIndex; i <= end.rowIndex; i++) {
        for (var j = start.columnIndex; j <= end.columnIndex; j++) {
          CellStyle cellStyle = mergedCellStyle.copyWith(
            topBorderVal: Border(),
            bottomBorderVal: Border(),
            leftBorderVal: Border(),
            rightBorderVal: Border(),
            diagonalBorderUpVal: false,
            diagonalBorderDownVal: false,
          );

          if (i == start.rowIndex) {
            cellStyle = cellStyle.copyWith(
              topBorderVal: mergedCellStyle.topBorder,
            );
          }
          if (i == end.rowIndex) {
            cellStyle = cellStyle.copyWith(
              bottomBorderVal: mergedCellStyle.bottomBorder,
            );
          }
          if (j == start.columnIndex) {
            cellStyle = cellStyle.copyWith(
              leftBorderVal: mergedCellStyle.leftBorder,
            );
          }
          if (j == end.columnIndex) {
            cellStyle = cellStyle.copyWith(
              rightBorderVal: mergedCellStyle.rightBorder,
            );
          }

          if (i == j ||
              start.rowIndex == end.rowIndex ||
              start.columnIndex == end.columnIndex) {
            cellStyle = cellStyle.copyWith(
              diagonalBorderUpVal: mergedCellStyle.diagonalBorderUp,
              diagonalBorderDownVal: mergedCellStyle.diagonalBorderDown,
            );
          }

          if (i == start.rowIndex && j == start.columnIndex) {
            cell(start).cellStyle = cellStyle;
          } else {
            _putData(i, j, null);
            _sheetData[i]![j]!.cellStyle = cellStyle;
          }
        }
      }
    }
  }

  ///
  /// Helps to find the interaction between the pre-existing span position and updates if with new span if there any interaction(Cross-Sectional Spanning) exists.
  ///
  List<int> _getSpanPosition(CellIndex start, CellIndex end) {
    int startColumn = start.columnIndex,
        startRow = start.rowIndex,
        endColumn = end.columnIndex,
        endRow = end.rowIndex;

    bool remove = false;

    if (startRow > endRow) {
      startRow = end.rowIndex;
      endRow = start.rowIndex;
    }
    if (endColumn < startColumn) {
      endColumn = start.columnIndex;
      startColumn = end.columnIndex;
    }

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }

      final locationChange = _isLocationChangeRequired(
          startColumn, startRow, endColumn, endRow, spanObj);

      if (locationChange.$1) {
        startColumn = locationChange.$2.$1;
        startRow = locationChange.$2.$2;
        endColumn = locationChange.$2.$3;
        endRow = locationChange.$2.$4;
        String sp = getSpanCellId(spanObj.columnSpanStart, spanObj.rowSpanStart,
            spanObj.columnSpanEnd, spanObj.rowSpanEnd);
        if (_spannedItems.contains(sp)) {
          _spannedItems.remove(sp);
        }
        remove = true;
        _spanList[i] = null;
      }
    }
    if (remove) {
      _cleanUpSpanMap();
    }

    return [startColumn, startRow, endColumn, endRow];
  }

  ///
  /// Appends [row] iterables just post the last filled `rowIndex`.
  ///
  void appendRow(List<CellValue?> row) {
    int targetRow = maxRows;
    insertRowIterables(row, targetRow);
  }

  /// getting the List of _Span Objects which have the rowIndex containing and
  /// also lower the range by giving the starting columnIndex
  List<_Span> _getSpannedObjects(int rowIndex, int startingColumnIndex) {
    List<_Span> obtained = <_Span>[];

    if (_spanList.isNotEmpty) {
      obtained = <_Span>[];
      _spanList.forEach((spanObject) {
        if (spanObject != null &&
            spanObject.rowSpanStart <= rowIndex &&
            rowIndex <= spanObject.rowSpanEnd &&
            startingColumnIndex <= spanObject.columnSpanEnd) {
          obtained.add(spanObject);
        }
      });
    }
    return obtained;
  }

  ///
  /// Checking if the columnIndex and the rowIndex passed is inside the spanObjectList which is got from calling function.
  ///
  bool _isInsideSpanObject(
      List<_Span> spanObjectList, int columnIndex, int rowIndex) {
    for (int i = 0; i < spanObjectList.length; i++) {
      _Span spanObject = spanObjectList[i];

      if (spanObject.columnSpanStart <= columnIndex &&
          columnIndex <= spanObject.columnSpanEnd &&
          spanObject.rowSpanStart <= rowIndex &&
          rowIndex <= spanObject.rowSpanEnd) {
        if (columnIndex < spanObject.columnSpanEnd) {
          return false;
        } else if (columnIndex == spanObject.columnSpanEnd) {
          return true;
        }
      }
    }
    return true;
  }

  ///
  /// Adds the [row] iterables in the given rowIndex = [rowIndex] in [sheet]
  ///
  /// [startingColumn] tells from where we should start putting the [row] iterables
  ///
  /// [overwriteMergedCells] when set to [true] will over-write mergedCell and does not jumps to next unqiue cell.
  ///
  /// [overwriteMergedCells] when set to [false] puts the cell value in next unique cell available and putting the value in merged cells only once.
  ///
  void insertRowIterables(
    List<CellValue?> row,
    int rowIndex, {
    int startingColumn = 0,
    bool overwriteMergedCells = true,
  }) {
    if (row.isEmpty || rowIndex < 0) {
      return;
    }

    _checkMaxRow(rowIndex);
    int columnIndex = 0;
    if (startingColumn > 0) {
      columnIndex = startingColumn;
    }
    _checkMaxColumn(columnIndex + row.length);
    int rowsLength = _maxRows,
        maxIterationIndex = row.length - 1,
        currentRowPosition = 0; // position in [row] iterables

    if (overwriteMergedCells || rowIndex >= rowsLength) {
      // Normally iterating and putting the data present in the [row] as we are on the last index.

      while (currentRowPosition <= maxIterationIndex) {
        _putData(rowIndex, columnIndex++, row[currentRowPosition++]);
      }
    } else {
      // expensive function as per time complexity
      _selfCorrectSpanMap(_excel);
      List<_Span> _spanObjectsList = _getSpannedObjects(rowIndex, columnIndex);

      if (_spanObjectsList.isEmpty) {
        while (currentRowPosition <= maxIterationIndex) {
          _putData(rowIndex, columnIndex++, row[currentRowPosition++]);
        }
      } else {
        while (currentRowPosition <= maxIterationIndex) {
          if (_isInsideSpanObject(_spanObjectsList, columnIndex, rowIndex)) {
            _putData(rowIndex, columnIndex, row[currentRowPosition++]);
          }
          columnIndex++;
        }
      }
    }
  }

  ///
  /// Internal function for putting the data in `_sheetData`.
  ///
  void _putData(int rowIndex, int columnIndex, CellValue? value) {
    var row = _sheetData[rowIndex];
    if (row == null) {
      _sheetData[rowIndex] = row = {};
    }
    var cell = row[columnIndex];
    if (cell == null) {
      row[columnIndex] = cell = Data.newData(this, rowIndex, columnIndex);
    }

    cell._value = value;
    cell._cellStyle = CellStyle(numberFormat: NumFormat.defaultFor(value));
    if (cell._cellStyle != NumFormat.standard_0) {
      _excel._styleChanges = true;
    }

    if ((_maxColumns - 1) < columnIndex) {
      _maxColumns = columnIndex + 1;
    }

    if ((_maxRows - 1) < rowIndex) {
      _maxRows = rowIndex + 1;
    }

    //_countRowsAndColumns();
  }

  double? get defaultRowHeight => _defaultRowHeight;

  double? get defaultColumnWidth => _defaultColumnWidth;

  ///
  /// returns map of auto fit columns
  ///
  Map<int, bool> get getColumnAutoFits => _columnAutoFit;

  ///
  /// returns map of custom width columns
  ///
  Map<int, double> get getColumnWidths => _columnWidths;

  ///
  /// returns map of custom height rows
  ///
  Map<int, double> get getRowHeights => _rowHeights;

  ///
  /// returns auto fit state of column index
  ///
  bool getColumnAutoFit(int columnIndex) {
    if (_columnAutoFit.containsKey(columnIndex)) {
      return _columnAutoFit[columnIndex]!;
    }
    return false;
  }

  ///
  /// returns width of column index
  ///
  double getColumnWidth(int columnIndex) {
    if (_columnWidths.containsKey(columnIndex)) {
      return _columnWidths[columnIndex]!;
    }
    return _defaultColumnWidth!;
  }

  ///
  /// returns height of row index
  ///
  double getRowHeight(int rowIndex) {
    if (_rowHeights.containsKey(rowIndex)) {
      return _rowHeights[rowIndex]!;
    }
    return _defaultRowHeight!;
  }

  ///
  /// Set the default column width.
  ///
  /// If both `setDefaultRowHeight` and `setDefaultColumnWidth` are not called,
  /// then the default row height and column width will be set by Excel.
  ///
  /// The default row height is 15.0 and the default column width is 8.43.
  ///
  void setDefaultColumnWidth([double columnWidth = _excelDefaultColumnWidth]) {
    if (columnWidth < 0) return;
    _defaultColumnWidth = columnWidth;
  }

  ///
  /// Set the default row height.
  ///
  /// If both `setDefaultRowHeight` and `setDefaultColumnWidth` are not called,
  /// then the default row height and column width will be set by Excel.
  ///
  /// The default row height is 15.0 and the default column width is 8.43.
  ///
  void setDefaultRowHeight([double rowHeight = _excelDefaultRowHeight]) {
    if (rowHeight < 0) return;
    _defaultRowHeight = rowHeight;
  }

  ///
  /// Set Column AutoFit
  ///
  void setColumnAutoFit(int columnIndex) {
    _checkMaxColumn(columnIndex);
    if (columnIndex < 0) return;
    _columnAutoFit[columnIndex] = true;
  }

  ///
  /// Set Column Width
  ///
  void setColumnWidth(int columnIndex, double columnWidth) {
    _checkMaxColumn(columnIndex);
    if (columnWidth < 0) return;
    _columnWidths[columnIndex] = columnWidth;
  }

  ///
  /// Set Row Height
  ///
  void setRowHeight(int rowIndex, double rowHeight) {
    _checkMaxRow(rowIndex);
    if (rowHeight < 0) return;
    _rowHeights[rowIndex] = rowHeight;
  }

  ///
  ///Returns the `count` of replaced `source` with `target`
  ///
  ///`source` is Pattern which allows you to pass your custom `RegExp` or a simple `String` providing more control over it.
  ///
  ///optional argument `first` is used to replace the number of first earlier occurrences
  ///
  ///If `first` is set to `3` then it will replace only first `3 occurrences` of the `source` with `target`.
  ///
  ///       excel.findAndReplace('MySheetName', 'sad', 'happy', first: 3);
  ///
  ///       or
  ///
  ///       var mySheet = excel['mySheetName'];
  ///       mySheet.findAndReplace('sad', 'happy', first: 3);
  ///
  ///In the above example it will replace all the occurences of `sad` with `happy` in the cells
  ///
  ///Other `options` are used to `narrow down` the `starting and ending ranges of cells`.
  ///
  int findAndReplace(Pattern source, String target,
      {int first = -1,
      int startingRow = -1,
      int endingRow = -1,
      int startingColumn = -1,
      int endingColumn = -1}) {
    int replaceCount = 0,
        _startingRow = 0,
        _endingRow = -1,
        _startingColumn = 0,
        _endingColumn = -1;

    if (startingRow != -1 && endingRow != -1) {
      if (startingRow > endingRow) {
        _endingRow = startingRow;
        _startingRow = endingRow;
      } else {
        _endingRow = endingRow;
        _startingRow = startingRow;
      }
    }

    if (startingColumn != -1 && endingColumn != -1) {
      if (startingColumn > endingColumn) {
        _endingColumn = startingColumn;
        _startingColumn = endingColumn;
      } else {
        _endingColumn = endingColumn;
        _startingColumn = startingColumn;
      }
    }

    int rowsLength = maxRows, columnLength = maxColumns;

    for (int i = _startingRow; i < rowsLength; i++) {
      if (_endingRow != -1 && i > _endingRow) {
        break;
      }
      for (int j = _startingColumn; j < columnLength; j++) {
        if (_endingColumn != -1 && j > _endingColumn) {
          break;
        }
        final sourceData = _sheetData[i]?[j]?.value;
        if (sourceData is! TextCellValue) {
          continue;
        }
        final result =
            sourceData.value.toString().replaceAllMapped(source, (match) {
          if (first == -1 || first != replaceCount) {
            ++replaceCount;
            return match.input.replaceRange(match.start, match.end, target);
          }
          return match.input;
        });
        _sheetData[i]![j]!.value = TextCellValue(result);
      }
    }

    return replaceCount;
  }

  ///
  /// returns `true` if the contents are successfully `cleared` else `false`.
  ///
  /// If the row is having any spanned-cells then it will not be cleared and hence returns `false`.
  ///
  bool clearRow(int rowIndex) {
    if (rowIndex < 0) {
      return false;
    }

    /// lets assume that this row is already cleared and is not inside spanList
    /// If this row exists then we check for the span condition
    bool isNotInside = true;

    if (_sheetData[rowIndex] != null && _sheetData[rowIndex]!.isNotEmpty) {
      /// lets start iterating the spanList and check that if the row is inside the spanList or not
      /// we will expect that value of isNotInside should not be changed to false
      /// If it changes to false then we can't clear this row as it is inside the spanned Cells
      for (int i = 0; i < _spanList.length; i++) {
        _Span? spanObj = _spanList[i];
        if (spanObj == null) {
          continue;
        }
        if (rowIndex >= spanObj.rowSpanStart &&
            rowIndex <= spanObj.rowSpanEnd) {
          isNotInside = false;
          break;
        }
      }

      /// As the row is not inside any SpanList so we can easily clear its content.
      if (isNotInside) {
        _sheetData[rowIndex]!.keys.toList().forEach((key) {
          /// Main concern here is to [clear the contents] and [not remove] the entire row or the cell block
          _sheetData[rowIndex]![key] = Data.newData(this, rowIndex, key);
        });
      }
    }
    //_countRowsAndColumns();
    return isNotInside;
  }

  ///
  ///It is used to check if cell at rowIndex, columnIndex is inside any spanning cell or not ?
  ///
  ///If it exist then the very first index of than spanned cells is returned in order to point to the starting cell
  ///otherwise the parameters are returned back.
  ///
  (int newRowIndex, int newColumnIndex) _isInsideSpanning(
      int rowIndex, int columnIndex) {
    int newRowIndex = rowIndex, newColumnIndex = columnIndex;

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }

      if (rowIndex >= spanObj.rowSpanStart &&
          rowIndex <= spanObj.rowSpanEnd &&
          columnIndex >= spanObj.columnSpanStart &&
          columnIndex <= spanObj.columnSpanEnd) {
        newRowIndex = spanObj.rowSpanStart;
        newColumnIndex = spanObj.columnSpanStart;
        break;
      }
    }

    return (newRowIndex, newColumnIndex);
  }

  ///
  ///Check if columnIndex is not out of `Excel Column limits`.
  ///
  void _checkMaxColumn(int columnIndex) {
    if (_maxColumns >= 16384 || columnIndex >= 16384) {
      throw ArgumentError('Reached Max (16384) or (XFD) columns value.');
    }
    if (columnIndex < 0) {
      throw ArgumentError('Negative columnIndex found: $columnIndex');
    }
  }

  ///
  ///Check if rowIndex is not out of `Excel Row limits`.
  ///
  void _checkMaxRow(int rowIndex) {
    if (_maxRows >= 1048576 || rowIndex >= 1048576) {
      throw ArgumentError('Reached Max (1048576) rows value.');
    }
    if (rowIndex < 0) {
      throw ArgumentError('Negative rowIndex found: $rowIndex');
    }
  }

  ///
  ///returns List of Spanned Cells as
  ///
  ///     ["A1:A2", "A4:G6", "Y4:Y6", ....]
  ///
  ///return type if String based cell-id
  ///
  List<String> get spannedItems {
    _spannedItems = FastList<String>();

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }
      String rC = getSpanCellId(spanObj.columnSpanStart, spanObj.rowSpanStart,
          spanObj.columnSpanEnd, spanObj.rowSpanEnd);
      if (!_spannedItems.contains(rC)) {
        _spannedItems.add(rC);
      }
    }

    return _spannedItems.keys;
  }

  ///
  ///Cleans the `_SpanList` by removing the indexes where null value exists.
  ///
  void _cleanUpSpanMap() {
    if (_spanList.isNotEmpty) {
      _spanList.removeWhere((value) {
        return value == null;
      });
    }
  }

  ///return `SheetName`
  String get sheetName {
    return _sheet;
  }

  ///returns row at index = `rowIndex`
  List<Data?> row(int rowIndex) {
    if (rowIndex < 0) {
      return <Data?>[];
    }
    if (rowIndex < _maxRows) {
      if (_sheetData[rowIndex] != null) {
        return List.generate(_maxColumns, (columnIndex) {
          if (_sheetData[rowIndex]![columnIndex] != null) {
            return _sheetData[rowIndex]![columnIndex]!;
          }
          return null;
        });
      } else {
        return List.generate(_maxColumns, (_) => null);
      }
    }
    return <Data?>[];
  }

  ///
  ///returns count of `rows` having data in `sheet`
  ///
  int get maxRows {
    return _maxRows;
  }

  ///
  ///returns count of `columns` having data in `sheet`
  ///
  int get maxColumns {
    return _maxColumns;
  }

  HeaderFooter? get headerFooter {
    return _headerFooter;
  }

  set headerFooter(HeaderFooter? headerFooter) {
    _headerFooter = headerFooter;
  }
}



================================================
FILE: lib/src/utilities/archive.dart
================================================
part of excel;

Archive _cloneArchive(
  Archive archive,
  Map<String, ArchiveFile> _archiveFiles, {
  String? excludedFile,
}) {
  var clone = Archive();
  archive.files.forEach((file) {
    if (file.isFile) {
      if (excludedFile != null &&
          file.name.toLowerCase() == excludedFile.toLowerCase()) {
        return;
      }
      ArchiveFile copy;
      if (_archiveFiles.containsKey(file.name)) {
        copy = _archiveFiles[file.name]!;
      } else {
        var content = file.content;
        var compression = _noCompression.contains(file.name)
            ? CompressionType.none
            : CompressionType.deflate;
        copy = ArchiveFile(file.name, content.length, content)
          ..compression = compression;
      }
      clone.addFile(copy);
    }
  });
  return clone;
}



================================================
FILE: lib/src/utilities/colors.dart
================================================
part of excel;

String _decimalToHexadecimal(int decimalVal) {
  if (decimalVal == 0) {
    return '0';
  }
  bool negative = false;
  if (decimalVal < 0) {
    negative = true;
    decimalVal *= -1;
  }
  String hexString = '';
  while (decimalVal > 0) {
    String hexVal = '';
    final int remainder = decimalVal % 16;
    decimalVal = decimalVal ~/ 16;
    if (_hexTable.containsKey(remainder)) {
      hexVal = _hexTable[remainder]!;
    } else {
      hexVal = remainder.toString();
    }
    hexString = hexVal + hexString;
  }
  return negative ? '-$hexString' : hexString;
}

bool _assertHexString(String hexString) {
  hexString = hexString.replaceAll('#', '').trim().toUpperCase();

  final bool isNegative = hexString[0] == '-';
  if (isNegative) hexString = hexString.substring(1);

  for (int i = 0; i < hexString.length; i++) {
    if (int.tryParse(hexString[i]) == null &&
        _hexTableReverse.containsKey(hexString[i]) == false) {
      return false;
    }
  }
  return true;
}

int _hexadecimalToDecimal(String hexString) {
  hexString = hexString.replaceAll('#', '').trim().toUpperCase();

  final bool isNegative = hexString[0] == '-';
  if (isNegative) hexString = hexString.substring(1);

  int decimalVal = 0;
  for (int i = 0; i < hexString.length; i++) {
    if (int.tryParse(hexString[i]) == null &&
        _hexTableReverse.containsKey(hexString[i]) == false) {
      throw Exception('Non-hex value was passed to the function');
    } else {
      decimalVal += (pow(16, hexString.length - i - 1) *
              (int.tryParse(hexString[i]) != null
                  ? int.parse(hexString[i])
                  : _hexTableReverse[hexString[i]]!))
          .toInt();
    }
  }
  return isNegative ? -1 * decimalVal : decimalVal;
}

const _hexTable = {
  10: 'A',
  11: 'B',
  12: 'C',
  13: 'D',
  14: 'E',
  15: 'F',
};

final _hexTableReverse = _hexTable.map((k, v) => MapEntry(v, k));

extension StringExt on String {
  /// Return [ExcelColor.black] if not a color hexadecimal
  ExcelColor get excelColor => this == 'none'
      ? ExcelColor.none
      : _assertHexString(this)
          ? ExcelColor.valuesAsMap[this] ?? ExcelColor._(this)
          : ExcelColor.black;
}

/// Copying from Flutter Material Color
class ExcelColor extends Equatable {
  const ExcelColor._(this._color, [this._name, this._type]);

  final String _color;
  final String? _name;
  final ColorType? _type;

  /// Return 'none' if [_color] is null, [black] if not match for safety
  String get colorHex =>
      _assertHexString(_color) || _color == 'none' ? _color : black.colorHex;

  /// Return [black] if [_color] is not match for safety
  int get colorInt =>
      _assertHexString(_color) ? _hexadecimalToDecimal(_color) : black.colorInt;

  ColorType? get type => _type;

  String? get name => _name;

  /// Warning! Highly unsafe method.
  /// Can break your excel file if you do not know what you are doing
  factory ExcelColor.fromInt(int colorIntValue) =>
      ExcelColor._(_decimalToHexadecimal(colorIntValue));

  /// Warning! Highly unsafe method.
  /// Can break your excel file if you do not know what you are doing
  factory ExcelColor.fromHexString(String colorHexValue) =>
      ExcelColor._(colorHexValue);

  static const none = ExcelColor._('none');

  static const black = ExcelColor._('FF000000', 'black', ColorType.color);
  static const black12 = ExcelColor._('1F000000', 'black12', ColorType.color);
  static const black26 = ExcelColor._('42000000', 'black26', ColorType.color);
  static const black38 = ExcelColor._('61000000', 'black38', ColorType.color);
  static const black45 = ExcelColor._('73000000', 'black45', ColorType.color);
  static const black54 = ExcelColor._('8A000000', 'black54', ColorType.color);
  static const black87 = ExcelColor._('DD000000', 'black87', ColorType.color);
  static const white = ExcelColor._('FFFFFFFF', 'white', ColorType.color);
  static const white10 = ExcelColor._('1AFFFFFF', 'white10', ColorType.color);
  static const white12 = ExcelColor._('1FFFFFFF', 'white12', ColorType.color);
  static const white24 = ExcelColor._('3DFFFFFF', 'white24', ColorType.color);
  static const white30 = ExcelColor._('4DFFFFFF', 'white30', ColorType.color);
  static const white38 = ExcelColor._('62FFFFFF', 'white38', ColorType.color);
  static const white54 = ExcelColor._('8AFFFFFF', 'white54', ColorType.color);
  static const white60 = ExcelColor._('99FFFFFF', 'white60', ColorType.color);
  static const white70 = ExcelColor._('B3FFFFFF', 'white70', ColorType.color);
  static const redAccent =
      ExcelColor._('FFFF5252', 'redAccent', ColorType.materialAccent);
  static const pinkAccent =
      ExcelColor._('FFFF4081', 'pinkAccent', ColorType.materialAccent);
  static const purpleAccent =
      ExcelColor._('FFE040FB', 'purpleAccent', ColorType.materialAccent);
  static const deepPurpleAccent =
      ExcelColor._('FF7C4DFF', 'deepPurpleAccent', ColorType.materialAccent);
  static const indigoAccent =
      ExcelColor._('FF536DFE', 'indigoAccent', ColorType.materialAccent);
  static const blueAccent =
      ExcelColor._('FF448AFF', 'blueAccent', ColorType.materialAccent);
  static const lightBlueAccent =
      ExcelColor._('FF40C4FF', 'lightBlueAccent', ColorType.materialAccent);
  static const cyanAccent =
      ExcelColor._('FF18FFFF', 'cyanAccent', ColorType.materialAccent);
  static const tealAccent =
      ExcelColor._('FF64FFDA', 'tealAccent', ColorType.materialAccent);
  static const greenAccent =
      ExcelColor._('FF69F0AE', 'greenAccent', ColorType.materialAccent);
  static const lightGreenAccent =
      ExcelColor._('FFB2FF59', 'lightGreenAccent', ColorType.materialAccent);
  static const limeAccent =
      ExcelColor._('FFEEFF41', 'limeAccent', ColorType.materialAccent);
  static const yellowAccent =
      ExcelColor._('FFFFFF00', 'yellowAccent', ColorType.materialAccent);
  static const amberAccent =
      ExcelColor._('FFFFD740', 'amberAccent', ColorType.materialAccent);
  static const orangeAccent =
      ExcelColor._('FFFFAB40', 'orangeAccent', ColorType.materialAccent);
  static const deepOrangeAccent =
      ExcelColor._('FFFF6E40', 'deepOrangeAccent', ColorType.materialAccent);
  static const red = ExcelColor._('FFF44336', 'red', ColorType.material);
  static const pink = ExcelColor._('FFE91E63', 'pink', ColorType.material);
  static const purple = ExcelColor._('FF9C27B0', 'purple', ColorType.material);
  static const deepPurple =
      ExcelColor._('FF673AB7', 'deepPurple', ColorType.material);
  static const indigo = ExcelColor._('FF3F51B5', 'indigo', ColorType.material);
  static const blue = ExcelColor._('FF2196F3', 'blue', ColorType.material);
  static const lightBlue =
      ExcelColor._('FF03A9F4', 'lightBlue', ColorType.material);
  static const cyan = ExcelColor._('FF00BCD4', 'cyan', ColorType.material);
  static const teal = ExcelColor._('FF009688', 'teal', ColorType.material);
  static const green = ExcelColor._('FF4CAF50', 'green', ColorType.material);
  static const lightGreen =
      ExcelColor._('FF8BC34A', 'lightGreen', ColorType.material);
  static const lime = ExcelColor._('FFCDDC39', 'lime', ColorType.material);
  static const yellow = ExcelColor._('FFFFEB3B', 'yellow', ColorType.material);
  static const amber = ExcelColor._('FFFFC107', 'amber', ColorType.material);
  static const orange = ExcelColor._('FFFF9800', 'orange', ColorType.material);
  static const deepOrange =
      ExcelColor._('FFFF5722', 'deepOrange', ColorType.material);
  static const brown = ExcelColor._('FF795548', 'brown', ColorType.material);
  static const grey = ExcelColor._('FF9E9E9E', 'grey', ColorType.material);
  static const blueGrey =
      ExcelColor._('FF607D8B', 'blueGrey', ColorType.material);
  static const redAccent100 =
      ExcelColor._('FFFF8A80', 'redAccent100', ColorType.materialAccent);
  static const redAccent400 =
      ExcelColor._('FFFF1744', 'redAccent400', ColorType.materialAccent);
  static const redAccent700 =
      ExcelColor._('FFD50000', 'redAccent700', ColorType.materialAccent);
  static const pinkAccent100 =
      ExcelColor._('FFFF80AB', 'pinkAccent100', ColorType.materialAccent);
  static const pinkAccent400 =
      ExcelColor._('FFF50057', 'pinkAccent400', ColorType.materialAccent);
  static const pinkAccent700 =
      ExcelColor._('FFC51162', 'pinkAccent700', ColorType.materialAccent);
  static const purpleAccent100 =
      ExcelColor._('FFEA80FC', 'purpleAccent100', ColorType.materialAccent);
  static const purpleAccent400 =
      ExcelColor._('FFD500F9', 'purpleAccent400', ColorType.materialAccent);
  static const purpleAccent700 =
      ExcelColor._('FFAA00FF', 'purpleAccent700', ColorType.materialAccent);
  static const deepPurpleAccent100 =
      ExcelColor._('FFB388FF', 'deepPurpleAccent100', ColorType.materialAccent);
  static const deepPurpleAccent400 =
      ExcelColor._('FF651FFF', 'deepPurpleAccent400', ColorType.materialAccent);
  static const deepPurpleAccent700 =
      ExcelColor._('FF6200EA', 'deepPurpleAccent700', ColorType.materialAccent);
  static const indigoAccent100 =
      ExcelColor._('FF8C9EFF', 'indigoAccent100', ColorType.materialAccent);
  static const indigoAccent400 =
      ExcelColor._('FF3D5AFE', 'indigoAccent400', ColorType.materialAccent);
  static const indigoAccent700 =
      ExcelColor._('FF304FFE', 'indigoAccent700', ColorType.materialAccent);
  static const blueAccent100 =
      ExcelColor._('FF82B1FF', 'blueAccent100', ColorType.materialAccent);
  static const blueAccent400 =
      ExcelColor._('FF2979FF', 'blueAccent400', ColorType.materialAccent);
  static const blueAccent700 =
      ExcelColor._('FF2962FF', 'blueAccent700', ColorType.materialAccent);
  static const lightBlueAccent100 =
      ExcelColor._('FF80D8FF', 'lightBlueAccent100', ColorType.materialAccent);
  static const lightBlueAccent400 =
      ExcelColor._('FF00B0FF', 'lightBlueAccent400', ColorType.materialAccent);
  static const lightBlueAccent700 =
      ExcelColor._('FF0091EA', 'lightBlueAccent700', ColorType.materialAccent);
  static const cyanAccent100 =
      ExcelColor._('FF84FFFF', 'cyanAccent100', ColorType.materialAccent);
  static const cyanAccent400 =
      ExcelColor._('FF00E5FF', 'cyanAccent400', ColorType.materialAccent);
  static const cyanAccent700 =
      ExcelColor._('FF00B8D4', 'cyanAccent700', ColorType.materialAccent);
  static const tealAccent100 =
      ExcelColor._('FFA7FFEB', 'tealAccent100', ColorType.materialAccent);
  static const tealAccent400 =
      ExcelColor._('FF1DE9B6', 'tealAccent400', ColorType.materialAccent);
  static const tealAccent700 =
      ExcelColor._('FF00BFA5', 'tealAccent700', ColorType.materialAccent);
  static const greenAccent100 =
      ExcelColor._('FFB9F6CA', 'greenAccent100', ColorType.materialAccent);
  static const greenAccent400 =
      ExcelColor._('FF00E676', 'greenAccent400', ColorType.materialAccent);
  static const greenAccent700 =
      ExcelColor._('FF00C853', 'greenAccent700', ColorType.materialAccent);
  static const lightGreenAccent100 =
      ExcelColor._('FFCCFF90', 'lightGreenAccent100', ColorType.materialAccent);
  static const lightGreenAccent400 =
      ExcelColor._('FF76FF03', 'lightGreenAccent400', ColorType.materialAccent);
  static const lightGreenAccent700 =
      ExcelColor._('FF64DD17', 'lightGreenAccent700', ColorType.materialAccent);
  static const limeAccent100 =
      ExcelColor._('FFF4FF81', 'limeAccent100', ColorType.materialAccent);
  static const limeAccent400 =
      ExcelColor._('FFC6FF00', 'limeAccent400', ColorType.materialAccent);
  static const limeAccent700 =
      ExcelColor._('FFAEEA00', 'limeAccent700', ColorType.materialAccent);
  static const yellowAccent100 =
      ExcelColor._('FFFFFF8D', 'yellowAccent100', ColorType.materialAccent);
  static const yellowAccent400 =
      ExcelColor._('FFFFEA00', 'yellowAccent400', ColorType.materialAccent);
  static const yellowAccent700 =
      ExcelColor._('FFFFD600', 'yellowAccent700', ColorType.materialAccent);
  static const amberAccent100 =
      ExcelColor._('FFFFE57F', 'amberAccent100', ColorType.materialAccent);
  static const amberAccent400 =
      ExcelColor._('FFFFC400', 'amberAccent400', ColorType.materialAccent);
  static const amberAccent700 =
      ExcelColor._('FFFFAB00', 'amberAccent700', ColorType.materialAccent);
  static const orangeAccent100 =
      ExcelColor._('FFFFD180', 'orangeAccent100', ColorType.materialAccent);
  static const orangeAccent400 =
      ExcelColor._('FFFF9100', 'orangeAccent400', ColorType.materialAccent);
  static const orangeAccent700 =
      ExcelColor._('FFFF6D00', 'orangeAccent700', ColorType.materialAccent);
  static const deepOrangeAccent100 =
      ExcelColor._('FFFF9E80', 'deepOrangeAccent100', ColorType.materialAccent);
  static const deepOrangeAccent400 =
      ExcelColor._('FFFF3D00', 'deepOrangeAccent400', ColorType.materialAccent);
  static const deepOrangeAccent700 =
      ExcelColor._('FFDD2C00', 'deepOrangeAccent700', ColorType.materialAccent);
  static const red50 = ExcelColor._('FFFFEBEE', 'red50', ColorType.material);
  static const red100 = ExcelColor._('FFFFCDD2', 'red100', ColorType.material);
  static const red200 = ExcelColor._('FFEF9A9A', 'red200', ColorType.material);
  static const red300 = ExcelColor._('FFE57373', 'red300', ColorType.material);
  static const red400 = ExcelColor._('FFEF5350', 'red400', ColorType.material);
  static const red600 = ExcelColor._('FFE53935', 'red600', ColorType.material);
  static const red700 = ExcelColor._('FFD32F2F', 'red700', ColorType.material);
  static const red800 = ExcelColor._('FFC62828', 'red800', ColorType.material);
  static const red900 = ExcelColor._('FFB71C1C', 'red900', ColorType.material);
  static const pink50 = ExcelColor._('FFFCE4EC', 'pink50', ColorType.material);
  static const pink100 =
      ExcelColor._('FFF8BBD0', 'pink100', ColorType.material);
  static const pink200 =
      ExcelColor._('FFF48FB1', 'pink200', ColorType.material);
  static const pink300 =
      ExcelColor._('FFF06292', 'pink300', ColorType.material);
  static const pink400 =
      ExcelColor._('FFEC407A', 'pink400', ColorType.material);
  static const pink600 =
      ExcelColor._('FFD81B60', 'pink600', ColorType.material);
  static const pink700 =
      ExcelColor._('FFC2185B', 'pink700', ColorType.material);
  static const pink800 =
      ExcelColor._('FFAD1457', 'pink800', ColorType.material);
  static const pink900 =
      ExcelColor._('FF880E4F', 'pink900', ColorType.material);
  static const purple50 =
      ExcelColor._('FFF3E5F5', 'purple50', ColorType.material);
  static const purple100 =
      ExcelColor._('FFE1BEE7', 'purple100', ColorType.material);
  static const purple200 =
      ExcelColor._('FFCE93D8', 'purple200', ColorType.material);
  static const purple300 =
      ExcelColor._('FFBA68C8', 'purple300', ColorType.material);
  static const purple400 =
      ExcelColor._('FFAB47BC', 'purple400', ColorType.material);
  static const purple600 =
      ExcelColor._('FF8E24AA', 'purple600', ColorType.material);
  static const purple700 =
      ExcelColor._('FF7B1FA2', 'purple700', ColorType.material);
  static const purple800 =
      ExcelColor._('FF6A1B9A', 'purple800', ColorType.material);
  static const purple900 =
      ExcelColor._('FF4A148C', 'purple900', ColorType.material);
  static const deepPurple50 =
      ExcelColor._('FFEDE7F6', 'deepPurple50', ColorType.material);
  static const deepPurple100 =
      ExcelColor._('FFD1C4E9', 'deepPurple100', ColorType.material);
  static const deepPurple200 =
      ExcelColor._('FFB39DDB', 'deepPurple200', ColorType.material);
  static const deepPurple300 =
      ExcelColor._('FF9575CD', 'deepPurple300', ColorType.material);
  static const deepPurple400 =
      ExcelColor._('FF7E57C2', 'deepPurple400', ColorType.material);
  static const deepPurple600 =
      ExcelColor._('FF5E35B1', 'deepPurple600', ColorType.material);
  static const deepPurple700 =
      ExcelColor._('FF512DA8', 'deepPurple700', ColorType.material);
  static const deepPurple800 =
      ExcelColor._('FF4527A0', 'deepPurple800', ColorType.material);
  static const deepPurple900 =
      ExcelColor._('FF311B92', 'deepPurple900', ColorType.material);
  static const indigo50 =
      ExcelColor._('FFE8EAF6', 'indigo50', ColorType.material);
  static const indigo100 =
      ExcelColor._('FFC5CAE9', 'indigo100', ColorType.material);
  static const indigo200 =
      ExcelColor._('FF9FA8DA', 'indigo200', ColorType.material);
  static const indigo300 =
      ExcelColor._('FF7986CB', 'indigo300', ColorType.material);
  static const indigo400 =
      ExcelColor._('FF5C6BC0', 'indigo400', ColorType.material);
  static const indigo600 =
      ExcelColor._('FF3949AB', 'indigo600', ColorType.material);
  static const indigo700 =
      ExcelColor._('FF303F9F', 'indigo700', ColorType.material);
  static const indigo800 =
      ExcelColor._('FF283593', 'indigo800', ColorType.material);
  static const indigo900 =
      ExcelColor._('FF1A237E', 'indigo900', ColorType.material);
  static const blue50 = ExcelColor._('FFE3F2FD', 'blue50', ColorType.material);
  static const blue100 =
      ExcelColor._('FFBBDEFB', 'blue100', ColorType.material);
  static const blue200 =
      ExcelColor._('FF90CAF9', 'blue200', ColorType.material);
  static const blue300 =
      ExcelColor._('FF64B5F6', 'blue300', ColorType.material);
  static const blue400 =
      ExcelColor._('FF42A5F5', 'blue400', ColorType.material);
  static const blue600 =
      ExcelColor._('FF1E88E5', 'blue600', ColorType.material);
  static const blue700 =
      ExcelColor._('FF1976D2', 'blue700', ColorType.material);
  static const blue800 =
      ExcelColor._('FF1565C0', 'blue800', ColorType.material);
  static const blue900 =
      ExcelColor._('FF0D47A1', 'blue900', ColorType.material);
  static const lightBlue50 =
      ExcelColor._('FFE1F5FE', 'lightBlue50', ColorType.material);
  static const lightBlue100 =
      ExcelColor._('FFB3E5FC', 'lightBlue100', ColorType.material);
  static const lightBlue200 =
      ExcelColor._('FF81D4FA', 'lightBlue200', ColorType.material);
  static const lightBlue300 =
      ExcelColor._('FF4FC3F7', 'lightBlue300', ColorType.material);
  static const lightBlue400 =
      ExcelColor._('FF29B6F6', 'lightBlue400', ColorType.material);
  static const lightBlue600 =
      ExcelColor._('FF039BE5', 'lightBlue600', ColorType.material);
  static const lightBlue700 =
      ExcelColor._('FF0288D1', 'lightBlue700', ColorType.material);
  static const lightBlue800 =
      ExcelColor._('FF0277BD', 'lightBlue800', ColorType.material);
  static const lightBlue900 =
      ExcelColor._('FF01579B', 'lightBlue900', ColorType.material);
  static const cyan50 = ExcelColor._('FFE0F7FA', 'cyan50', ColorType.material);
  static const cyan100 =
      ExcelColor._('FFB2EBF2', 'cyan100', ColorType.material);
  static const cyan200 =
      ExcelColor._('FF80DEEA', 'cyan200', ColorType.material);
  static const cyan300 =
      ExcelColor._('FF4DD0E1', 'cyan300', ColorType.material);
  static const cyan400 =
      ExcelColor._('FF26C6DA', 'cyan400', ColorType.material);
  static const cyan600 =
      ExcelColor._('FF00ACC1', 'cyan600', ColorType.material);
  static const cyan700 =
      ExcelColor._('FF0097A7', 'cyan700', ColorType.material);
  static const cyan800 =
      ExcelColor._('FF00838F', 'cyan800', ColorType.material);
  static const cyan900 =
      ExcelColor._('FF006064', 'cyan900', ColorType.material);
  static const teal50 = ExcelColor._('FFE0F2F1', 'teal50', ColorType.material);
  static const teal100 =
      ExcelColor._('FFB2DFDB', 'teal100', ColorType.material);
  static const teal200 =
      ExcelColor._('FF80CBC4', 'teal200', ColorType.material);
  static const teal300 =
      ExcelColor._('FF4DB6AC', 'teal300', ColorType.material);
  static const teal400 =
      ExcelColor._('FF26A69A', 'teal400', ColorType.material);
  static const teal600 =
      ExcelColor._('FF00897B', 'teal600', ColorType.material);
  static const teal700 =
      ExcelColor._('FF00796B', 'teal700', ColorType.material);
  static const teal800 =
      ExcelColor._('FF00695C', 'teal800', ColorType.material);
  static const teal900 =
      ExcelColor._('FF004D40', 'teal900', ColorType.material);
  static const green50 =
      ExcelColor._('FFE8F5E9', 'green50', ColorType.material);
  static const green100 =
      ExcelColor._('FFC8E6C9', 'green100', ColorType.material);
  static const green200 =
      ExcelColor._('FFA5D6A7', 'green200', ColorType.material);
  static const green300 =
      ExcelColor._('FF81C784', 'green300', ColorType.material);
  static const green400 =
      ExcelColor._('FF66BB6A', 'green400', ColorType.material);
  static const green600 =
      ExcelColor._('FF43A047', 'green600', ColorType.material);
  static const green700 =
      ExcelColor._('FF388E3C', 'green700', ColorType.material);
  static const green800 =
      ExcelColor._('FF2E7D32', 'green800', ColorType.material);
  static const green900 =
      ExcelColor._('FF1B5E20', 'green900', ColorType.material);
  static const lightGreen50 =
      ExcelColor._('FFF1F8E9', 'lightGreen50', ColorType.material);
  static const lightGreen100 =
      ExcelColor._('FFDCEDC8', 'lightGreen100', ColorType.material);
  static const lightGreen200 =
      ExcelColor._('FFC5E1A5', 'lightGreen200', ColorType.material);
  static const lightGreen300 =
      ExcelColor._('FFAED581', 'lightGreen300', ColorType.material);
  static const lightGreen400 =
      ExcelColor._('FF9CCC65', 'lightGreen400', ColorType.material);
  static const lightGreen600 =
      ExcelColor._('FF7CB342', 'lightGreen600', ColorType.material);
  static const lightGreen700 =
      ExcelColor._('FF689F38', 'lightGreen700', ColorType.material);
  static const lightGreen800 =
      ExcelColor._('FF558B2F', 'lightGreen800', ColorType.material);
  static const lightGreen900 =
      ExcelColor._('FF33691E', 'lightGreen900', ColorType.material);
  static const lime50 = ExcelColor._('FFF9FBE7', 'lime50', ColorType.material);
  static const lime100 =
      ExcelColor._('FFF0F4C3', 'lime100', ColorType.material);
  static const lime200 =
      ExcelColor._('FFE6EE9C', 'lime200', ColorType.material);
  static const lime300 =
      ExcelColor._('FFDCE775', 'lime300', ColorType.material);
  static const lime400 =
      ExcelColor._('FFD4E157', 'lime400', ColorType.material);
  static const lime600 =
      ExcelColor._('FFC0CA33', 'lime600', ColorType.material);
  static const lime700 =
      ExcelColor._('FFAFB42B', 'lime700', ColorType.material);
  static const lime800 =
      ExcelColor._('FF9E9D24', 'lime800', ColorType.material);
  static const lime900 =
      ExcelColor._('FF827717', 'lime900', ColorType.material);
  static const yellow50 =
      ExcelColor._('FFFFFDE7', 'yellow50', ColorType.material);
  static const yellow100 =
      ExcelColor._('FFFFF9C4', 'yellow100', ColorType.material);
  static const yellow200 =
      ExcelColor._('FFFFF59D', 'yellow200', ColorType.material);
  static const yellow300 =
      ExcelColor._('FFFFF176', 'yellow300', ColorType.material);
  static const yellow400 =
      ExcelColor._('FFFFEE58', 'yellow400', ColorType.material);
  static const yellow600 =
      ExcelColor._('FFFDD835', 'yellow600', ColorType.material);
  static const yellow700 =
      ExcelColor._('FFFBC02D', 'yellow700', ColorType.material);
  static const yellow800 =
      ExcelColor._('FFF9A825', 'yellow800', ColorType.material);
  static const yellow900 =
      ExcelColor._('FFF57F17', 'yellow900', ColorType.material);
  static const amber50 =
      ExcelColor._('FFFFF8E1', 'amber50', ColorType.material);
  static const amber100 =
      ExcelColor._('FFFFECB3', 'amber100', ColorType.material);
  static const amber200 =
      ExcelColor._('FFFFE082', 'amber200', ColorType.material);
  static const amber300 =
      ExcelColor._('FFFFD54F', 'amber300', ColorType.material);
  static const amber400 =
      ExcelColor._('FFFFCA28', 'amber400', ColorType.material);
  static const amber600 =
      ExcelColor._('FFFFB300', 'amber600', ColorType.material);
  static const amber700 =
      ExcelColor._('FFFFA000', 'amber700', ColorType.material);
  static const amber800 =
      ExcelColor._('FFFF8F00', 'amber800', ColorType.material);
  static const amber900 =
      ExcelColor._('FFFF6F00', 'amber900', ColorType.material);
  static const orange50 =
      ExcelColor._('FFFFF3E0', 'orange50', ColorType.material);
  static const orange100 =
      ExcelColor._('FFFFE0B2', 'orange100', ColorType.material);
  static const orange200 =
      ExcelColor._('FFFFCC80', 'orange200', ColorType.material);
  static const orange300 =
      ExcelColor._('FFFFB74D', 'orange300', ColorType.material);
  static const orange400 =
      ExcelColor._('FFFFA726', 'orange400', ColorType.material);
  static const orange600 =
      ExcelColor._('FFFB8C00', 'orange600', ColorType.material);
  static const orange700 =
      ExcelColor._('FFF57C00', 'orange700', ColorType.material);
  static const orange800 =
      ExcelColor._('FFEF6C00', 'orange800', ColorType.material);
  static const orange900 =
      ExcelColor._('FFE65100', 'orange900', ColorType.material);
  static const deepOrange50 =
      ExcelColor._('FFFBE9E7', 'deepOrange50', ColorType.material);
  static const deepOrange100 =
      ExcelColor._('FFFFCCBC', 'deepOrange100', ColorType.material);
  static const deepOrange200 =
      ExcelColor._('FFFFAB91', 'deepOrange200', ColorType.material);
  static const deepOrange300 =
      ExcelColor._('FFFF8A65', 'deepOrange300', ColorType.material);
  static const deepOrange400 =
      ExcelColor._('FFFF7043', 'deepOrange400', ColorType.material);
  static const deepOrange600 =
      ExcelColor._('FFF4511E', 'deepOrange600', ColorType.material);
  static const deepOrange700 =
      ExcelColor._('FFE64A19', 'deepOrange700', ColorType.material);
  static const deepOrange800 =
      ExcelColor._('FFD84315', 'deepOrange800', ColorType.material);
  static const deepOrange900 =
      ExcelColor._('FFBF360C', 'deepOrange900', ColorType.material);
  static const brown50 =
      ExcelColor._('FFEFEBE9', 'brown50', ColorType.material);
  static const brown100 =
      ExcelColor._('FFD7CCC8', 'brown100', ColorType.material);
  static const brown200 =
      ExcelColor._('FFBCAAA4', 'brown200', ColorType.material);
  static const brown300 =
      ExcelColor._('FFA1887F', 'brown300', ColorType.material);
  static const brown400 =
      ExcelColor._('FF8D6E63', 'brown400', ColorType.material);
  static const brown600 =
      ExcelColor._('FF6D4C41', 'brown600', ColorType.material);
  static const brown700 =
      ExcelColor._('FF5D4037', 'brown700', ColorType.material);
  static const brown800 =
      ExcelColor._('FF4E342E', 'brown800', ColorType.material);
  static const brown900 =
      ExcelColor._('FF3E2723', 'brown900', ColorType.material);
  static const grey50 = ExcelColor._('FFFAFAFA', 'grey50', ColorType.material);
  static const grey100 =
      ExcelColor._('FFF5F5F5', 'grey100', ColorType.material);
  static const grey200 =
      ExcelColor._('FFEEEEEE', 'grey200', ColorType.material);
  static const grey300 =
      ExcelColor._('FFE0E0E0', 'grey300', ColorType.material);
  static const grey350 =
      ExcelColor._('FFD6D6D6', 'grey350', ColorType.material);
  static const grey400 =
      ExcelColor._('FFBDBDBD', 'grey400', ColorType.material);
  static const grey600 =
      ExcelColor._('FF757575', 'grey600', ColorType.material);
  static const grey700 =
      ExcelColor._('FF616161', 'grey700', ColorType.material);
  static const grey800 =
      ExcelColor._('FF424242', 'grey800', ColorType.material);
  static const grey850 =
      ExcelColor._('FF303030', 'grey850', ColorType.material);
  static const grey900 =
      ExcelColor._('FF212121', 'grey900', ColorType.material);
  static const blueGrey50 =
      ExcelColor._('FFECEFF1', 'blueGrey50', ColorType.material);
  static const blueGrey100 =
      ExcelColor._('FFCFD8DC', 'blueGrey100', ColorType.material);
  static const blueGrey200 =
      ExcelColor._('FFB0BEC5', 'blueGrey200', ColorType.material);
  static const blueGrey300 =
      ExcelColor._('FF90A4AE', 'blueGrey300', ColorType.material);
  static const blueGrey400 =
      ExcelColor._('FF78909C', 'blueGrey400', ColorType.material);
  static const blueGrey600 =
      ExcelColor._('FF546E7A', 'blueGrey600', ColorType.material);
  static const blueGrey700 =
      ExcelColor._('FF455A64', 'blueGrey700', ColorType.material);
  static const blueGrey800 =
      ExcelColor._('FF37474F', 'blueGrey800', ColorType.material);
  static const blueGrey900 =
      ExcelColor._('FF263238', 'blueGrey900', ColorType.material);

  static List<ExcelColor> get values => [
        black,
        black12,
        black26,
        black38,
        black45,
        black54,
        black87,
        white,
        white10,
        white12,
        white24,
        white30,
        white38,
        white54,
        white60,
        white70,
        redAccent,
        pinkAccent,
        purpleAccent,
        deepPurpleAccent,
        indigoAccent,
        blueAccent,
        lightBlueAccent,
        cyanAccent,
        tealAccent,
        greenAccent,
        lightGreenAccent,
        limeAccent,
        yellowAccent,
        amberAccent,
        orangeAccent,
        deepOrangeAccent,
        red,
        pink,
        purple,
        deepPurple,
        indigo,
        blue,
        lightBlue,
        cyan,
        teal,
        green,
        lightGreen,
        lime,
        yellow,
        amber,
        orange,
        deepOrange,
        brown,
        grey,
        blueGrey,
        redAccent100,
        redAccent400,
        redAccent700,
        pinkAccent100,
        pinkAccent400,
        pinkAccent700,
        purpleAccent100,
        purpleAccent400,
        purpleAccent700,
        deepPurpleAccent100,
        deepPurpleAccent400,
        deepPurpleAccent700,
        indigoAccent100,
        indigoAccent400,
        indigoAccent700,
        blueAccent100,
        blueAccent400,
        blueAccent700,
        lightBlueAccent100,
        lightBlueAccent400,
        lightBlueAccent700,
        cyanAccent100,
        cyanAccent400,
        cyanAccent700,
        tealAccent100,
        tealAccent400,
        tealAccent700,
        greenAccent100,
        greenAccent400,
        greenAccent700,
        lightGreenAccent100,
        lightGreenAccent400,
        lightGreenAccent700,
        limeAccent100,
        limeAccent400,
        limeAccent700,
        yellowAccent100,
        yellowAccent400,
        yellowAccent700,
        amberAccent100,
        amberAccent400,
        amberAccent700,
        orangeAccent100,
        orangeAccent400,
        orangeAccent700,
        deepOrangeAccent100,
        deepOrangeAccent400,
        deepOrangeAccent700,
        red50,
        red100,
        red200,
        red300,
        red400,
        red600,
        red700,
        red800,
        red900,
        pink50,
        pink100,
        pink200,
        pink300,
        pink400,
        pink600,
        pink700,
        pink800,
        pink900,
        purple50,
        purple100,
        purple200,
        purple300,
        purple400,
        purple600,
        purple700,
        purple800,
        purple900,
        deepPurple50,
        deepPurple100,
        deepPurple200,
        deepPurple300,
        deepPurple400,
        deepPurple600,
        deepPurple700,
        deepPurple800,
        deepPurple900,
        indigo50,
        indigo100,
        indigo200,
        indigo300,
        indigo400,
        indigo600,
        indigo700,
        indigo800,
        indigo900,
        blue50,
        blue100,
        blue200,
        blue300,
        blue400,
        blue600,
        blue700,
        blue800,
        blue900,
        lightBlue50,
        lightBlue100,
        lightBlue200,
        lightBlue300,
        lightBlue400,
        lightBlue600,
        lightBlue700,
        lightBlue800,
        lightBlue900,
        cyan50,
        cyan100,
        cyan200,
        cyan300,
        cyan400,
        cyan600,
        cyan700,
        cyan800,
        cyan900,
        teal50,
        teal100,
        teal200,
        teal300,
        teal400,
        teal600,
        teal700,
        teal800,
        teal900,
        green50,
        green100,
        green200,
        green300,
        green400,
        green600,
        green700,
        green800,
        green900,
        lightGreen50,
        lightGreen100,
        lightGreen200,
        lightGreen300,
        lightGreen400,
        lightGreen600,
        lightGreen700,
        lightGreen800,
        lightGreen900,
        lime50,
        lime100,
        lime200,
        lime300,
        lime400,
        lime600,
        lime700,
        lime800,
        lime900,
        yellow50,
        yellow100,
        yellow200,
        yellow300,
        yellow400,
        yellow600,
        yellow700,
        yellow800,
        yellow900,
        amber50,
        amber100,
        amber200,
        amber300,
        amber400,
        amber600,
        amber700,
        amber800,
        amber900,
        orange50,
        orange100,
        orange200,
        orange300,
        orange400,
        orange600,
        orange700,
        orange800,
        orange900,
        deepOrange50,
        deepOrange100,
        deepOrange200,
        deepOrange300,
        deepOrange400,
        deepOrange600,
        deepOrange700,
        deepOrange800,
        deepOrange900,
        brown50,
        brown100,
        brown200,
        brown300,
        brown400,
        brown600,
        brown700,
        brown800,
        brown900,
        grey50,
        grey100,
        grey200,
        grey300,
        grey350,
        grey400,
        grey600,
        grey700,
        grey800,
        grey850,
        grey900,
        blueGrey50,
        blueGrey100,
        blueGrey200,
        blueGrey300,
        blueGrey400,
        blueGrey600,
        blueGrey700,
        blueGrey800,
        blueGrey900,
      ];

  static Map<String, ExcelColor> get valuesAsMap =>
      values.asMap().map((_, v) => MapEntry(v.colorHex, v));
  @override
  List<Object?> get props => [
        _name,
        _color,
        _type,
        colorHex,
        colorInt,
      ];
}

enum ColorType {
  color,
  material,
  materialAccent,
  ;
}



================================================
FILE: lib/src/utilities/constants.dart
================================================
part of excel;

const _relationshipsStyles =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

const _relationshipsWorksheet =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

const _relationshipsSharedStrings =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

const _relationships =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const _spreasheetXlsx = 'xlsx';

// reference: https://support.microsoft.com/en-gb/office/change-the-column-width-and-row-height-72f5e3cc-994d-43e8-ae58-9774a0905f46
const _excelDefaultColumnWidth = 8.43;
const _excelDefaultRowHeight = 15.0;

const _newSheet =
    'UEsDBBQACAgIAPwDN1AAAAAAAAAAAAAAAAAYAAAAeGwvZHJhd2luZ3MvZHJhd2luZzEueG1sndBdbsIwDAfwE+wOVd5pWhgTQxRe0E4wDuAlbhuRj8oOo9x+0Uo2aXsBHm3LP/nvzW50tvhEYhN8I+qyEgV6FbTxXSMO72+zlSg4gtdgg8dGXJDFbvu0GTWtz7ynIu17XqeyEX2Mw1pKVj064DIM6NO0DeQgppI6qQnOSXZWzqvqRfJACJp7xLifJuLqwQOaA+Pz/k3XhLY1CvdBnRz6OCGEFmL6Bfdm4KypB65RPVD8AcZ/gjOKAoc2liq46ynZSEL9PAk4/hr13chSvsrVX8jdFMcBHU/DLLlDesiHsSZevpNlRnfugbdoAx2By8i4OPjj3bEqyTa1KCtssV7ercyzIrdfUEsHCAdiaYMFAQAABwMAAFBLAwQUAAgICAD8AzdQAAAAAAAAAAAAAAAAGAAAAHhsL3dvcmtzaGVldHMvc2hlZXQxLnhtbJ2TzW7DIAyAn2DvEHFvaLZ2W6Mklbaq2m5TtZ8zI06DCjgC0qRvP5K20bpeot2MwZ8/gUmWrZLBHowVqFMShVMSgOaYC71Nycf7evJIAuuYzplEDSk5gCXL7CZp0OxsCeACD9A2JaVzVUyp5SUoZkOsQPudAo1izi/NltrKAMv7IiXp7XR6TxUTmhwJsRnDwKIQHFbIawXaHSEGJHNe35aismeaaq9wSnCDFgsXclQnkjfgFFoOvdDjhZDiY4wUM7u6mnhk5S2+hRTu0HsNmH1KaqPjE2MyaHQ1se8f75U8H26j2Tjvq8tc0MWFfRvN/0eKpjSK/qBm7PouxmsxPpDUOMzwIqcRyZIe+WayBGsnhYY3E9ha+cs/PIHEJiV+cE+JjdiWrkvQLKFDXR98CmjsrzjoxvgbcdctXvOLot9n1/2D+568tg7VCxxbRCTIoWC1dM8ov0TuSp+bhbO7Ib/BZjg8Dx/mHb4nrphjPs4Na/xXC0wsfHfzmke9wPC7sh9QSwcILzuxOoEBAAChAwAAUEsDBBQACAgIAPwDN1AAAAAAAAAAAAAAAAAjAAAAeGwvd29ya3NoZWV0cy9fcmVscy9zaGVldDEueG1sLnJlbHONz0sKwjAQBuATeIcwe5PWhYg07UaEbqUeYEimD2weJPHR25uNouDC5czPfMNfNQ8zsxuFODkroeQFMLLK6ckOEs7dcb0DFhNajbOzJGGhCE29qk40Y8o3cZx8ZBmxUcKYkt8LEdVIBiN3nmxOehcMpjyGQXhUFxxIbIpiK8KnAfWXyVotIbS6BNYtnv6xXd9Pig5OXQ3Z9OOF0AHvuVgmMQyUJHD+2r3DkmcWRF2Jr4r1E1BLBwitqOtNswAAACoBAABQSwMEFAAICAgA/AM3UAAAAAAAAAAAAAAAABMAAAB4bC90aGVtZS90aGVtZTEueG1szVfbbtwgEP2C/gPivcHXvSm7UbKbVR9aVeq26jOx8aXB2AI2af6+GHttfEuiZiNlXwLjM4czM8CQy6u/GQUPhIs0Z2toX1gQEBbkYcriNfz1c/95AYGQmIWY5oys4RMR8Grz6RKvZEIyApQ7Eyu8homUxQohESgzFhd5QZj6FuU8w1JNeYxCjh8VbUaRY1kzlOGUwdqfv8Y/j6I0ILs8OGaEyYqEE4qlki6StBAQMJwpjYeEECng5iTylpLSQ5SGgPJDoJUPsOG9Xf4RPL7bUg4eMF1DS/8g2lyiBkDlELfXvxpXA8J75yU+p+Ib4np8GoCDQEUxXNtzFv7eq7EGqBoOuW+vPdf1O3iD3x1qubnZWl1+t8V7A7zrXS98t4P3Wrw/EutsZ9kdvN/iZ8N4Zze77ayD16CEpux+gLZt399ua3QDiXL65WV4i0LGzqn8mZzaRxn+k/O9Aujiqu3JgHwqSIQDhbvmKaYlPV4RPG4PxJgd9YizlL3TKi0xMgPVYWfdqL/rI6mjjlJKD/KJkq9CSxI5TcO9MuqJdmqSXCRqWC/XwcUc6zHgufydyuSQ4EItY+sVYlFTxwIUuVCHCU5y66Qcs295eCrr6dwpByxbu+U3dpVCWVln8/aQNvR6FgtTgK9JXy/CWKwrwh0RMXdfJ8K2zqViOaJiYT+nAhlVUQcF4LJr+F6lCIgAUxKWdar8T9U9e6WnktkN2xkJb+mdrdIdEcZ264owtmGCQ9I3n7nWy+V4qZ1RGfPFe9QaDe8Gyroz8KjOnOsrmgAXaxip60wNs0LxCRZDgGmsHieBrBP9PzdLwYXcYZFUMP2pij9LJeGAppna62YZKGu12c7c+rjiltbHyxzqF5lEEQnkhKWdqm8VyejXN4LLSX5Uog9J+Aju6JH/wCpR/twuEximQjbZDFNubO42i73rqj6KIy88/YChRYLrjmJe5hVcjxs5RhxaaT8qNJbCu3h/jq77slPv0pxoIPPJW+z9mryhyh1X5Y/edcuF9XyXeHtDMKQtxqW549KmescZHwTGcrOJvDmT1XxjN+jvWmS8K/Ws90/bybL5B1BLBwhlo4FhKAMAAK0OAABQSwMEFAAICAgA/AM3UAAAAAAAAAAAAAAAABQAAAB4bC9zaGFyZWRTdHJpbmdzLnhtbA3LQQ7CIBBA0RN4BzJ7C7owxpR21xPoASZlLCQwEGZi9Pay/Hn58/ot2XyoS6rs4TI5MMR7DYkPD6/ndr6DEUUOmCuThx8JrMtpFlEzVhYPUbU9rJU9UkGZaiMe8q69oI7sh5XWCYNEIi3ZXp272YKJwS5/UEsHCK+9gnR0AAAAgAAAAFBLAwQUAAgICAD8AzdQAAAAAAAAAAAAAAAADQAAAHhsL3N0eWxlcy54bWylU01v3CAQ/QX9D4h7FieKqiayHeXiKpf2kK3UK8awRgHGAja1++s7gPdLG6mVygXmzfBm3jDUT7M15F36oME19HZTUSKdgEG7XUN/bLubL5SEyN3ADTjZ0EUG+tR+qkNcjHwdpYwEGVxo6Bjj9MhYEKO0PGxgkg49CrzlEU2/Y2Hykg8hXbKG3VXVZ2a5drQwPM6391xc8VgtPARQcSPAMlBKC3nN9MAeGBcHJntN80E5lvu3/XSDtBOPutdGxyVXRdtagYuBCNi7iF1ZgbYOv8k7N4hU2CjW1gIMeOJ3fUO7rsorwY5bWQKfveYmQawQ5C0gnTbmyH9HC9DWWEiU3nVokPW8XSZsu8PmF5oc95doo3dj/Or5cnYlb5i5Bz/gc59rK1AKXZ0oTBrzmp74p7oInRUpMS9DQ3FWEunhiMrWo9vbzh4MPk1mecaSnJWFpkAdFCvlPU9Xkv9/3ln9YwFtzQ9OksYKR/97SpUvh9Fr97aFTsds41eJWqSn7SFGsJT88nzayjm7k5ZZrYKOWrKyCzlH9FRlmpmGfkvzaSjp99pE7YrvokPIOcyn5hTv6Te2fwBQSwcIzh0LebYBAADSAwAAUEsDBBQACAgIAPwDN1AAAAAAAAAAAAAAAAAPAAAAeGwvd29ya2Jvb2sueG1snZJLbsIwEIZP0DtE3oNjRCuISNhUldhUldoewNgTYuFHZJs03L6TkESibKKu/JxvPtn/bt8anTTgg3I2J2yZkgSscFLZU06+v94WG5KEyK3k2lnIyRUC2RdPux/nz0fnzgnW25CTKsY6ozSICgwPS1eDxZPSecMjLv2JhtoDl6ECiEbTVZq+UMOVJTdC5ucwXFkqAa9OXAzYeIN40DyifahUHUaaaR9wRgnvgivjUjgzkNBAUGgF9EKbOyEj5hgZ7s+XeoHIGi2OSqt47b0mTJOTi7fZwFhMGl1Nhv2zxujxcsvW87wfHnNLt3f2LXv+H4mllLE/qDV/fIv5WlxMJDMPM/3IEJFiituHp8Wu54dh7NIZMZiNCuqogSSWG1x+dmcMs9uNB4nRJonPFE78Qa4JUuiIkVAqC/Id6wLuC65F34aOTYtfUEsHCE3Koq1HAQAAJgMAAFBLAwQUAAgICAD8AzdQAAAAAAAAAAAAAAAAGgAAAHhsL19yZWxzL3dvcmtib29rLnhtbC5yZWxzrZJBasMwEEVP0DuI2deyk1JKiZxNKGTbpgcQ0tgysSUhTdr69p024DoQQhdeif/F/P/QaLP9GnrxgSl3wSuoihIEehNs51sF74eX+ycQmbS3ug8eFYyYYVvfbV6x18Qz2XUxCw7xWYEjis9SZuNw0LkIET3fNCENmlimVkZtjrpFuSrLR5nmGVBfZIq9VZD2tgJxGCP+Jzs0TWdwF8xpQE9XKiTxLHKgTi2Sgl95NquCw0BeZ1gtyZBp7PkNJ4izvlW/XrTe6YT2jRIveE4xt2/BPCwJ8xnSMTtE+gOZrB9UPqbFyIsfV38DUEsHCJYZwVPqAAAAuQIAAFBLAwQUAAgICAD8AzdQAAAAAAAAAAAAAAAACwAAAF9yZWxzLy5yZWxzjc9BDoIwEAXQE3iHZvZScGGMobAxJmwNHqC2QyFAp2mrwu3tUo0Ll5P5836mrJd5Yg/0YSAroMhyYGgV6cEaAdf2vD0AC1FaLSeyKGDFAHW1KS84yZhuQj+4wBJig4A+RnfkPKgeZxkycmjTpiM/y5hGb7iTapQG+S7P99y/G1B9mKzRAnyjC2Dt6vAfm7puUHgidZ/Rxh8VX4kkS28wClgm/iQ/3ojGLKHAq5J/PFi9AFBLBwikb6EgsgAAACgBAABQSwMEFAAICAgA/AM3UAAAAAAAAAAAAAAAABMAAABbQ29udGVudF9UeXBlc10ueG1stVPLTsMwEPwC/iHyFTVuOSCEmvbA4whIlA9Y7E1j1S953dffs0laJKoggdRevLbHOzPrtafznbPFBhOZ4CsxKceiQK+CNn5ZiY/F8+hOFJTBa7DBYyX2SGI+u5ou9hGp4GRPlWhyjvdSkmrQAZUhomekDslB5mVayghqBUuUN+PxrVTBZ/R5lFsOMZs+Yg1rm4uHfr+lrgTEaI2CzL4kk4niacdgb7Ndyz/kbbw+MTM6GCkT2u4MNSbS9akAo9QqvPLNJKPxXxKhro1CHdTacUpJMSFoahCzs+U2pFU37zXfIOUXcEwqd1Z+gyS7MCkPlZ7fBzWQUL/nxI2mIS8/DpzTh06wZc4hzQNEx8kl6897i8OFd8g5lTN/CxyS6oB+vGirOZYOjP/tzX2GsDrqy+5nz74AUEsHCG2ItFA1AQAAGQQAAFBLAQIUABQACAgIAPwDN1AHYmmDBQEAAAcDAAAYAAAAAAAAAAAAAAAAAAAAAAB4bC9kcmF3aW5ncy9kcmF3aW5nMS54bWxQSwECFAAUAAgICAD8AzdQLzuxOoEBAAChAwAAGAAAAAAAAAAAAAAAAABLAQAAeGwvd29ya3NoZWV0cy9zaGVldDEueG1sUEsBAhQAFAAICAgA/AM3UK2o602zAAAAKgEAACMAAAAAAAAAAAAAAAAAEgMAAHhsL3dvcmtzaGVldHMvX3JlbHMvc2hlZXQxLnhtbC5yZWxzUEsBAhQAFAAICAgA/AM3UGWjgWEoAwAArQ4AABMAAAAAAAAAAAAAAAAAFgQAAHhsL3RoZW1lL3RoZW1lMS54bWxQSwECFAAUAAgICAD8AzdQr72CdHQAAACAAAAAFAAAAAAAAAAAAAAAAAB/BwAAeGwvc2hhcmVkU3RyaW5ncy54bWxQSwECFAAUAAgICAD8AzdQzh0LebYBAADSAwAADQAAAAAAAAAAAAAAAAA1CAAAeGwvc3R5bGVzLnhtbFBLAQIUABQACAgIAPwDN1BNyqKtRwEAACYDAAAPAAAAAAAAAAAAAAAAACYKAAB4bC93b3JrYm9vay54bWxQSwECFAAUAAgICAD8AzdQlhnBU+oAAAC5AgAAGgAAAAAAAAAAAAAAAACqCwAAeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHNQSwECFAAUAAgICAD8AzdQpG+hILIAAAAoAQAACwAAAAAAAAAAAAAAAADcDAAAX3JlbHMvLnJlbHNQSwECFAAUAAgICAD8AzdQbYi0UDUBAAAZBAAAEwAAAAAAAAAAAAAAAADHDQAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLBQYAAAAACgAKAJoCAAA9DwAAAAA=';



================================================
FILE: lib/src/utilities/enum.dart
================================================
part of excel;

///enum for `wrapping` up the text
///
enum TextWrapping {
  WrapText,
  Clip,
}

///
///enum for setting `vertical alignment`
///
enum VerticalAlign {
  Top,
  Center,
  Bottom,
}

///
///enum for setting `horizontal alignment`
///
enum HorizontalAlign {
  Left,
  Center,
  Right,
}

///
///`Underline`
///
enum Underline {
  None,
  Single,
  Double,
}

///
///enum for setting `font scheme`
///
enum FontScheme { Unset, Major, Minor }



================================================
FILE: lib/src/utilities/fast_list.dart
================================================
part of excel;

// A helper class to optimized the usage of Maps
class FastList<K> {
  Map<K, int> _map = <K, int>{};
  int _index = 0;

  FastList();

  FastList.from(FastList<K> other)
      : _map = Map<K, int>.from(other._map),
        _index = other._index;

  void add(K key) {
    if (_map[key] == null) {
      _map[key] = _index;
      _index += 1;
    }
  }

  bool contains(K key) {
    return _map[key] != null;
  }

  void remove(K key) {
    _map.remove(key);
  }

  void clear() {
    _index = 0;
    _map = <K, int>{};
  }

  List<K> get keys => _map.keys.toList();

  bool get isNotEmpty => _map.isNotEmpty;
}



================================================
FILE: lib/src/utilities/span.dart
================================================
part of excel;

// For Spanning the columns and rows
class _Span extends Equatable {
  final int rowSpanStart;
  final int columnSpanStart;
  final int rowSpanEnd;
  final int columnSpanEnd;

  _Span({
    required this.rowSpanStart,
    required this.columnSpanStart,
    required this.rowSpanEnd,
    required this.columnSpanEnd,
  });

  _Span.fromCellIndex({
    required CellIndex start,
    required CellIndex end,
  })  : rowSpanStart = start.rowIndex,
        columnSpanStart = start.columnIndex,
        rowSpanEnd = end.rowIndex,
        columnSpanEnd = end.columnIndex;

  @override
  List<Object?> get props => [
        rowSpanStart,
        columnSpanStart,
        rowSpanEnd,
        columnSpanEnd,
      ];
}



================================================
FILE: lib/src/utilities/utility.dart
================================================
part of excel;

final List<String> _noCompression = <String>[
  'mimetype',
  'Thumbnails/thumbnail.png'
];

String getCellId(int columnIndex, int rowIndex) {
  return '${_numericToLetters(columnIndex + 1)}${rowIndex + 1}';
}

String _isColorAppropriate(String value) {
  switch (value.length) {
    case 7:
      return value.replaceAll(RegExp(r'#'), 'FF');
    case 9:
      return value.replaceAll(RegExp(r'#'), '');
    default:
      return value;
  }
}

/// Convert a character based column
int lettersToNumeric(String letters) {
  var sum = 0, mul = 1, n = 1;
  for (var index = letters.length - 1; index >= 0; index--) {
    var c = letters[index].codeUnitAt(0);
    n = 1;
    if (65 <= c && c <= 90) {
      n += c - 65;
    } else if (97 <= c && c <= 122) {
      n += c - 97;
    }
    sum += n * mul;
    mul = mul * 26;
  }
  return sum;
}

Iterable<XmlElement> _findRows(XmlElement table) {
  return table.findElements('row');
}

Iterable<XmlElement> _findCells(XmlElement row) {
  return row.findElements('c');
}

int? _getCellNumber(XmlElement cell) {
  var r = cell.getAttribute('r');
  if (r == null) {
    return null;
  }
  return _cellCoordsFromCellId(r).$2;
}

int? _getRowNumber(XmlElement row) {
  return int.tryParse(row.getAttribute('r').toString());
}

int _checkPosition(List<CellStyle> list, CellStyle cellStyle) {
  return list.indexOf(cellStyle);
}

int _letterOnly(int rune) {
  if (65 <= rune && rune <= 90) {
    return rune;
  } else if (97 <= rune && rune <= 122) {
    return rune - 32;
  }
  return 0;
}

String _twoDigits(int n) {
  if (n > 9) {
    return '$n';
  }
  return '0$n';
}

/// Convert a number to character based column
String _numericToLetters(int number) {
  var letters = '';

  while (number != 0) {
    // Set remainder from 1..26
    var remainder = number % 26;

    if (remainder == 0) {
      remainder = 26;
    }

    // Convert the remainder to a character.
    var letter = String.fromCharCode(65 + remainder - 1);

    // Accumulate the column letters, right to left.
    letters = letter + letters;

    // Get the next order of magnitude using bit shift.
    number = (number - 1) ~/ 26;
  }

  return letters;
}

/// Normalize line
String _normalizeNewLine(String text) {
  return text.replaceAll('\r\n', '\n');
}

///
///Returns the coordinates from a cell name.
///
///       cellCoordsFromCellId("A2"); // returns [2, 1]
///       cellCoordsFromCellId("B3"); // returns [3, 2]
///
///It is useful to convert CellId to Indexing.
///
(int x, int y) _cellCoordsFromCellId(String cellId) {
  var letters = cellId.runes.map(_letterOnly);
  var lettersPart = utf8.decode(letters.where((rune) {
    return rune > 0;
  }).toList(growable: false));
  var numericsPart = cellId.substring(lettersPart.length);

  return (
    int.parse(numericsPart) - 1,
    lettersToNumeric(lettersPart) - 1
  ); // [x , y]
}

///
///Throw error at situation where further processing is not possible
///It is also called when important parts of excel files are missing as corrupted excel file is used
///
void _damagedExcel({String text = ''}) {
  throw ArgumentError('\nDamaged Excel file: $text\n');
}

///
///return A2:B2 for spanning storage in unmerge list when [0,2] [2,2] is passed
///
String getSpanCellId(int startColumn, int startRow, int endColumn, int endRow) {
  return '${getCellId(startColumn, startRow)}:${getCellId(endColumn, endRow)}';
}

///
///returns updated SpanObject location as there might be cross-sectional interaction between the two spanning objects.
///
(
  bool changeValue,
  (int startColumn, int startRow, int endColumn, int endRow)
) _isLocationChangeRequired(
    int startColumn, int startRow, int endColumn, int endRow, _Span spanObj) {
  bool changeValue = (
          // Overlapping checker
          startRow <= spanObj.rowSpanStart &&
              startColumn <= spanObj.columnSpanStart &&
              endRow >= spanObj.rowSpanEnd &&
              endColumn >= spanObj.columnSpanEnd)
      // first check starts here
      ||
      ( // outwards checking
          ((startColumn < spanObj.columnSpanStart &&
                      endColumn >= spanObj.columnSpanStart) ||
                  (startColumn <= spanObj.columnSpanEnd &&
                      endColumn > spanObj.columnSpanEnd))
              // inwards checking
              &&
              ((startRow >= spanObj.rowSpanStart &&
                      startRow <= spanObj.rowSpanEnd) ||
                  (endRow >= spanObj.rowSpanStart &&
                      endRow <= spanObj.rowSpanEnd)))

      // second check starts here
      ||
      (
          // outwards checking
          ((startRow < spanObj.rowSpanStart &&
                      endRow >= spanObj.rowSpanStart) ||
                  (startRow <= spanObj.rowSpanEnd &&
                      endRow > spanObj.rowSpanEnd))
              // inwards checking
              &&
              ((startColumn >= spanObj.columnSpanStart &&
                      startColumn <= spanObj.columnSpanEnd) ||
                  (endColumn >= spanObj.columnSpanStart &&
                      endColumn <= spanObj.columnSpanEnd)));

  if (changeValue) {
    if (startColumn > spanObj.columnSpanStart) {
      startColumn = spanObj.columnSpanStart;
    }
    if (endColumn < spanObj.columnSpanEnd) {
      endColumn = spanObj.columnSpanEnd;
    }
    if (startRow > spanObj.rowSpanStart) {
      startRow = spanObj.rowSpanStart;
    }
    if (endRow < spanObj.rowSpanEnd) {
      endRow = spanObj.rowSpanEnd;
    }
  }

  return (changeValue, (startColumn, startRow, endColumn, endRow));
}

///
///Returns Column based String alphabet when column index is passed
///
///     `getColumnAlphabet(0); // returns A`
///     `getColumnAlphabet(5); // returns F`
///
String getColumnAlphabet(int columnIndex) {
  return _numericToLetters(columnIndex + 1);
}

///
///Returns Column based int index when column alphabet is passed
///
///    `getColumnAlphabet("A"); // returns 0`
///    `getColumnAlphabet("F"); // returns 5`
///
int getColumnIndex(String columnAlphabet) {
  return _cellCoordsFromCellId(columnAlphabet).$2;
}

///
///Checks if the fontStyle is already present in the list or not
///
int _fontStyleIndex(List<_FontStyle> list, _FontStyle fontStyle) {
  return list.indexOf(fontStyle);
}



================================================
FILE: lib/src/web_helper/client_save_excel.dart
================================================
class SavingHelper {
// A wrapper to save the excel file in client
  static List<int>? saveFile(List<int>? val, String fileName) {
    return val;
  }
}



================================================
FILE: lib/src/web_helper/web_save_excel_browser.dart
================================================
import 'dart:js_interop';
import 'dart:typed_data';

import 'package:web/web.dart';

// A wrapper to save the excel file in browser
class SavingHelper {
  static List<int>? saveFile(List<int>? val, String fileName) {
    if (val == null) {
      return null;
    }

    final blob = Blob(JSArray.from(Uint8List.fromList(val).toJS));
    final url = URL.createObjectURL(blob);
    final anchor = HTMLAnchorElement()
      ..href = url
      ..download = '$fileName';

    document.body?.append(anchor);

    // download the file
    anchor.click();

    // cleanup
    anchor.remove();
    URL.revokeObjectURL(url);
    return val;
  }
}
