# セル結合がエクスポート時に壊れる不具合 — 原因分析レポート

## 概要

Dart Excel ライブラリにおいて、インポートした Excel ファイル内のセルに値を埋め込みエクスポートすると、インポート時に設定されていたセル結合（マージ）が解除されてしまう不具合について、根本原因を調査した。

---

## 1. ライブラリのアーキテクチャ概要

### データ構造

| 構造 | 場所 | 役割 |
|------|------|------|
| `_excel._archive` | `lib/src/excel.dart` | 元のZIPアーカイブ |
| `_excel._xmlFiles` | `Map<String, XmlDocument>` | パースされたXMLドキュメント群 |
| `_excel._sheets` | `Map<String, XmlNode>` | 各シートの `<sheetData>` 要素への参照 |
| `_excel._sheetMap` | `Map<String, Sheet>` | 各シートのインメモリデータモデル |
| `_excel._xmlSheetId` | `Map<String, String>` | シート名→XMLファイルパスのマッピング |

### マージ関連のデータ構造

| 構造 | 場所 | 役割 |
|------|------|------|
| `sheet._spannedItems` | `FastList<String>` (L4826) | マージ範囲文字列リスト（例: "A1:B2"） |
| `sheet._spanList` | `List<_Span?>` (L4827) | マージ範囲オブジェクトリスト |
| `_excel._mergeChanges` | `bool` (L110) | マージ変更フラグ |
| `_excel._mergeChangeLook` | `List<String>` (L123) | マージ変更があったシート名リスト |

### `_Span` クラス（L7371-7398）

```dart
class _Span extends Equatable {
  final int rowSpanStart;
  final int columnSpanStart;
  final int rowSpanEnd;
  final int columnSpanEnd;
}
```

---

## 2. 処理フローの詳細追跡

### 2.1 インポート時のフロー

**`_startParsing()`**（L1222-1229）の実行順序:

```
1. _putContentXml()         → [Content_Types].xml をパース
2. _parseRelations()        → xl/_rels/workbook.xml.rels をパース
3. _parseStyles()           → xl/styles.xml をパース
4. _parseSharedStrings()    → xl/sharedStrings.xml をパース
5. _parseContent()          → xl/workbook.xml をパース → 各シートの _parseTable() を呼出
6. _parseMergedCells()      → 各シートのマージ情報をパース
```

#### `_parseTable()`（L1738-1778）

```dart
var content = XmlDocument.parse(utf8.decode(file.content));  // L1751
var worksheet = content.findElements('worksheet').first;      // L1752
var sheet = worksheet.findElements('sheetData').first;        // L1763

_findRows(sheet).forEach((child) {
    _parseRow(child, sheetObject, name);  // 各セルのデータを _sheetData に格納
});

_excel._sheets[name] = sheet;                    // L1772: <sheetData> ノードを保存
_excel._xmlFiles['xl/$target'] = content;        // L1774: XmlDocument 全体を保存
_excel._xmlSheetId[name] = 'xl/$target';         // L1775: シート名→パスのマッピング
```

**重要**: `_excel._sheets[name]` は `<sheetData>` 要素（`<worksheet>` の子）への参照。`<mergeCells>` は `<sheetData>` の兄弟要素として `<worksheet>` の直下に存在する。

```xml
<worksheet>
  <sheetData>          ← _excel._sheets[name] が参照するノード
    <row>...</row>
  </sheetData>
  <mergeCells count="1">  ← <sheetData> の兄弟要素
    <mergeCell ref="A1:B2"/>
  </mergeCells>
</worksheet>
```

#### `_parseCell()`（L1791-1874）

各セルのパースで `sheetObject.updateCell()` を呼ぶ（L1869）。この時点では `_spanList` は空であるため、`_isInsideSpanning()` によるリダイレクトは行われず、全セルが元の位置に格納される。

#### `_parseMergedCells()`（L1383-1421）

```dart
void _parseMergedCells() {
  _excel._sheets.forEach((sheetName, node) {
    _excel._availSheet(sheetName);
    XmlElement sheetDataNode = node as XmlElement;
    final sheet = _excel._sheetMap[sheetName]!;
    final worksheetNode = sheetDataNode.parent;  // <worksheet> 要素

    worksheetNode!.findAllElements('mergeCell').forEach((element) {
      String? ref = element.getAttribute('ref');       // 例: "A1:B2"
      if (ref != null && ref.contains(':') && ref.split(':').length == 2) {
        sheet._spannedItems.add(ref);                  // L1396: 文字列リストに追加

        CellIndex startIndex = CellIndex.indexByString(startCell);
        CellIndex endIndex = CellIndex.indexByString(endCell);
        _Span spanObj = _Span.fromCellIndex(start: startIndex, end: endIndex);  // L1408

        sheet._spanList.add(spanObj);                  // L1413: Spanリストに追加
        _deleteAllButTopLeftCellsOfSpanObj(spanObj, sheet);  // L1415: 左上以外のセルを削除

        _excel._mergeChangeLookup = sheetName;         // L1417: ★ここが重要
      }
    });
  });
}
```

#### `_deleteAllButTopLeftCellsOfSpanObj()`（L1432-1449）

マージ範囲内の左上セル以外のデータを `_sheetData` から削除する:

```dart
void _deleteAllButTopLeftCellsOfSpanObj(_Span spanObj, Sheet sheet) {
  for (var columnI = columnSpanStart; columnI <= columnSpanEnd; columnI++) {
    for (var rowI = rowSpanStart; rowI <= rowSpanEnd; rowI++) {
      bool isTopLeftCellThatShouldNotBeDeleted =
          columnI == columnSpanStart && rowI == rowSpanStart;
      if (isTopLeftCellThatShouldNotBeDeleted) {
        continue;
      }
      sheet._removeCell(rowI, columnI);  // L1446
    }
  }
}
```

### 2.2 ユーザー操作時のフロー

ユーザーが `sheet.cell(CellIndex.indexByString("C1")).value = TextCellValue("hello")` を実行:

1. **`cell()`**（L4945-4984）: 指定位置に `Data` オブジェクトが存在しなければ新規作成。`_maxRows`/`_maxColumns` を更新。
2. **`value` setter**（L4019-4021）: `_sheet.updateCell(cellIndex, val)` を呼出。
3. **`updateCell()`**（L5502-5546）:
   - `_checkMaxColumn()` / `_checkMaxRow()` でサイズ更新
   - `_spanList.isNotEmpty` の場合、`_isInsideSpanning()` でマージ範囲内なら左上セルにリダイレクト
   - `_putData()` でデータ格納
   - **`_mergeChanges` は変更されない**

### 2.3 保存時のフロー

**`_save()`**（L2647-2671）:

```dart
List<int>? _save() {
    if (_excel._styleChanges) {
      _processStylesFile();           // スタイル変更時のみ
    }
    _setSheetElements();              // L2651: ★全シートの sheetData を再構築
    if (_excel._defaultSheet != null) {
      _setDefaultSheet(_excel._defaultSheet);
    }
    _setSharedStrings();              // L2655: 共有文字列を再構築

    if (_excel._mergeChanges) {       // L2657: ★ここが問題の条件分岐
      _setMerge();                    // L2658: マージ情報をXMLに書き戻す
    }

    if (_excel._rtlChanges) {
      _setRTL();
    }

    // 全XMLファイルをシリアライズしてアーカイブに格納
    for (var xmlFile in _excel._xmlFiles.keys) {
      var xml = _excel._xmlFiles[xmlFile].toString();
      var content = utf8.encode(xml);
      _archiveFiles[xmlFile] = ArchiveFile(xmlFile, content.length, content);
    }
    return ZipEncoder().encode(_cloneArchive(_excel._archive, _archiveFiles));
}
```

#### `_setSheetElements()`（L3012-3071）

```dart
void _setSheetElements() {
    _excel._sharedStrings.clear();              // L3013

    _excel._sheetMap.forEach((sheetName, sheetObject) {
      // 新規シートの場合XMLを作成
      if (_excel._sheets[sheetName] == null) {
        parser._createSheet(sheetName);          // L3019: ★<mergeCells>を含まない新XMLを作成
      }

      // <sheetData> の子要素（<row>）をすべてクリア
      if (_excel._sheets[sheetName]?.children.isNotEmpty ?? false) {
        _excel._sheets[sheetName]!.children.clear();  // L3026: ★<sheetData>の中身のみクリア
      }

      XmlDocument? xmlFile = _excel._xmlFiles[_excel._xmlSheetId[sheetName]];
      if (xmlFile == null) return;

      // sheetFormatPr の処理（L3038-3063）
      // ...

      _setColumns(sheetObject, xmlFile);         // L3065: 列幅の設定
      _setRows(sheetName, sheetObject);          // L3067: 行・セルデータの再構築
      _setHeaderFooter(sheetName);               // L3069: ヘッダー/フッターの設定
    });
}
```

#### `_setRows()`（L2731-2757）

```dart
void _setRows(String sheetName, Sheet sheetObject) {
    for (var rowIndex = 0; rowIndex < sheetObject._maxRows; rowIndex++) {
      if (sheetObject._sheetData[rowIndex] == null) {
        continue;  // データのない行はスキップ
      }
      var foundRow = _createNewRow(
          _excel._sheets[sheetName]! as XmlElement, rowIndex, height);
      for (var columnIndex = 0; columnIndex < sheetObject._maxColumns; columnIndex++) {
        var data = sheetObject._sheetData[rowIndex]![columnIndex];
        if (data == null) {
          continue;  // データのないセルはスキップ
        }
        _updateCell(sheetName, foundRow, columnIndex, rowIndex,
            data.value, data.cellStyle?.numberFormat);
      }
    }
}
```

#### `_setMerge()`（L2816-2884）— 呼ばれない場合の影響

```dart
void _setMerge() {
    _selfCorrectSpanMap(_excel);  // 重複マージの自動補正
    _excel._mergeChangeLook.forEach((s) {
      if (_excel._sheetMap[s] != null &&
          _excel._sheetMap[s]!._spanList.isNotEmpty &&
          _excel._xmlSheetId.containsKey(s) &&
          _excel._xmlFiles.containsKey(_excel._xmlSheetId[s])) {

        // <mergeCells> 要素を検索、なければ作成
        Iterable<XmlElement>? iterMergeElement = _excel
            ._xmlFiles[_excel._xmlSheetId[s]]
            ?.findAllElements('mergeCells');

        if (iterMergeElement?.isNotEmpty ?? false) {
          mergeElement = iterMergeElement!.first;       // 既存要素を使用
        } else {
          // <sheetData> の直後に新規 <mergeCells> を挿入（L2845-2856）
        }

        // _spanList から spannedItems を再導出
        List<String> _spannedItems =
            List<String>.from(_excel._sheetMap[s]!.spannedItems);

        // count 属性を更新
        mergeElement.getAttributeNode('count')!.value = _spannedItems.length.toString();

        // 既存の <mergeCell> 子要素をすべてクリアして再構築
        mergeElement.children.clear();
        _spannedItems.forEach((value) {
          mergeElement.children.add(XmlElement(XmlName('mergeCell'),
              [XmlAttribute(XmlName('ref'), '$value')], []));
        });
      }
    });
}
```

---

## 3. 根本原因の特定

### 主原因: `_mergeChangeLookup` セッター内の `_mergeChanges = true;` がコメントアウト

**場所: `lib/src/excel.dart` L658-663**

```dart
set _mergeChangeLookup(String value) {
    if (!_mergeChangeLook.contains(value)) {
      _mergeChangeLook.add(value);
      //_mergeChanges = true;   // ← ★ コメントアウトされている
    }
}
```

### `_mergeChanges` が `true` になるケース（現状）

`_mergeChanges = true` が設定されるのは以下の場合**のみ**:

| 操作 | 行番号 |
|------|--------|
| `removeColumn()` 内でスパン境界がシフトした場合 | L5176 |
| `insertColumn()` 内でスパン境界がシフトした場合 | L5263 |
| `removeRow()` 内でスパン境界がシフトした場合 | L5366 |
| `insertRow()` 内でスパン境界がシフトした場合 | L5448 |
| `merge()` が明示的に呼ばれた場合 | L5573 |

**単純なセル値の更新（`updateCell()`）では `_mergeChanges` は設定されない。**

### 結果として起きること

1. **インポート**: `_parseMergedCells()` がマージ情報を `_spanList` / `_spannedItems` に正しく格納。`_mergeChangeLook` にシート名を追加。しかし `_mergeChanges` は `false` のまま。
2. **ユーザー操作**: セル値を更新。`_mergeChanges` は変更されない。
3. **保存**: `if (_excel._mergeChanges)` が `false` → `_setMerge()` が呼ばれない。

### XML要素の保持について

`_setSheetElements()` は `_excel._sheets[sheetName]!.children.clear()` で `<sheetData>` 内の `<row>` 要素のみをクリアする。`<mergeCells>` は `<worksheet>` の直接の子要素であり `<sheetData>` の兄弟要素であるため、この操作では削除されない。

したがって、**最もシンプルなケース**（インポート → セル値変更 → エクスポート）では、`<mergeCells>` XML要素はXmlDocumentに残存し、シリアライズ時に出力される**可能性がある**。

しかし以下のケースでは確実に失われる:

#### ケース1: シートのリネーム

```dart
// rename() の内部処理
delete(oldSheetName);   // _sheets.remove(oldSheetName) が呼ばれる（L377）
```

リネーム後、新しいシート名は `_sheetMap` には存在するが `_sheets` には存在しない。
`_setSheetElements()` で `_excel._sheets[newName] == null` となり、`_createSheet(newName)` が呼ばれる。
`_createSheet()` は `<mergeCells>` を含まない新しいXMLを作成する（L1982）。
`_setMerge()` が呼ばれないため、マージ情報は復元されない。

#### ケース2: シートのコピー

```dart
void copy(String fromSheet, String toSheet) {
    _availSheet(toSheet);  // _sheetMap に追加されるが _sheets には追加されない
    this[toSheet] = this[fromSheet];  // データをクローン
}
```

コピー先シートは `_sheets` にエントリがないため、同様に `_createSheet()` が呼ばれ、マージ情報が失われる。

#### ケース3: `_setMerge()` が必要なのに呼ばれないケース

行・列の挿入/削除はマージのスパン境界を更新し `_mergeChanges = true` を設定するため、`_setMerge()` が呼ばれる。しかし、インポートしたマージ情報のみの場合は `_mergeChanges` が `false` のままなので、XMLドキュメントが何らかの理由で再構築された場合にマージ情報は失われる。

---

## 4. `_setMerge()` の動作詳細

`_setMerge()` が正しく呼ばれた場合の処理:

1. **`_selfCorrectSpanMap()`**（L3121-3182）: 重複するスパンを検出・統合
2. **`_mergeChangeLook` のシートを順次処理**:
   - `_spanList.isNotEmpty` のシートのみ処理
   - 既存の `<mergeCells>` 要素があれば取得、なければ `<sheetData>` の直後に新規作成
   - `spannedItems` ゲッター（L6219-6235）で `_spanList` から文字列リストを再導出
   - `<mergeCells>` の `count` 属性を更新
   - 既存の `<mergeCell>` 子要素をすべてクリア
   - `_spanList` の各エントリから `<mergeCell ref="..."/>` を再構築

### `spannedItems` ゲッター（L6219-6235）

```dart
List<String> get spannedItems {
  _spannedItems = FastList<String>();  // 毎回新規作成
  for (int i = 0; i < _spanList.length; i++) {
    _Span? spanObj = _spanList[i];
    if (spanObj == null) continue;
    String rC = getSpanCellId(spanObj.columnSpanStart, spanObj.rowSpanStart,
        spanObj.columnSpanEnd, spanObj.rowSpanEnd);
    if (!_spannedItems.contains(rC)) {
      _spannedItems.add(rC);
    }
  }
  return _spannedItems;
}
```

このゲッターは `_spanList` から毎回 `_spannedItems` を再導出する。`_spanList` がインポート時に正しく設定されていれば、`_setMerge()` が呼ばれた場合は正しいマージ情報がXMLに書き込まれる。

---

## 5. `_createSheet()` の問題点（L1905-2011）

新規シート作成時のXMLテンプレート（L1982）:

```xml
<worksheet xmlns="...">
  <dimension ref="A1"/>
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetData/>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
```

**`<mergeCells>` 要素は含まれていない**。このXMLが `_excel._xmlFiles` に設定されるため、元のXMLドキュメント（マージ情報を含む）は置き換えられる。

---

## 6. 副次的な問題点

### `_countRowsAndColumns()` の広範なコメントアウト

以下の箇所で `_countRowsAndColumns()` がコメントアウトされている:

| 場所 | メソッド | 行番号 |
|------|----------|--------|
| `insertColumn()` | 列挿入後 | L5323 |
| `removeRow()` | 行削除後 | L5397 |
| `insertRow()` | 行挿入後 | L5490 |
| `_putData()` | データ格納後 | L5931 |
| `clearRow()` | 行クリア後 | L6155 |

直接マージの問題には関係しないが、`_maxRows`/`_maxColumns` の不整合を引き起こす可能性があり、`_setRows()` のイテレーション範囲に影響を与えうる。

---

## 7. 修正方針

### 最小修正: L661 のコメントアウト解除

**ファイル: `lib/src/excel.dart`**

```dart
// 修正前（L658-663）
set _mergeChangeLookup(String value) {
    if (!_mergeChangeLook.contains(value)) {
      _mergeChangeLook.add(value);
      //_mergeChanges = true;
    }
}

// 修正後
set _mergeChangeLookup(String value) {
    if (!_mergeChangeLook.contains(value)) {
      _mergeChangeLook.add(value);
      _mergeChanges = true;
    }
}
```

### 修正の効果

- インポート時に `_parseMergedCells()` → `_mergeChangeLookup = sheetName` → `_mergeChanges = true`
- 保存時に `if (_excel._mergeChanges)` が `true` → `_setMerge()` が呼ばれる
- `_setMerge()` が `_spanList` からマージ情報を再構築してXMLに書き込む
- シートのリネーム・コピーで XMLドキュメントが置換されても、マージ情報が復元される

### 修正の影響範囲

- `_setMerge()` は `_selfCorrectSpanMap()` を呼び出し、全対象シートのスパンデータを処理する
- マージデータが大量にある場合、保存処理の時間が若干増加する可能性がある
- ただしこれは本来の正しい動作であり、コメントアウトされていること自体がバグである

### 代替案: `_save()` での無条件呼び出し

```dart
// 修正前
if (_excel._mergeChanges) {
    _setMerge();
}

// 修正後: 条件を削除して常に呼び出す
_setMerge();
```

より安全だが、マージデータのないファイルでも `_setMerge()` が実行されるため、不要な処理が発生する。`_setMerge()` 内部で `_mergeChangeLook` が空なら何もしないので実害は小さい。

---

## 8. 検証方法

1. マージされたセル（例: A1:B2）を含むExcelファイルを用意する
2. ライブラリでインポート:
   ```dart
   var bytes = File('template.xlsx').readAsBytesSync();
   var excel = Excel.decodeBytes(bytes);
   ```
3. 任意のセルに値を設定:
   ```dart
   var sheet = excel['Sheet1'];
   sheet.cell(CellIndex.indexByString("C1")).value = TextCellValue("test");
   ```
4. エクスポート:
   ```dart
   var output = excel.encode();
   File('output.xlsx').writeAsBytesSync(output!);
   ```
5. `output.xlsx` をExcelで開き、A1:B2 のセル結合が保持されていることを確認
6. 追加テスト: シートのリネーム・コピー後もマージが保持されることを確認
7. 既存テストスイートの実行で回帰がないことを確認

---

## 9. 関連ファイル一覧

| ファイル | 主な関連箇所 |
|----------|-------------|
| `lib/src/excel.dart` | L110(`_mergeChanges`), L123(`_mergeChangeLook`), L658-663(`_mergeChangeLookup` setter) |
| `lib/src/parser/parse.dart` | L1222-1229(`_startParsing`), L1383-1449(`_parseMergedCells`), L1738-1778(`_parseTable`) |
| `lib/src/save/save_file.dart` | L2647-2671(`_save`), L2816-2884(`_setMerge`), L3012-3071(`_setSheetElements`), L3121-3182(`_selfCorrectSpanMap`) |
| `lib/src/sheet/sheet.dart` | L4826-4827(マージデータ構造), L5502-5546(`updateCell`), L5549-5625(`merge`), L5627-5665(`unMerge`), L6219-6235(`spannedItems` getter) |
| `lib/src/utilities/span.dart` | L7371-7398(`_Span` class) |

※行番号はすべて `gitingest.md` 内の行番号を基準としている。
