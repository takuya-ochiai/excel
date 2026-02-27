# xlsmパススルーエクスポート XMLエラー調査報告・修正方針

## 現象

`test/export_test.dart` の「xlsmファイルを読み込み、変更せずそのままエクスポートする」テストで出力された `tmp/passthrough_out.xlsm` をExcelで開くとエラーが発生する。

```
置き換えられたパーツ: /xl/worksheets/sheet1.xml パーツに XML エラーがありました。
読み込みエラーが発生しました。場所は、行 2、列 0 です。
```

---

## 検証結果

### 問題1（XMLエラーの直接原因）: `<headerFooter>` 要素の順序違反

**場所**: `lib/src/save/save_file.dart` 690-707行目 `_setHeaderFooter()`

```dart
void _setHeaderFooter(String sheetName) {
    ...
    final results = sheetXmlElement.findAllElements("headerFooter");
    if (results.isNotEmpty) {
      sheetXmlElement.children.remove(results.first);  // 既存の<headerFooter>を削除
    }
    if (sheet.headerFooter == null) return;
    sheetXmlElement.children.add(sheet.headerFooter!.toXmlElement());  // ← 末尾に追加
}
```

**原因**: `children.add()` により `<headerFooter>` が `<extLst>` の**後**に配置される。Open XMLスキーマ（ISO/IEC 29500）では `<headerFooter>` は `<extLst>` より前でなければならない。

**要素順序の比較**:

```
元ファイル:                    エクスポート後:
  ...                            ...
  <pageSetup>                    <pageSetup>
  <headerFooter>  ← 正しい位置    <extLst>
  <extLst>                       <headerFooter>  ← スキーマ違反！
```

Open XMLスキーマの正しい順序: `... → pageMargins → pageSetup → headerFooter → ... → extLst`（`extLst`は必ず最後）

---

### 問題2: スタイルのみのセルが失われる（23,813セル消失）

**場所**: `lib/src/save/save_file.dart` 625-651行目 `_setRows()`

```dart
var data = sheetObject._sheetData[rowIndex]![columnIndex];
if (data == null) {
  continue;  // ← スタイルのみのセルがスキップされる
}
```

**数値**:
- 元ファイル: 83,853セル
- エクスポート後: 60,040セル
- **23,813セルが消失**

**消失セルの特徴**: すべて `type=none, val=false, formula=false, style=true`
（例: `<c r="E5" s="41"/>` — 値なし、スタイルのみ）

**原因**: パース時に値のないセルは `_sheetData` に格納されない。保存時に `data == null` でスキップされるため、スタイル情報だけを持つセルが消失する。

---

### 問題3: `<sheetViews>` の属性消失

```
元: <sheetView showGridLines="0" tabSelected="1" zoomScale="85" zoomScaleNormal="85" workbookViewId="0">
後: <sheetView workbookViewId="0"/>
```

多数の表示関連属性（グリッド線、ズーム等）と子要素が失われている。

---

### 問題4: `<sheetFormatPr>` の属性変化

```
元: defaultColWidth="9" defaultRowHeight="13" x14ac:dyDescent="0.2"
後: defaultRowHeight="13.00" defaultColWidth="9.00"
```

- `x14ac:dyDescent` 属性が消失
- 整数値に不要な小数点が追加される（`"9"` → `"9.00"`）

---

### 問題5: `<cols>` の子要素数増加

- 元ファイル: 45個のcol定義
- エクスポート後: 119個のcol定義

saveプロセスでカラム定義が再生成され、数が増加している。

---

## 修正方針

### 問題1の修正（XMLエラーの直接原因 — 最優先）

**ファイル**: `lib/src/save/save_file.dart` `_setHeaderFooter()` メソッド

**方針**: `children.add()`（末尾追加）ではなく、`<extLst>` の前に `insert` する。

```dart
// 修正案
void _setHeaderFooter(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    if (sheet == null) return;
    final xmlFile = _excel._xmlFiles[_excel._xmlSheetId[sheetName]];
    if (xmlFile == null) return;

    final sheetXmlElement = xmlFile.findAllElements("worksheet").first;

    // 既存の<headerFooter>を削除
    final results = sheetXmlElement.findAllElements("headerFooter");
    if (results.isNotEmpty) {
      sheetXmlElement.children.remove(results.first);
    }

    if (sheet.headerFooter == null) return;

    // <extLst>の前に挿入（Open XMLスキーマ順序を維持）
    var extLstIndex = sheetXmlElement.children.indexWhere(
      (child) => child is XmlElement && (child as XmlElement).name.local == 'extLst'
    );
    if (extLstIndex != -1) {
      sheetXmlElement.children.insert(extLstIndex, sheet.headerFooter!.toXmlElement());
    } else {
      sheetXmlElement.children.add(sheet.headerFooter!.toXmlElement());
    }
}
```

### 問題2〜5について

これらは「パススルー時のデータ忠実性」に関する問題であり、XMLエラーの直接原因ではない。save処理全体の設計に関わるため、問題1を先に修正し、影響範囲を確認した後に段階的に対応する。

---

## 検証方法

1. `dart test test/export_test.dart` を実行しテストがpassすることを確認
2. `tmp/passthrough_out.xlsm` をExcelで開き、**自動修復メッセージが表示されない**ことを確認
3. `tmp/structure.dart` で要素順序が正しいことを確認（`<headerFooter>` が `<extLst>` の前）

---

## 関連ファイル一覧

| ファイル | 主な関連箇所 |
|----------|-------------|
| `lib/src/save/save_file.dart` | L690-707(`_setHeaderFooter`), L625-651(`_setRows`), L541-565(`_save`), L906-965(`_setSheetElements`) |
| `lib/src/parser/parse.dart` | L536-562(ワークシートXMLパース), L607-612(共有文字列セルパース) |
| `lib/src/sharedStrings/shared_strings.dart` | L36-44(`addFromParsedXml`) |
| `test/export_test.dart` | テストコード |
| `test/test_resources/report-template.xlsm` | テスト用テンプレート |
