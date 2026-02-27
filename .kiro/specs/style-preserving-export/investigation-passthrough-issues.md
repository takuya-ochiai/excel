# パススルーエクスポート問題の調査結果

## 報告日: 2026-02-27

## 概要

`report-template.xlsm` を読み込み、変更せずにエクスポートした `passthrough_out.xlsm` で以下の3つの視覚的問題が発生している。

1. **一部のセル（結合セル含む）が黒塗りに変化**
2. **一部のセル（結合セル含む）の枠線が消失**
3. **白で塗りつぶしているセル（結合セル含む）が初期化され、グリッド線が見える**

---

## 診断結果サマリ

| 指標 | 元ファイル | エクスポート後 | 差分 |
|------|-----------|-------------|------|
| Fonts count | 8 | 8 | ±0 |
| Fills count | 6 | 7 | +1 |
| Borders count | 22 | 22 | ±0 |
| CellXfs count | 73 | 131 | +58 |
| セル数（s属性あり） | 83,853 | 60,040 | -23,813 |
| s値が変更されたセル | — | — | 1,095 |

---

## 根本原因一覧

### RC1: `_styleChanges` がパース時に常に `true` になる

- **重要度**: 最高（全問題の起点）
- **場所**:
  - `lib/src/sheet/sheet.dart:732` — `updateCell` で cellStyle が非 null の場合に `_excel._styleChanges = true` を設定
  - `lib/src/sheet/sheet.dart:1107-1108` — `_putData` で `cell._cellStyle != NumFormat.standard_0`（型が異なるため常に true）
- **メカニズム**:
  1. パーサーが `_parseCell` で `sheetObject.updateCell(cellStyle: _cellStyleList[s])` を呼ぶ
  2. `updateCell` 内で `_styleChanges = true` が設定される
  3. エクスポート時に `_processStylesFile()` が**常に**実行される
  4. styles.xml の fonts/fills/borders/cellXfs セクションが再構築される
- **影響**:
  - 変更なしのパススルーでも styles.xml が変更される
  - `_cellStyleReferenced` による元の `s` 属性値の透過パスが**使用されない**
  - `_processStylesFile()` で `_innerCellStyle`（全セルのユニーク CellStyle）が cellXfs に追加される（73 → 131 エントリ）

### RC2: `numberFormat.accepts()` が numFmtId=49 で `TextCellValue` を拒否する

- **重要度**: 最高（大量のスタイルインデックス不整合の直接原因）
- **場所**: `lib/src/number_format/num_format.dart:310`
- **問題コード**:
  ```dart
  // StandardNumericNumFormat.accepts()
  TextCellValue() => numFmtId == 0,  // numFmtId=49(@テキスト形式)でも false
  ```
- **メカニズム**:
  1. テンプレートの大半の cellXfs エントリが `numFmtId="49"`（テキスト形式 `@`）を使用
  2. セル値が `TextCellValue` の場合、`accepts()` が `false` を返す（`49 == 0` → `false`）
  3. `updateCell` が `cellStyle.copyWith(numberFormat: NumFormat.defaultFor(value))` を呼び、`NumFormat.standard_0` で**新しい CellStyle オブジェクト**を生成
  4. この新 CellStyle は元の `_cellStyleList` にないため `indexOf` が `-1` を返す
  5. `_innerCellStyle` 内のインデックス + 73 が使われ、**元とは異なる `s` 値**が出力される
- **診断証拠**: `B2: s=1 → s=76`（73 + _innerCellStyle 内インデックス 3）
- **影響**: numFmtId=49 を使う全セルのスタイルインデックスが狂い、フォント・塗りつぶし・枠線の組み合わせが変わる → **黒塗り・枠線消失の主要原因**

### RC3: CellStyle の重複排除で `indexOf`（値等価比較）を使用

- **重要度**: 高
- **場所**: `lib/src/utilities/utility.dart:60-62`
- **問題コード**:
  ```dart
  int _checkPosition(List<CellStyle> list, CellStyle cellStyle) {
    return list.indexOf(cellStyle);  // Equatable == で最初の一致を返す
  }
  ```
- **メカニズム**:
  1. 異なる cellXfs エントリが同一の Equatable props を持つ場合がある
  2. `indexOf` は**最初にマッチしたインデックス**を返すため、元のインデックスとは異なる値になる
  3. セルが参照する cellXfs エントリが変わり、fontId/fillId/borderId の組み合わせが変わる
- **検出された重複ペア**:

  | ペア | 原因 |
  |------|------|
  | xf[3] ↔ xf[7] | `horizontal` 属性の有無。xf[3] は指定なし（デフォルト Left）、xf[7] は `horizontal="left"`。CellStyle のデフォルト値が `HorizontalAlign.Left` のため同一視される |
  | xf[25] ↔ xf[41] | fontId=3, fillId=0, borderId=2, numFmtId=49, alignment=center,center |
  | xf[28] ↔ xf[43] | fontId=3, fillId=2, borderId=2, numFmtId=49, alignment=center,center |
  | xf[29] ↔ xf[58] | fontId=3, fillId=0, borderId=1, numFmtId=49, alignment=center,center |
  | xf[30] ↔ xf[60] | fontId=3, fillId=0, borderId=11, numFmtId=49, alignment=center,center |

- **診断証拠**: `D4: s=7 → s=3`、計 1,095 セルの s 値が変更
- **影響**: 一部のセルが異なるスタイル（異なる fill/border/font 参照）で描画される

### RC4: 結合セルの非左上セルが `_sheetData` から削除される

- **重要度**: 高（枠線消失・白塗りリセットの主要原因）
- **場所**: `lib/src/parser/parse.dart:220-237`（`_deleteAllButTopLeftCellsOfSpanObj`）
- **メカニズム**:
  1. `_parseMergedCells` が結合領域の左上セル以外を `sheet._removeCell()` で削除
  2. エクスポート時に `_setSheetElements` → `_setRows` が `_sheetData` から行・セルを再構築
  3. 削除されたセルは**XML に一切出力されない**（`<c>` 要素自体が存在しない）
  4. 元の Excel ファイルでは結合領域の全セルに `s` 属性が付与されている場合がある
- **診断証拠**:
  - 結合領域数: 7,898
  - 消失セル数: 23,813（s 属性付きセルが 83,853 → 60,040）
- **影響**:
  - 結合セルの**右端・下端の枠線が消失**（枠線情報を持つセルの XML 要素が消えるため）
  - 結合セルの**白塗りつぶし・背景色がリセット**（塗りつぶし情報を持つセルの XML 要素が消えるため）
  - グリッド線が見えるようになる

### RC5: `findAllElements` のスコープが `<dxf>` セクションを含む（潜在的リスク）

- **重要度**: 中（今回のテンプレートでは発生しないが、他のファイルで発生しうる）
- **場所**:
  - `lib/src/parser/parse.dart:262` — `document.findAllElements('patternFill')`
  - `lib/src/parser/parse.dart:282` — `document.findAllElements('border')`
- **メカニズム**:
  1. `findAllElements` はドキュメント全体を再帰的に検索する
  2. `<dxf>`（条件付き書式）セクション内の `<patternFill>` や `<border>` も `_patternFill` / `_borderSetList` に追加される
  3. fills/borders のインデックスが実際の `<fills>` / `<borders>` セクションと不一致になる
- **今回の影響**: テンプレートは `<dxfs count="0"/>` のため影響なし
- **潜在的影響**: 条件付き書式を含むファイルで fill/border インデックスが狂う

---

## 問題との対応関係

| 視覚的問題 | RC1 | RC2 | RC3 | RC4 | RC5 |
|-----------|-----|-----|-----|-----|-----|
| 黒塗りセル | ● 起点 | ●● 主因 | ● 副因 | | (潜在) |
| 枠線消失 | ● 起点 | ●● 主因 | ● 副因 | ●● 主因 | (潜在) |
| 白塗りリセット | ● 起点 | ● 副因 | ● 副因 | ●● 主因 | (潜在) |

### 因果関係フロー

```
RC1: _styleChanges = true (パース時)
  │
  ├─→ _processStylesFile() が呼ばれる
  │     │
  │     ├─→ RC2: accepts() バグで CellStyle が copyWith される
  │     │     └─→ _cellStyleList で見つからず、s 値が 73+ に変化
  │     │           └─→ 新 cellXfs エントリの fontId/fillId/borderId が異なる
  │     │                 └─→ 黒塗り・枠線消失
  │     │
  │     └─→ RC3: indexOf で最初の一致を返す
  │           └─→ s 値が別のインデックスに変化
  │                 └─→ fontId/fillId/borderId が異なる
  │                       └─→ 黒塗り・枠線消失
  │
  └─→ _cellStyleReferenced パス（元の s 値を透過）が使われない

RC4: 結合セルの非左上セル削除
  └─→ セルの XML 要素自体が消失
        └─→ 枠線・塗りつぶし情報の欠落
              └─→ 枠線消失・白塗りリセット
```

---

## 修正方針（案）

### 優先度 1: RC1 の修正 — パーサーで `_styleChanges` を設定しない

パーサー内の `updateCell` 呼び出しでは `_styleChanges` を設定しないようにする。これにより、パススルー時は `_cellStyleReferenced` パスが使われ、元の `s` 属性値がそのまま出力される。styles.xml も変更されない。

**方法案**: `updateCell` に内部フラグ（`_isParsing` 等）を追加し、パース中は `_styleChanges` を設定しないようにする。または `_parseCell` で `updateCell` を使わず直接 `_sheetData` に設定する。

### 優先度 2: RC2 の修正 — `accepts()` で numFmtId=49 を許可

```dart
// Before:
TextCellValue() => numFmtId == 0,
// After:
TextCellValue() => numFmtId == 0 || numFmtId == 49,
```

### 優先度 3: RC4 の修正 — 結合セルの非左上セルのスタイル情報を保持

非左上セルを `_sheetData` から削除する際、`_cellStyleReferenced` に元の `s` 属性を保持し、エクスポート時に空セル（値なし・スタイルあり）として出力する。

### 優先度 4: RC3 の修正 — `_checkPosition` でオブジェクト同一性を優先

```dart
int _checkPosition(List<CellStyle> list, CellStyle cellStyle) {
  // まずオブジェクト同一性で検索
  for (int i = 0; i < list.length; i++) {
    if (identical(list[i], cellStyle)) return i;
  }
  // フォールバック: 値等価比較
  return list.indexOf(cellStyle);
}
```

### 優先度 5: RC5 の修正 — `findAllElements` のスコープを限定

```dart
// Before:
document.findAllElements('patternFill').forEach(...)
// After:
document.findAllElements('fills').first.findAllElements('patternFill').forEach(...)
```

---

## 検証方法

上記修正後、以下を確認する:

1. `report-template.xlsm` のパススルーで styles.xml が元ファイルと同一であること
2. セルの `s` 属性が元ファイルと同一であること（83,853 セル全て）
3. 結合セルの非左上セルが XML に出力されること
4. Excel で開いた際に黒塗り・枠線消失・白塗りリセットが発生しないこと
5. 既存テストスイートが全件パスすること
