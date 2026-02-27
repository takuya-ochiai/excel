# ギャップ分析: style-preserving-export

## 概要

インポートしたXLSXファイルのスタイル情報をエクスポート時に保持する機能について、現行コードベースと要件のギャップを分析した。主要な問題は3カテゴリに分類される:

1. **パーサーバグ**: アライメント属性の読み取りノード誤り、二重下線の上書き等
2. **エクスポーターバグ**: `applyFont`/`applyBorder` の未出力、パターン塗りつぶしの不完全な処理
3. **データモデルの不足**: テーマカラー、セル保護、取り消し線等のOOXMLプロパティが `CellStyle` に存在しない

---

## 1. 現状調査

### 1.1 主要ファイルと役割

| ファイル | 役割 | 行数 |
|---|---|---|
| `lib/src/parser/parse.dart` | `xl/styles.xml` のパース、`<xf>` → `CellStyle` 変換 | ~660 |
| `lib/src/save/save_file.dart` | `CellStyle` → XML再構築、`<xf>` 追加 | ~565 |
| `lib/src/sheet/cell_style.dart` | 公開APIデータモデル（20プロパティ） | ~371 |
| `lib/src/sheet/font_style.dart` | 内部フォントスタイル（重複排除用） | ~124 |
| `lib/src/sheet/border_style.dart` | `Border` + `_BorderSet` | ~94 |
| `lib/src/sheet/data_model.dart` | `Data` クラス（`cellStyle` setter で `_styleChanges` トリガー） | ~411 |
| `lib/src/excel.dart` | `Excel` クラス本体（スタイルリスト・変更フラグ管理） | ~337 |
| `lib/src/utilities/enum.dart` | `TextWrapping`, `VerticalAlign`, `HorizontalAlign`, `Underline`, `FontScheme` | ~40 |
| `lib/src/number_format/num_format.dart` | `NumFormat` sealed hierarchy + `NumFormatMaintainer` | ~300 |

### 1.2 アーキテクチャパターン

- **2フェーズ処理**: Import（ZIP→XML→In-Memory） / Export（In-Memory→XML→ZIP）
- **変更追跡ゲート**: `_styleChanges` が `false` なら `xl/styles.xml` は透過（完全保持）
- **追記型エクスポート**: `_styleChanges == true` 時、既存XML要素は保持し新規要素を末尾に追加
- **疎行列モデル**: `Map<int, Map<int, Data>>` でセルデータ管理
- **Equatable重複排除**: `CellStyle`, `_FontStyle`, `_BorderSet` はEquatableで値比較

### 1.3 既存テスト基盤

9つのラウンドトリップテストが存在（`test/excel_test.dart`, `test/export_test.dart`, `test/custom_excel_test.dart`）。スタイル関連テスト:

- 罫線ラウンドトリップ（`borders.xlsx`）
- 全罫線スタイル列挙（`borders2.xlsx`）
- 結合セル+罫線（`mergedBorders.xlsx`）
- リッチテキスト（`richText.xlsx`）
- カスタム数値書式ラウンドトリップ
- xlsx/xlsm透過ラウンドトリップ

**不足**: アライメント・フォント装飾・塗りつぶしパターン・テーマカラーのラウンドトリップテストは存在しない。

---

## 2. 要件-資産マッピング

### Requirement 1: アライメント情報の正確なインポート

| 受入基準 | 現行資産 | ギャップ |
|---|---|---|
| `horizontal` を `<alignment>` から読み取る | `parse.dart:442` — `node.getAttribute('horizontal')` | **バグ**: `node`（`<xf>`）から読んでいるため常に `null`。`child`（`<alignment>`）から読むべき |
| `vertical` を `<alignment>` から読み取る | `parse.dart:433` — 同上 | **バグ**: 同上 |
| `textRotation` を `<alignment>` から読み取る | `parse.dart:451` — 同上 | **バグ**: 同上 |
| `indent` を保持 | なし | **Missing**: `CellStyle` にフィールドなし、パーサーにコードなし |
| `readingOrder` を保持 | なし | **Missing**: 同上 |

### Requirement 2: フォント情報の正確なインポート

| 受入基準 | 現行資産 | ギャップ |
|---|---|---|
| 二重下線の正確なパース | `parse.dart:377-386` | **バグ**: `Underline.Double` を設定後、直後の単一下線チェックで `Underline.Single` に上書きされる |
| 取り消し線の保持 | なし | **Missing**: `CellStyle` にフィールドなし、パーサーに `<strike>` 読み取りなし |
| 上付き/下付き文字 | なし | **Missing**: `CellStyle` にフィールドなし |
| `fontScheme` の等値比較 | `font_style.dart:117-124` — `props` リスト | **バグ**: `_fontScheme` が `props` に含まれていない |

### Requirement 3: 塗りつぶし情報のラウンドトリップ

| 受入基準 | 現行資産 | ギャップ |
|---|---|---|
| 全パターンタイプの再出力 | `save_file.dart:306-332` | **Missing**: `solid`/`none`/`gray125`/`lightGray` 以外のパターンタイプは無視される |
| `fgColor`/`bgColor` の区別 | `save_file.dart:312-319` | **バグ**: エクスポート時に `fgColor` と `bgColor` を同一値で書き出す |
| テーマカラー参照の保持 | `parse.dart:258` | **Missing**: `rgb` 属性のみ読み取り、`theme`/`tint` は完全に無視 |

### Requirement 4: 罫線情報の完全なエクスポート

| 受入基準 | 現行資産 | ギャップ |
|---|---|---|
| `applyBorder` の出力 | なし | **Missing**: `<xf>` に `applyBorder="1"` を書くコードが存在しない |
| `applyFont` の正しい出力 | `save_file.dart:439-442` | **バグ**: `&&`（AND）条件で相互排他的リストを検査するため常に `false` |
| `applyNumberFormat` の出力 | なし | **Missing**: コードが存在しない |

### Requirement 5: 未対応XML属性の透過的保持

| 受入基準 | 現行資産 | ギャップ |
|---|---|---|
| `protection` の保持 | なし | **Missing**: パース・モデル・エクスポートすべてに存在しない |
| `cellStyleXfs`/`cellStyles` の保持 | 追記型アーキテクチャ | **Constraint**: 既存XMLノードは変更されないため、`_styleChanges == true` でも元のセクションは保持される（追加の対応不要） |
| `xfId` の保持 | `save_file.dart:428` | **バグ**: 新規 `<xf>` で `xfId` が常に `0` にハードコード |
| `<alignment>` の全属性保持 | なし | **Missing**: `indent`, `readingOrder`, `justifyLastLine`, `relativeIndent` はモデルにない |

### Requirement 6: テーマカラー参照の保持

| 受入基準 | 現行資産 | ギャップ |
|---|---|---|
| フォントのテーマカラー | `parse.dart:354-357` — `rgb` のみ | **Missing**: `theme`/`tint`/`indexed` は完全に未対応 |
| 塗りつぶしのテーマカラー | `parse.dart:258` — `rgb` のみ | **Missing**: 同上 |
| 罫線のテーマカラー | `parse.dart:296` — `rgb` のみ | **Missing**: 同上 |

### Requirement 7: ラウンドトリップの整合性検証

| 受入基準 | 現行資産 | ギャップ |
|---|---|---|
| 変更なしの透過保持 | `_styleChanges` ゲート機構 | **OK**: 既に実装済み |
| 変更ありで有効なOOXML出力 | 追記型エクスポーター | **Constraint**: 上記バグ修正により達成可能 |
| 未変更セルの `s` 属性保持 | `_cellStyleReferenced` | **Constraint**: `_styleChanges == true` 時のパスで一部の条件分岐に問題あり（`_createCell` line 69-77） |
| スタイルインデックスの整合性 | `_createCell` のオフセット計算 | **Constraint**: `_innerCellStyle` と `_cellStyleList` 間の重複排除なし |

### Requirement 8: 後方互換性

| 受入基準 | 現行資産 | ギャップ |
|---|---|---|
| 既存プロパティの動作維持 | 全getter/setter | **OK**: バグ修正はAPI非互換を発生させない |
| 新規プロパティのデフォルト値 | — | **要設計**: 新フィールド追加時のデフォルト値設計が必要 |
| 既存テストの通過 | 9つのラウンドトリップテスト | **OK**: バグ修正は正しい方向への変更 |

---

## 3. 実装アプローチの選択肢

### Option A: 既存コンポーネントの拡張（バグ修正 + 最小限の新フィールド追加）

**対象**: Requirement 1-4, 7, 8（パーサーバグ修正、エクスポーターバグ修正、必須フィールド追加）

#### 変更対象ファイル

| ファイル | 変更内容 |
|---|---|
| `parse.dart` | アライメント属性の読み取りノード修正（3行）、下線パース順序修正（5行程度） |
| `save_file.dart` | `applyFont` 条件修正（`&&` → `||`）、`applyBorder`/`applyNumberFormat` 追加 |
| `cell_style.dart` | `strikethrough`, `verticalText`（上付き/下付き）フィールド追加 |
| `font_style.dart` | `_fontScheme` を `props` に追加 |
| `enum.dart` | `HorizontalAlign` に `General`/`Fill`/`Justify`/`Distributed` 追加、`VerticalAlign` に `Justify`/`Distributed` 追加 |

**トレードオフ**:
- ✅ 最小変更量で最大効果（致命的バグ5件の修正）
- ✅ APIの後方互換性を維持しやすい
- ✅ 既存の追記型アーキテクチャを活用
- ❌ テーマカラー・セル保護・パターン塗りつぶし等は未対応のまま
- ❌ 段階的に追加すると設計の一貫性が低下するリスク

### Option B: 生XML保持アプローチ（新コンポーネント作成）

**対象**: Requirement 5-6を中心に全要件

既存の `<xf>` 要素の生XMLを `CellStyle` に紐付けて保持し、エクスポート時に `CellStyle` でモデル化されていない属性を元のXMLから復元する。

#### 新規コンポーネント

| コンポーネント | 役割 |
|---|---|
| `_RawXfData` クラス | パース時に `<xf>` の生XML要素を保持する内部データ構造 |
| `_ColorReference` クラス | `rgb`/`theme`+`tint`/`indexed` を統一的に扱うカラー参照モデル |

#### 変更対象ファイル

| ファイル | 変更内容 |
|---|---|
| `parse.dart` | 各 `<xf>` のXmlElementを `_RawXfData` に保存 |
| `save_file.dart` | 新規 `<xf>` 作成時に、ベースとなる `_RawXfData` から未対応属性をコピー |
| `cell_style.dart` | `_RawXfData?` フィールドを内部保持（公開APIには含めない） |
| `excel.dart` | `_rawXfDataList: List<_RawXfData>` を追加 |
| `colors.dart` | `_ColorReference` 型を導入し `theme`/`tint` を保持 |

**トレードオフ**:
- ✅ 未対応プロパティを自動的に保持（将来のOOXML要素追加にも対応）
- ✅ テーマカラー・セル保護・インデント等を一括で保持
- ✅ `CellStyle` に大量のフィールドを追加する必要がない
- ❌ 生XMLの管理が複雑化（`CellStyle` 変更時の整合性）
- ❌ 設計・実装の工数が大きい
- ❌ `_RawXfData` と `CellStyle` の二重管理による認知負荷

### Option C: ハイブリッドアプローチ（推奨）

**フェーズ1**: Option A のバグ修正 + 必須フィールド追加
**フェーズ2**: テーマカラー参照の `_ColorReference` モデル導入
**フェーズ3**: 既存 `<xf>` の未対応属性透過保持

#### フェーズ1（即時対応）

| 対象 | 変更内容 | Requirement |
|---|---|---|
| パーサーバグ3件 | アライメント読み取りノード修正、下線上書き修正 | Req 1, 2 |
| エクスポーターバグ3件 | `applyFont` 条件修正、`applyBorder`/`applyNumberFormat` 追加 | Req 4 |
| `_FontStyle.props` | `_fontScheme` 追加 | Req 2 |
| `CellStyle` 拡張 | `strikethrough` フィールド追加 | Req 2 |
| `HorizontalAlign`/`VerticalAlign` | 不足値追加 | Req 1 |
| テスト追加 | アライメント・フォント装飾のラウンドトリップテスト | Req 7 |

#### フェーズ2（カラーモデル改善）

| 対象 | 変更内容 | Requirement |
|---|---|---|
| `_ColorReference` 導入 | `rgb`/`theme`+`tint`/`indexed` の統一型 | Req 6 |
| パーサー拡張 | `<color>` 要素から全属性を読み取り | Req 6 |
| エクスポーター拡張 | `_ColorReference` → XML変換 | Req 6 |
| `CellStyle`/`Border` | カラー型を `ExcelColor` から `_ColorReference` へ内部拡張 | Req 3, 6 |

#### フェーズ3（未対応属性の透過保持）

| 対象 | 変更内容 | Requirement |
|---|---|---|
| `_RawXfData` 導入 | 既存 `<xf>` の未解釈属性を保持 | Req 5 |
| エクスポーター修正 | 新規 `<xf>` に `_RawXfData` から `protection`/`indent` 等をコピー | Req 5 |
| `xfId` 修正 | パース時に `xfId` を保持し、エクスポート時に復元 | Req 5 |

**トレードオフ**:
- ✅ 段階的デリバリーで各フェーズでテスト可能
- ✅ フェーズ1だけで主要なバグが全修正される
- ✅ 各フェーズの変更量が管理可能
- ❌ フェーズ間でのAPIの微調整が必要になる可能性
- ❌ フェーズ2のカラーモデル変更は `ExcelColor` APIへの影響を慎重に設計する必要あり

---

## 4. 複雑性・リスク評価

### 工数見積

| アプローチ | 工数 | 根拠 |
|---|---|---|
| Option A（バグ修正のみ） | **S**（1-3日） | 既存パターン内の修正、影響範囲が限定的 |
| Option B（生XML保持） | **L**（1-2週） | 新規データモデル+パーサー/エクスポーター大規模変更 |
| Option C フェーズ1 | **S**（1-3日） | バグ修正+少数のフィールド追加 |
| Option C フェーズ2 | **M**（3-7日） | カラーモデルの再設計+後方互換性の維持 |
| Option C フェーズ3 | **M**（3-7日） | 生XML保持の設計+整合性管理 |
| Option C 全体 | **L**（1-2週） | 3フェーズの合計 |

### リスク評価

| リスク要因 | レベル | 根拠 |
|---|---|---|
| パーサーバグ修正 | **Low** | 原因が明確、修正が限定的 |
| `applyFont`/`applyBorder` 修正 | **Low** | ロジック条件の修正のみ |
| `CellStyle` フィールド追加 | **Low** | デフォルト値付きの追加で後方互換 |
| テーマカラーモデル導入 | **Medium** | `ExcelColor` APIへの影響、`theme1.xml` パースの追加検討 |
| 生XML属性の透過保持 | **Medium** | `CellStyle` 変更時の整合性管理が複雑 |
| パターン塗りつぶしの全対応 | **Medium** | OOXMLパターンタイプ全18種の正確な再出力 |

---

## 5. 設計フェーズへの推奨事項

### 推奨アプローチ: **Option C（ハイブリッド）**

フェーズ1のバグ修正だけでも大幅な品質改善が得られ、フェーズ2-3は後続イテレーションとして段階的に実装可能。

### 設計フェーズで決定すべき事項

1. **カラーモデル設計**: `ExcelColor` の公開APIを維持しつつ `theme`/`tint` を内部保持する方法（`ExcelColor` を拡張するか、新型 `_ColorReference` を内部に持つか）
2. **`_RawXfData` のスコープ**: どの `<xf>` 属性を生XMLとして保持し、どれを `CellStyle` フィールドに昇格させるか
3. **`CellStyle` in-place変更の検出**: 現在 `data.cellStyle = ...` のsetterでのみ `_styleChanges` がトリガーされるが、`cellStyle.isBold = true` のようなin-place変更は検出されない。この問題の対処方針
4. **パターン塗りつぶしモデル**: `_patternFill` の `List<String>` 表現を構造化された型に変更するか

### Research Needed（設計フェーズで調査）

- テーマカラー解決（`xl/theme/theme1.xml` のパースと `tint` 計算）の必要性と実装コスト
- `cellStyleXfs` / `cellStyles`（名前付きスタイル）の保持が実用上必要かどうか
- OOXML Strict vs Transitional の差異がスタイル処理に影響するか
- `<dxf>`（条件付き書式の差分スタイル）の保持が要件スコープ内か

---

## 6. 発見された致命的バグ一覧

| # | 箇所 | 説明 | 影響 |
|---|---|---|---|
| 1 | `parse.dart:433,442,451` | `<alignment>` の `vertical`/`horizontal`/`textRotation` を `<xf>` から読み取っている | 全XLSXファイルの配置情報が消失 |
| 2 | `parse.dart:377-386` | 二重下線パース後に単一下線チェックで上書き | 二重下線が常に単一下線になる |
| 3 | `save_file.dart:439-442` | `applyFont` の条件が `&&` で相互排他的リスト | `applyFont="1"` が書かれない |
| 4 | `save_file.dart` | `applyBorder` 出力コードが存在しない | 罫線が有効化されない |
| 5 | `font_style.dart:117-124` | `_fontScheme` が `props` に含まれていない | 異なる `fontScheme` のフォントが同一視される |
| 6 | `save_file.dart:428` | `xfId` が常に `0` にハードコード | 名前付きスタイル参照が消失 |
