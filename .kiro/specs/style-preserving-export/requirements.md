# Requirements Document

## Introduction

本仕様は、XLSXファイルをインポートした際のセルスタイル情報（フォント、配置、罫線、塗りつぶし、数値書式等）を正確にパースし、エクスポート時に欠落なく再出力する「スタイル保持エクスポート」機能の要件を定義する。

現状、`_styleChanges` が `false` の場合は `xl/styles.xml` がそのまま透過されるため問題ないが、ユーザーがセルスタイルを変更した場合（`_styleChanges == true`）に新規追加される `<xf>` エントリや関連要素で多くのスタイル情報が欠落・誤変換される。また、インポート時のパーサーにも配置属性の読み取りノード誤り、下線スタイルの上書き等の既知バグが存在する。

本機能により、ラウンドトリップ（インポート → 編集 → エクスポート）で元ファイルのスタイル情報を最大限保持し、出力ファイルが MS Excel で正常に開ける状態を実現する。

## Requirements

### Requirement 1: アライメント情報の正確なインポート

**Objective:** As a ライブラリ利用者, I want インポート時にセルの配置（水平・垂直・回転・インデント）情報が正確にパースされること, so that 元ファイルの配置設定をプログラムから参照・保持できる

#### Acceptance Criteria

1. When XLSXファイルをインポートする, the Excel library shall `<xf>` の子要素 `<alignment>` から `horizontal` 属性を正確に読み取り `CellStyle.horizontalAlign` に反映する
2. When XLSXファイルをインポートする, the Excel library shall `<xf>` の子要素 `<alignment>` から `vertical` 属性を正確に読み取り `CellStyle.verticalAlign` に反映する
3. When XLSXファイルをインポートする, the Excel library shall `<xf>` の子要素 `<alignment>` から `textRotation` 属性を正確に読み取り `CellStyle.rotation` に反映する
4. When `<alignment>` 要素に `indent` 属性が存在する, the Excel library shall その値を `CellStyle` に保持する
5. When `<alignment>` 要素に `readingOrder` 属性が存在する, the Excel library shall その値を `CellStyle` に保持する

### Requirement 2: フォント情報の正確なインポート

**Objective:** As a ライブラリ利用者, I want インポート時にフォントの全プロパティが正確にパースされること, so that 元ファイルのフォント設定が正しく保持される

#### Acceptance Criteria

1. When `<u val="double">` を含むフォント定義をインポートする, the Excel library shall `Underline.Double` として正確にパースし、後続の処理で `Underline.Single` に上書きしない
2. When `<font>` に `<strike/>` 要素が存在する, the Excel library shall 取り消し線スタイルを `CellStyle` に保持する
3. When `<font>` に `<vertAlign val="superscript"/>` または `<vertAlign val="subscript"/>` が存在する, the Excel library shall その値を `CellStyle` に保持する
4. The Excel library shall `fontScheme` プロパティをフォントスタイルの等値比較（Equatable）の対象に含める

### Requirement 3: 塗りつぶし情報の正確なインポート/エクスポート

**Objective:** As a ライブラリ利用者, I want セルの塗りつぶし情報（パターン塗りつぶし・色）がラウンドトリップで保持されること, so that 元ファイルのセル背景が変化しない

#### Acceptance Criteria

1. When `_styleChanges` が `true` の状態でエクスポートする, the Excel library shall 既存セルのパターン塗りつぶしタイプ（"solid"、"gray125"、"darkGrid" 等すべての OOXML パターン）を正確に再出力する
2. When `<patternFill>` に `fgColor` と `bgColor` が個別に設定されている, the Excel library shall 両者を区別して保持し、エクスポート時に元の値を再出力する
3. If テーマカラー参照（`theme` + `tint` 属性）が塗りつぶしに使用されている, then the Excel library shall その参照情報を保持し、エクスポート時に再出力する

### Requirement 4: 罫線情報の完全なエクスポート

**Objective:** As a ライブラリ利用者, I want セルに設定された罫線スタイルがエクスポート時に正しく出力されること, so that 元ファイルの罫線表示がMS Excelで再現される

#### Acceptance Criteria

1. When セルに罫線スタイルが設定されている, the Excel library shall エクスポート時に `<xf>` 要素に `applyBorder="1"` 属性を出力する
2. When セルにフォントスタイルが設定されている, the Excel library shall エクスポート時に `<xf>` 要素に `applyFont="1"` 属性を正しく出力する
3. When セルに数値書式が設定されている, the Excel library shall エクスポート時に `<xf>` 要素に `applyNumberFormat="1"` 属性を出力する

### Requirement 5: 未対応XML属性の透過的保持

**Objective:** As a ライブラリ利用者, I want ライブラリが直接操作しないXML属性・要素がエクスポート時に消失しないこと, so that 元ファイルの未対応機能（セル保護、名前付きスタイル等）が維持される

#### Acceptance Criteria

1. When `_styleChanges` が `true` の状態でエクスポートする, the Excel library shall 既存の `<xf>` 要素に含まれる `protection` 子要素（`locked`, `hidden` 属性）を保持して再出力する
2. When `_styleChanges` が `true` の状態でエクスポートする, the Excel library shall 既存の `cellStyleXfs`、`cellStyles` セクションを変更せずに保持する
3. When `_styleChanges` が `true` の状態でエクスポートする, the Excel library shall 既存の `<xf>` の `xfId` 属性値を保持し、`0` にハードコードしない
4. When `_styleChanges` が `true` の状態でエクスポートする, the Excel library shall 既存の `<alignment>` 要素の全属性（`indent`, `readingOrder`, `justifyLastLine`, `relativeIndent` を含む）を保持して再出力する

### Requirement 6: テーマカラー参照の保持

**Objective:** As a ライブラリ利用者, I want テーマカラー参照（`theme` + `tint` 属性）を使用したスタイルがラウンドトリップで保持されること, so that テーマ依存の色指定が維持される

#### Acceptance Criteria

1. When フォントカラーが `theme` 属性で指定されている, the Excel library shall インポート時にテーマカラー参照情報を保持し、エクスポート時に `rgb` ではなく元の `theme`（+ `tint`）属性として再出力する
2. When 塗りつぶしカラーが `theme` 属性で指定されている, the Excel library shall インポート時にテーマカラー参照情報を保持し、エクスポート時に元の `theme`（+ `tint`）属性として再出力する
3. When 罫線カラーが `theme` 属性で指定されている, the Excel library shall インポート時にテーマカラー参照情報を保持し、エクスポート時に元の `theme`（+ `tint`）属性として再出力する

### Requirement 7: ラウンドトリップの整合性検証

**Objective:** As a ライブラリ利用者, I want インポート→（スタイル変更あり/なし）→エクスポートしたファイルが常にMS Excelで正常に開けること, so that 出力ファイルの互換性が保証される

#### Acceptance Criteria

1. When スタイル変更なしでエクスポートする, the Excel library shall `xl/styles.xml` を元ファイルと同一の内容で出力する（既存の透過動作を維持）
2. When スタイル変更ありでエクスポートする, the Excel library shall 有効な OOXML 構造を持つ `xl/styles.xml` を出力し、MS Excel でエラーなく開けるファイルを生成する
3. When スタイル変更ありでエクスポートする, the Excel library shall 変更されていないセルのスタイルインデックス（`s` 属性）を元の値のまま保持する
4. When 新規セルにスタイルを適用してエクスポートする, the Excel library shall 新規スタイルエントリを既存スタイルリストに正しく追加し、インデックスの整合性を維持する
5. If エクスポート時にスタイル情報の矛盾（重複フォント、未参照スタイル等）が検出された場合, then the Excel library shall 矛盾を自動修正し、有効なファイルを出力する

### Requirement 8: 既存APIとの後方互換性

**Objective:** As a ライブラリ利用者, I want 既存の `CellStyle` API を使ったコードが変更なく動作すること, so that 本変更によるリグレッションが発生しない

#### Acceptance Criteria

1. The Excel library shall 既存の `CellStyle` プロパティ（`fontFamily`, `fontSize`, `bold`, `italic`, `underline`, `fontColorHex`, `backgroundColorHex`, `horizontalAlign`, `verticalAlign`, `textWrapping`, `rotation`, `leftBorder`, `rightBorder`, `topBorder`, `bottomBorder`, `diagonalBorder`, `diagonalBorderUp`, `diagonalBorderDown`, `numberFormat`）の getter/setter 動作を変更しない
2. The Excel library shall 新規に追加するプロパティにはデフォルト値を設定し、既存コードが明示的に指定しなくても動作する
3. When 既存のテストスイートを実行する, the Excel library shall すべての既存テストがパスする
