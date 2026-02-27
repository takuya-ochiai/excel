# Research & Design Decisions

## Summary
- **Feature**: style-preserving-export
- **Discovery Scope**: Extension（既存システムの改善）
- **Key Findings**:
  - パーサーにアライメント属性の読み取り元ノード誤り（`node` vs `child`）のバグが存在
  - 下線パース処理で `Underline.Double` が `Underline.Single` に上書きされるバグ
  - エクスポーターの `applyFont` 判定ロジックに論理バグあり（両リストに存在を要求）
  - `applyBorder`、`applyNumberFormat` 属性がエクスポート時に未出力
  - テーマカラー参照、パターン塗りつぶし詳細、セル保護属性がモデルに未定義
  - `_FontStyle` の Equatable に `_fontScheme` が含まれていない

## Research Log

### アライメント属性パースのバグ調査
- **Context**: requirements.md 要件1.1-1.3 の調査中に発見
- **Sources Consulted**: `lib/src/parser/parse.dart` Lines 425-456
- **Findings**:
  - `_parseStyles()` 内の `<cellXfs>` パースにおいて、`<alignment>` 子要素の属性（`horizontal`, `vertical`, `textRotation`）を読み取る際、`child`（alignment要素）ではなく `node`（xf要素）から `getAttribute()` を呼んでいる
  - OOXML の構造: `<xf><alignment horizontal="center" vertical="top"/></xf>`
  - `node` は `<xf>` 要素であり `horizontal` 等の属性を持たないため、常に `null` が返る
  - 結果: 全セルのアライメントがデフォルト値（Left / Bottom / 0°）になる
- **Implications**: パーサーの3箇所で `node.getAttribute()` → `child.getAttribute()` に修正が必要

### 下線パース順序のバグ調査
- **Context**: requirements.md 要件2.1 の調査中に発見
- **Sources Consulted**: `lib/src/parser/parse.dart` Lines 377-387
- **Findings**:
  - 現行コード:
    1. `_nodeChildren(font, 'u', attribute: 'val')` が非null → `Underline.Double` を設定
    2. `_nodeChildren(font, 'u')` が非null → `Underline.Single` を設定（上書き）
  - `<u val="double"/>` の場合、両方の条件が `true` になり、最終的に `Single` に上書きされる
  - OOXML 仕様: `<u/>` = Single、`<u val="double"/>` = Double、`<u val="singleAccounting"/>` 等
- **Implications**: パース順序の逆転、または `val` 属性の値に基づく分岐ロジックへ修正が必要

### applyFont ロジックバグの調査
- **Context**: requirements.md 要件4.2 の調査中に発見
- **Sources Consulted**: `lib/src/save/save_file.dart` Lines 439-441
- **Findings**:
  - 現行コード: `_fontStyleIndex(_excel._fontStyleList, _fs) != -1 && _fontStyleIndex(innerFontStyle, _fs) != -1`
  - 元のフォントリスト（`_fontStyleList`）と新規収集リスト（`innerFontStyle`）の**両方**にフォントが存在することを要求
  - 新規追加フォントは `_fontStyleList` に存在せず、既存フォントは `innerFontStyle` に存在しない場合がある
  - 結果: `applyFont="1"` がほぼ出力されない
- **Implications**: `fontId > 0` などの単純な条件に修正が必要

### _FontStyle Equatable の不備
- **Context**: requirements.md 要件2.4 の調査中に発見
- **Sources Consulted**: `lib/src/sheet/font_style.dart` Lines 117-124
- **Findings**:
  - `_FontStyle` の `props` リストに `_fontScheme` が含まれていない
  - `fontScheme` が異なるフォントが同一と判定され、スタイル重複排除時にエントリが失われる可能性
- **Implications**: `props` への `_fontScheme` 追加が必要

### OOXML パターン塗りつぶしタイプ
- **Context**: requirements.md 要件3.1 の調査
- **Sources Consulted**: ECMA-376 ST_PatternType 定義
- **Findings**:
  - OOXML 定義の全パターンタイプ（18種）: none, solid, mediumGray, darkGray, lightGray, darkHorizontal, darkVertical, darkDown, darkUp, darkGrid, darkTrellis, lightHorizontal, lightVertical, lightDown, lightUp, lightGrid, lightTrellis, gray125, gray0625
  - 現行エクスポートは "none", "gray125", "lightGray", "solid"（hex色）のみ対応
  - `fgColor` と `bgColor` は別の色を持つことが可能（パターン前景色 vs パターン背景色）
- **Implications**: FillValue モデルの導入により全パターンタイプ・個別色に対応

### テーマカラー参照の構造
- **Context**: requirements.md 要件6.1-6.3 の調査
- **Sources Consulted**: OOXML `<color>` 要素仕様
- **Findings**:
  - `<color>` 要素は以下の属性を持つ:
    - `rgb="AARRGGBB"` — 直接指定のARGB値
    - `theme="0-11"` — テーマカラーインデックス
    - `tint="-1.0 ~ 1.0"` — テーマカラーの明度調整
    - `indexed="0-63"` — レガシーインデックスカラー
    - `auto="true|false"` — 自動色（システム色）
  - テーマカラーは `theme1.xml` で定義される基本色を参照
  - RGB変換するとテーマ切り替え時に追従しなくなるため、参照を保持する必要がある
- **Implications**: ColorValue 値オブジェクトで RGB/テーマ/インデックスの全表現をカバー

### cellStyleXfs / cellStyles セクション
- **Context**: requirements.md 要件5.2 の調査
- **Sources Consulted**: `lib/src/save/save_file.dart` `_processStylesFile()`
- **Findings**:
  - `cellStyleXfs`: 名前付きスタイル（Normal, Heading 1 等）のベースフォーマット定義
  - `cellStyles`: 名前付きスタイルのメタデータ（名前、builtinId等）
  - `cellXfs` の各 `<xf>` の `xfId` は `cellStyleXfs` 内のインデックスを参照
  - 現行の `_processStylesFile()` はこれらセクションを一切考慮していない
  - `_styleChanges == true` 時に完全再生成するため、これらのセクションが消失する
- **Implications**: Import 時にこれらのセクションを Raw XML として保存し、Export 時に再挿入

### 既存エクスポートにおけるスタイルインデックス管理
- **Context**: requirements.md 要件7.3-7.4 の調査
- **Sources Consulted**: `lib/src/save/save_file.dart` Lines 625-651 (`_createCell`)
- **Findings**:
  - 現行フロー:
    1. `_processStylesFile()` で全セルの CellStyle を収集 → `_innerCellStyle`
    2. `_createCell()` でセルの `s` 属性を決定: `_cellStyleList` 内 → そのインデックス、`_innerCellStyle` 内 → `_cellStyleList.length + そのインデックス`
  - 問題: 元スタイルリストの内容が再生成されたスタイルと一致しない場合、インデックスがずれる
  - 特に、元のスタイルに含まれていた属性（protection、xfId等）がモデルに無いため、等値比較で不一致となる
- **Implications**: CellStyle モデルを完全にし、元スタイルとの正確な等値比較を可能にする

## Architecture Pattern Evaluation

| Option | 説明 | 強み | リスク/制限 | 備考 |
|--------|------|------|-------------|------|
| A: Raw XML 保存 | 元の XML ノードをそのまま保存し、変更時のみ再生成 | 最大限の忠実度 | 二重データモデル、複雑性増加 | 未知の属性も保持可能 |
| B: 完全モデル拡張 | CellStyle を全 OOXML 属性に対応拡張 | 単一データソース、型安全 | 全属性のモデル化が必要 | エッジケースのリスク |
| C: ハイブリッド | モデル拡張 + 未対応セクションの Raw XML 透過 | バランスの良いアプローチ | 2つの保存戦略の管理 | **選択** |

**選択**: Option C — ハイブリッドアプローチ
- 主要な属性は CellStyle モデルに追加（型安全・操作可能）
- `cellStyleXfs`/`cellStyles` は Raw XML として透過保持（操作不要のため）
- 未知の `<xf>` 属性は拡張モデルでカバー

## Design Decisions

### Decision: テーマカラーの表現方法
- **Context**: フォント、塗りつぶし、罫線の各色でテーマカラー参照を保持する必要がある（要件6.1-6.3）
- **Alternatives Considered**:
  1. 既存の hex String にテーマ情報をエンコード（例: "THEME:1:0.5"）
  2. ColorValue 値オブジェクトで RGB/テーマ/インデックスを統一表現
  3. 各プロパティにテーマ用フィールドを個別追加（themeIndex, tint）
- **Selected Approach**: Option 2 — ColorValue 値オブジェクト
- **Rationale**: OOXML の `<color>` 要素と1:1対応し、表現が統一的。既存の hex getter/setter との後方互換も維持可能
- **Trade-offs**: 新規クラス追加が必要だが、コードの意図が明確になる
- **Follow-up**: 既存テストで hex ベースの色指定が引き続き動作することを確認

### Decision: パターン塗りつぶしの表現方法
- **Context**: 現行の `_patternFill` は `List<String>` で色または "none"/"gray125" のみ。全パターンタイプと fgColor/bgColor の個別保持が必要（要件3.1-3.3）
- **Alternatives Considered**:
  1. `_patternFill` を `List<Map<String, dynamic>>` に拡張
  2. FillValue 値オブジェクトの導入
- **Selected Approach**: Option 2 — FillValue 値オブジェクト
- **Rationale**: 型安全、Equatable による正確な重複排除、ColorValue との統合
- **Trade-offs**: 既存の `backgroundColorHex` API との互換レイヤーが必要

### Decision: cellStyleXfs / cellStyles の保持方法
- **Context**: ライブラリが直接操作しないセクションの透過保持（要件5.2）
- **Alternatives Considered**:
  1. セクションを完全にモデル化
  2. Raw XML ノードとして保存・再挿入
- **Selected Approach**: Option 2 — Raw XML 保存
- **Rationale**: これらのセクションはライブラリの操作対象外であり、モデル化のコストに見合わない。Raw XML をそのまま保持することで変更リスクを最小化
- **Trade-offs**: これらのセクションの編集機能は提供されない（Non-Goals に含む）

### Decision: スタイルインデックス管理戦略
- **Context**: `_styleChanges == true` 時に元スタイルのインデックスを保持する必要がある（要件7.3）
- **Alternatives Considered**:
  1. 元スタイルリストを起点に差分のみ追加
  2. 全スタイルを再収集して再番号付け
- **Selected Approach**: Option 1 — 元スタイルリスト保持 + 差分追加
- **Rationale**: 未変更セルの `s` 属性を変更せずに済み、ラウンドトリップの安定性が向上。新規/変更スタイルのみ末尾に追加
- **Trade-offs**: 使われなくなった元スタイルエントリが残る可能性があるが、OOXML 仕様上問題なし

## Risks & Mitigations
- **R1: 後方互換性の破壊** — 全既存プロパティの getter/setter を互換レイヤーで維持。既存テスト全パスを確認条件とする
- **R2: テーマカラー解決の複雑さ** — テーマ→RGB 変換は行わず、参照情報の保持のみに集中。RGB 変換は将来の拡張に委ねる
- **R3: エッジケースの OOXML 属性** — 主要属性はモデル化、未知の属性は要件範囲外として Non-Goals に明記
- **R4: パフォーマンス影響** — FillValue/ColorValue の導入による Equatable 比較コスト増。実運用サイズのファイルでは無視できる程度と判断

## References
- ECMA-376 Part 1 (Office Open XML) — styles.xml 要素定義
- `lib/src/parser/parse.dart` — 現行パーサー実装
- `lib/src/save/save_file.dart` — 現行エクスポーター実装
- `lib/src/sheet/cell_style.dart` — 現行 CellStyle モデル
- `lib/src/sheet/font_style.dart` — 現行 _FontStyle モデル
