# Project Structure

## Organization Philosophy

**Monolithic Library + 機能別ディレクトリ**構成。`lib/excel.dart` を単一のライブラリエントリポイントとし、`part`/`part of` ディレクティブで全ソースを結合する。`lib/src/` 以下は機能ドメインごとにサブディレクトリへ分類。

## Directory Patterns

### Library Entry Point
**Location**: `lib/excel.dart`
**Purpose**: ライブラリの公開 API 定義。外部パッケージの import と全 part ファイルの宣言
**Pattern**: `library excel;` + `part 'src/...';` の列挙

### Core Logic
**Location**: `lib/src/excel.dart`
**Purpose**: `Excel` クラス本体。ファクトリコンストラクタ（`decodeBytes`, `createExcel`）、シートアクセス、保存処理
**Pattern**: ライブラリのメインファサード

### Sheet Model
**Location**: `lib/src/sheet/`
**Purpose**: シート・セル・スタイルに関するデータモデルと操作ロジック
**Example**: `sheet.dart`（Sheet クラス）、`data_model.dart`（Data, CellValue sealed class）、`cell_style.dart`

### Parser
**Location**: `lib/src/parser/`
**Purpose**: XLSX ファイル（ZIP + XML）のデコードとパース
**Example**: `parse.dart` — Archive を展開し XML を解析して In-Memory Model を構築

### Save
**Location**: `lib/src/save/`
**Purpose**: In-Memory Model から XLSX ファイルへのシリアライズ
**Example**: `save_file.dart`（XML 生成 + ZIP 圧縮）、`self_correct_span.dart`（結合セル補正）

### Shared Strings
**Location**: `lib/src/sharedStrings/`
**Purpose**: OOXML 共有文字列テーブルの管理、リッチテキスト（`TextSpan`）パース
**Pattern**: `_SharedStringsMaintainer` がインデックス管理。パース時は `addFromParsedXml()` で位置保持、通常追加は重複排除

### Number Format
**Location**: `lib/src/number_format/`
**Purpose**: OOXML 数値フォーマットの sealed class 階層と ID レジストリ管理
**Pattern**: `NumFormat` sealed class + `NumFormatMaintainer` がカスタム ID（164〜）を採番

### Utilities
**Location**: `lib/src/utilities/`
**Purpose**: 共通ヘルパー、定数、列挙型
**Example**: `constants.dart`（OOXML 定数）、`enum.dart`（TextWrapping, VerticalAlign, FontVerticalAlign 等）、`span.dart`

### Platform Abstraction
**Location**: `lib/src/web_helper/`
**Purpose**: Conditional import でブラウザ/ネイティブの保存処理を切り替え
**Pattern**: `client_save_excel.dart`（デフォルト）、`web_save_excel_browser.dart`（Web 向け）

### Tests
**Location**: `test/`
**Purpose**: 機能別テストファイル + サンプルデータ
**Pattern**: 機能ドメインごとにテストファイルを分割（`excel_test.dart`, `parser_test.dart`, `exporter_test.dart`, `style_index_test.dart`, `style_model_test.dart` 等）。`test_resources/` に XLSX/XLSM サンプルファイル

### Examples
**Location**: `example/`
**Purpose**: 機能別のデモスクリプト
**Pattern**: `excel_*.dart` — 各ファイルが特定機能（スタイル、ボーダー、サイズ等）をデモ

## Naming Conventions

- **Files**: `snake_case.dart`（例: `cell_style.dart`, `num_format.dart`）
- **Classes**: `PascalCase`（例: `Excel`, `Sheet`, `Data`, `CellStyle`）
- **Private members**: `_` プレフィックス（例: `_sheetData`, `_Span`）
- **Methods/Properties**: `camelCase`（例: `updateCell()`, `maxRows`）
- **Enums**: `PascalCase` 名 + `camelCase` 値（例: `TextWrapping.WrapText`）

## Import Organization

```dart
// lib/excel.dart (ライブラリルート)
library excel;

import 'dart:convert';
import 'dart:typed_data';
// ... stdlib imports

import 'package:archive/archive.dart';
import 'package:xml/xml.dart';
// ... package imports

// Conditional import (Web 対応)
import 'src/web_helper/client_save_excel.dart'
    if (dart.library.js_interop) 'src/web_helper/web_save_excel_browser.dart';

// Part declarations
part 'src/excel.dart';
part 'src/sheet/sheet.dart';
// ... all part files
```

**Pattern**: stdlib → packages → conditional imports → part declarations

## Code Organization Principles

- **単一ライブラリ**: 全 `.dart` ファイルは `part of excel;` で同一名前空間を共有
- **疎行列モデル**: `Map<int, Map<int, Data>>` でシートデータを表現（空セルはメモリ不使用）
- **Value Object**: `Equatable` を活用し、CellValue・CellStyle・Data・ColorValue・FillValue・CellProtection を値オブジェクトとして扱う
- **変更追跡**: `_mergeChanges`、`_styleChanges` 等のフラグで差分保存を制御
- **XML パススルー**: 未対応の XML セクションは生 `XmlElement` として保持し Export 時に再挿入

---
_Document patterns, not file trees. New files following patterns shouldn't require updates_
