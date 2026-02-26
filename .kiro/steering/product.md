# Product Overview

Pure Dart で実装された XLSX (Microsoft Excel) ファイルの読み書きライブラリ。Flutter (iOS/Android/Web) およびサーバーサイド Dart の両方で動作し、プラットフォーム固有の外部依存なしに Excel ファイルを操作できる。

## Core Capabilities

- **XLSX の読み込み・書き込み**: 既存ファイルのデコードと新規ファイルの作成
- **型安全なセル操作**: Sealed class `CellValue` による 8 つの型（Text, Int, Double, Bool, Date, Time, DateTime, Formula）の網羅的パターンマッチング
- **セルスタイリング**: フォント、色、配置、ボーダー、折り返し、数値フォーマット
- **シート管理**: 作成、リネーム、コピー、削除、結合セル、列幅・行高の制御
- **クロスプラットフォーム**: Web（ブラウザ保存）とネイティブ（ファイル I/O）の両対応

## Target Use Cases

- ビジネスアプリケーションでのレポート・帳票の動的生成
- データエクスポート（テーブルデータを Excel 形式で出力）
- 既存 Excel ファイルの読み込み・変換・再出力パイプライン
- MS Excel / Google Sheets / LibreOffice で作成されたファイルの相互運用

## Value Proposition

- **Pure Dart**: ネイティブ拡張やプラットフォーム固有 API に依存しない
- **型安全**: `dynamic` を排除し、Dart 3 の sealed class で網羅的な型チェックを実現（v4.0.0〜）
- **OOXML 準拠**: ZIP + XML ベースの XLSX フォーマットを直接操作
- **軽量依存**: archive, xml, equatable, collection, web の 5 パッケージのみ

---
_Focus on patterns and purpose, not exhaustive feature lists_
