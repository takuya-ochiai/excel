# Technology Stack

## Architecture

OOXML (Office Open XML) ベースの 2 フェーズ処理アーキテクチャ:

1. **Import**: ZIP Archive → XML Documents → In-Memory Objects
2. **Export**: Modified Objects → XML Documents → ZIP Archive

変更追跡フラグ（`_mergeChanges`, `_styleChanges`, `_rtlChanges`）により、変更された部分のみを再シリアライズする最適化を行う。

## Core Technologies

- **Language**: Dart 3.6.0+ (< 4.0.0)
- **Platform**: Flutter (iOS/Android/Web) + Dart VM (サーバーサイド)
- **Package Manager**: Dart pub

## Key Libraries

| Package | Purpose |
|---------|---------|
| `archive` ^4.0.4 | ZIP エンコード/デコード（XLSX は ZIP 形式） |
| `xml` >=5.0.0 <7.0.0 | XML パース/生成（OOXML ドキュメント構造） |
| `equatable` ^2.0.0 | 値オブジェクトの等値比較（CellValue, CellStyle, Data） |
| `collection` ^1.15.0 | ユーティリティコレクション |
| `web` ^1.1.1 | Web API（ブラウザ保存、dart:html の後継） |

## Development Standards

### Type Safety
- Dart 3 sealed class による網羅的パターンマッチング（`CellValue`）
- `dynamic` は排除、明示的な型を使用
- `Equatable` による値ベースの等値比較

### Code Quality
- `lints: ^5.1.1`（Dart 標準 lint ルール）
- `dart analyze` によるスタティック解析

### Testing
- `package:test ^1.23.0`
- 70+ テストケース、22 個のサンプル XLSX ファイルによる実データ検証
- MS Excel / Google Sheets / LibreOffice のクロスアプリケーション互換性テスト

## Development Environment

### Required Tools
- Dart SDK 3.6.0+

### Common Commands
```bash
# 依存関係インストール
dart pub get

# スタティック解析
dart analyze

# テスト実行
dart test

# サンプル実行
dart run example/excel_example.dart
```

## Key Technical Decisions

- **Parts-based monolithic library**: `part`/`part of` で単一ライブラリとして構成し、pub visibility の複雑さを回避
- **Sealed CellValue (v4.0.0)**: `dynamic` からの脱却、コンパイル時の型安全性を確保
- **web パッケージ移行 (v5.0.0)**: `dart:html` から `package:web` へ移行、Wasm 対応
- **Conditional imports**: `web_helper/` でブラウザ/ネイティブのプラットフォーム差異を吸収
- **In-Memory Model**: `Map<int, Map<int, Data>>` による疎行列表現でシートデータを管理

## CI/CD

- **GitHub Actions**: PR/Push → `dart analyze` + `dart test`
- **Publishing**: `publish` ブランチへの Push で pub.dev へ自動公開

---
_Document standards and patterns, not every dependency_
