# PDI - Professional Development Initiative

このリポジトリはVBA（Visual Basic for Applications）開発のためのリソースとコードを管理します。

## 構成

```
PDI/
├── config/          # 設定ファイル
├── docs/            # ドキュメント
├── examples/        # サンプルコードとテンプレート
├── src/            # 本番用ソースコード
├── templates/      # プロジェクトテンプレート
├── tests/          # テストコード
└── tools/          # 開発支援ツール
```

## ディレクトリ構成の説明

- `config/`: プロジェクトの設定ファイルを格納
- `docs/`: プロジェクトのドキュメント（コーディング規約など）
- `examples/`: サンプルコードとその説明
- `src/`: 本番環境で使用する実装コード
- `templates/`: 新規プロジェクト用のテンプレート
- `tests/`: 単体テストやインテグレーションテスト
- `tools/`: 開発効率を向上させるユーティリティツール

## 開発環境

- Microsoft Office 2016以降
- VBA Editor (Alt + F11)
- Git

## 使用方法

1. このリポジトリをクローン
2. `/examples` ディレクトリ内のサンプルコードを参照
3. 必要に応じてコードをカスタマイズして使用

## コーディング規約

- プロシージャ名は動詞で始める（例：`GetData`, `ProcessReport`）
- 変数名はキャメルケースを使用
- モジュール名は機能を表す名詞を使用
- すべての関数とサブプロシージャにコメントを付ける

## ライセンス

MIT License
