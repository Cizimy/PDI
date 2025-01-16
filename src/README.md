# ソースコード

このディレクトリには本番環境で使用する実装コードを格納します。

## ディレクトリ構造

```
src/
├── data/          # データ層モジュール
│   ├── modDBHandler.bas       # データベース操作
│   ├── modFileSystem.bas      # ファイル操作
│   └── modDataValidator.bas   # データ検証
│
├── ui/            # UI層モジュール
│   ├── modUIController.bas    # UI制御
│   ├── modInputValidator.bas  # 入力検証
│   └── modFormFactory.bas     # フォーム生成
│
├── business/      # ビジネスロジック層
│   ├── modBusinessRules.bas   # ビジネスルール
│   ├── modWorkflowManager.bas # ワークフロー管理
│   └── modCalculator.bas      # 計算処理
│
├── utils/         # ユーティリティ層
│   ├── modStringHelper.bas    # 文字列操作
│   ├── modDateHelper.bas      # 日付処理
│   └── modArrayHelper.bas     # 配列操作
│
├── security/      # セキュリティ層
│   ├── modSecurity.bas        # セキュリティ機能
│   ├── modEncryption.bas      # 暗号化処理
│   └── modAccessControl.bas   # アクセス制御
│
└── system/        # システム基盤層
    ├── modErrorHandler.bas    # エラー処理
    ├── modLogger.bas          # ログ管理
    └── modConfigManager.bas   # 設定管理
```

## モジュール命名規則

- 標準モジュール: `mod〈機能名〉`
- クラスモジュール: `cls〈クラス名〉`
- フォーム: `frm〈画面名〉`

## 実装ガイドライン

1. レイヤー間の依存関係
   - UI層 → ビジネスロジック層 → データ層
   - すべての層 → ユーティリティ層
   - すべての層 → システム基盤層

2. モジュール設計原則
   - 単一責任の原則に従う
   - インターフェースを明確に定義
   - 実装の詳細はプライベート関数で隠蔽
   - 各モジュールは独立してテスト可能に

3. エラー処理とログ
   - すべてのパブリック関数でエラー処理を実装
   - エラーは`modErrorHandler`を通じて一元管理
   - 重要な操作は`modLogger`でログを記録

4. セキュリティ考慮事項
   - 機密データは`modEncryption`で暗号化
   - ユーザー権限は`modAccessControl`で管理
   - 入力値は必ず検証してから処理

5. 保守性と拡張性
   - コードは十分なコメントを付加
   - 設定値は`modConfigManager`で一元管理
   - 新機能追加時は該当する層に配置
