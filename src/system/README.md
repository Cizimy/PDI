# システム基盤層

このディレクトリには、アプリケーション全体の基盤となるシステムレベルのモジュールが含まれています。

## モジュール構成

### modErrorHandler.bas
エラー処理の中央管理システム。アプリケーション全体のエラーハンドリングを担当します。

```vb
' エラーハンドリングの使用例
Public Sub SomeFunction()
    On Error GoTo ErrorHandler
    ' 処理内容
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "SomeModule.SomeFunction", _
                               etDatabase, _
                               "追加情報"
End Sub
```

### modLogger.bas
ログ管理システム。アプリケーションの動作ログを記録します。

```vb
' ログ出力の使用例
Public Sub ProcessData()
    ' モジュールの初期化
    modLogger.Initialize
    
    ' 各種ログレベルでの出力
    modLogger.Debug "デバッグ情報", "ProcessData"
    modLogger.Info "処理開始", "ProcessData"
    modLogger.Warning "警告メッセージ", "ProcessData"
    modLogger.Error "エラー発生", "ProcessData"
    modLogger.Critical "重大なエラー", "ProcessData"
End Sub
```

### modConfigManager.bas
設定管理システム。アプリケーションの設定を一元管理します。

```vb
' 設定管理の使用例
Public Sub InitializeApp()
    ' モジュールの初期化
    modConfigManager.Initialize "config/custom_settings.json"
    
    ' 設定値の取得
    Dim logLevel As String
    logLevel = modConfigManager.GetValue("logging", "level", "INFO")
    
    ' 設定値の更新
    modConfigManager.SetValue "logging", "level", "DEBUG"
End Sub
```

## モジュール間の連携

1. エラー処理とログ記録
```vb
Private Sub HandleDatabaseError(ByRef errorInfo As ErrorInfo)
    ' エラーをログに記録
    modLogger.Error errorInfo.Description, _
                    "Database Error: " & errorInfo.Source
    
    ' エラーメッセージを表示
    MsgBox "データベースエラーが発生しました。" & vbNewLine & _
           "詳細はログを確認してください。", _
           vbCritical + vbOKOnly, _
           "データベースエラー"
End Sub
```

2. 設定に基づくログレベルの制御
```vb
Public Sub ConfigureLogging()
    Dim logLevel As String
    Dim logPath As String
    
    ' 設定から値を取得
    logLevel = modConfigManager.GetValue("logging", "level", "INFO")
    logPath = modConfigManager.GetValue("logging", "path", "logs/app.log")
    
    ' ログ設定を更新
    modLogger.Configure logPath, _
                       GetLogLevelFromString(logLevel)
End Sub
```

## 初期化順序

システム基盤層のモジュールは以下の順序で初期化する必要があります：

1. 設定マネージャー（modConfigManager）
2. ログマネージャー（modLogger）
3. エラーハンドラー（modErrorHandler）

```vb
Public Sub InitializeSystem()
    ' 1. 設定の初期化
    modConfigManager.Initialize
    
    ' 2. ログの初期化（設定値を使用）
    modLogger.Initialize
    modLogger.Configure modConfigManager.GetValue("logging", "path"), _
                       modConfigManager.GetValue("logging", "level")
    
    ' 3. エラーハンドラーの準備は自動的に行われます
    
    ' 初期化完了のログ
    modLogger.Info "システム基盤層の初期化が完了しました。", "SystemInitialization"
End Sub
```

## エラーコード体系

システム基盤層で使用するエラーコードの範囲：

- vbObjectError + 513: 設定ファイル読み込みエラー
- vbObjectError + 514: デフォルト設定作成エラー
- vbObjectError + 515: 無効なセクション名エラー
- vbObjectError + 516: 設定値更新エラー
- vbObjectError + 517: 設定保存エラー

## 設定ファイル構造

```json
{
    "logging": {
        "level": "INFO",
        "path": "logs/app.log",
        "maxSize": 5242880,
        "rotateCount": 5
    },
    "database": {
        "server": "",
        "database": "",
        "username": "",
        "password": ""
    },
    "security": {
        "encryptionKey": "",
        "sessionTimeout": 30
    },
    "ui": {
        "theme": "default",
        "language": "ja"
    }
}
```

## セキュリティ考慮事項

1. 機密情報の取り扱い
   - パスワードなどの機密情報は暗号化して保存
   - ログにセンシティブな情報を出力しない
   - エラーメッセージは一般ユーザー向けに抽象化

2. エラー情報の制御
   - 本番環境では詳細なエラー情報を非表示
   - エラーログは適切なアクセス制御の下で管理
   - スタックトレースは開発環境でのみ表示

## 保守性とパフォーマンス

1. ログローテーション
   - ファイルサイズの制限
   - 古いログの自動アーカイブ
   - パフォーマンスへの影響を最小化

2. 設定のキャッシュ
   - メモリ内でのキャッシュ
   - 必要な場合のみファイルI/O
   - 変更監視による自動更新

## 今後の拡張計画

1. モジュールの追加
   - 国際化対応（modI18N）
   - キャッシュ管理（modCache）
   - メトリクス収集（modMetrics）

2. 機能強化
   - 非同期ログ出力
   - 設定の暗号化
   - リモート設定管理
