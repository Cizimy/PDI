# clsDefaultErrorHandler 修正計画

## 高優先度の修正

### 1. コーディング規約違反の修正

#### XMLドキュメントコメントの追加
- すべてのパブリックメソッドとプロパティにXMLドキュメントコメントを追加
- コメントには`<summary>`, `<param>`, `<returns>`, `<remarks>`を適切に使用

#### 定数名の修正
現在:
```vb
Private Const MILLISECONDS_PER_SECOND As Long = 1000
Private Const JITTER_RANGE As Double = 0.1
Private Const FIBONACCI_INITIAL_VALUE As Long = 0
Private Const BURST_THRESHOLD As Long = 5
Private Const BURST_INTERVAL As Long = 60
Private Const ENV_PREFIX As String = "ERROR_HANDLER_"
```

修正後:
```vb
Private Const MILLISECONDS_PER_SECOND As Long = 1000
Private Const JITTER_RANGE As Double = 0.1
Private Const FIBONACCI_INITIAL_VALUE As Long = 0
Private Const ERROR_BURST_THRESHOLD As Long = 5
Private Const ERROR_BURST_INTERVAL As Long = 60
Private Const ERROR_HANDLER_ENV_PREFIX As String = "ERROR_HANDLER_"
```

### 2. ExecuteRetry メソッドの重複排除
- ExecuteRetryメソッドをAttemptRetryメソッドに統合
- 重複するリトライロジックを一箇所にまとめる

## 中優先度の修正

### 1. マジックナンバーの定数化

#### AttemptRetry メソッド
```vb
Private Const MILLISECONDS_PER_SECOND As Long = 1000  ' 既存の定数を使用
```

#### CalculateBackoffInterval メソッド
```vb
Private Const JITTER_RANGE As Double = 0.1  ' 既存の定数を使用
```

#### CalculateFibonacciInterval メソッド
```vb
Private Const FIBONACCI_INITIAL_VALUE As Long = 0  ' 既存の定数を使用
```

#### LoadSettings メソッド
```vb
Private Const DEFAULT_MAX_RETRY_COUNT As Long = 3
Private Const DEFAULT_RETRY_INTERVAL As Long = 1000
Private Const DEFAULT_MIN_RETRY_INTERVAL As Long = 100
Private Const DEFAULT_MAX_RETRY_INTERVAL As Long = 30000
Private Const DEFAULT_TIMEOUT_MS As Long = 30000
Private Const DEFAULT_CONNECTION_TIMEOUT_MS As Long = 5000
Private Const DEFAULT_OPERATION_TIMEOUT_MS As Long = 30000
Private Const DEFAULT_MAX_ACTIVE_RESOURCES As Long = 100
Private Const DEFAULT_RESOURCE_CLEANUP_INTERVAL As Long = 300000
Private Const DEFAULT_MAX_RESOURCE_AGE_MS As Long = 3600000
Private Const DEFAULT_MAX_LOG_SIZE As Long = 5242880
Private Const DEFAULT_LOG_ROTATION_COUNT As Long = 5
Private Const DEFAULT_ERROR_HISTORY_SIZE As Long = 1000
Private Const DEFAULT_PATTERN_ANALYSIS_WINDOW As Long = 3600
Private Const DEFAULT_METRICS_SAMPLE_INTERVAL As Long = 60000
Private Const DEFAULT_METRICS_HISTORY_SIZE As Long = 1000
Private Const DEFAULT_MAX_RECOVERY_CHAIN_LENGTH As Long = 5
Private Const DEFAULT_RECOVERY_TIMEOUT_MS As Long = 60000
```

### 2. エラー処理の改善
以下のメソッドのエラーハンドリングにErr.Descriptionを含める：

1. BeginTransaction
2. CommitTransaction
3. AttemptRetry
4. ExecuteDefaultRetry
5. LoadSettingsFromFile
6. ValidateSettings
7. ExecuteRetry
8. RetryDatabaseOperation
9. RetryNetworkOperation
10. RetryFileOperation
11. RetryTimedOutOperation

例：
```vb
ErrorHandler:
    LogError "メソッド名", "エラーの説明: " & Err.Description & " (ErrorCode: " & Err.Number & ")"
    Err.Raise Err.Number, "メソッド名", "処理に失敗しました: " & Err.Description
```

## 低優先度の修正

### メソッド名の命名規則修正

現在の名前 | 修正後の名前
-----------|-------------
UpdateRetryStatistics | UpdateRetryStats
ExecuteDefaultRetry | ExecuteDefaultRetryOperation
GetTransactionLevel | GetCurrentTransactionLevel
CalculateBackoffInterval | GetBackoffInterval
CalculateFibonacciInterval | GetFibonacciInterval
RestoreContext | RestoreExecutionContext
LoadSettings | LoadErrorHandlerSettings
LoadSettingsFromFile | LoadSettingsFromConfigFile
LoadSettingsFromEnvironment | LoadSettingsFromEnvVars
CheckEnvironmentSetting | LoadSettingFromEnvironment
ValidateSettings | ValidateErrorHandlerSettings
ValidateRange | ValidateValueRange
InitializeMinimalSettings | InitializeMinimalErrorHandlerSettings
Class_Initialize | Initialize
LogError | WriteErrorLog

## 実装の注意点

1. 各修正は既存の機能を損なわないように慎重に行う
2. 修正後は単体テストを実行して機能の正常性を確認
3. 変更はバックワードコンパティビリティを維持
4. コードの可読性と保守性を向上させることを意識
5. エラーハンドリングの強化により、システムの安定性を向上

## 修正の影響範囲

1. エラーハンドリング機能全般
2. リトライメカニズム
3. トランザクション管理
4. パフォーマンスメトリクス収集
5. 設定管理
6. ログ機能

これらの修正により、コードの品質、保守性、エラー処理の堅牢性が向上します。