- `clsLogger` : モジュール名
  - `[概要]` : アプリケーションのログ記録を管理するメインクラス。様々な出力先（ファイル、データベース、イベントログ、コンソール）へのログ出力をサポートする。
  - `[依存関係]` :
    - ILoggerSettings
    - IQueue
    - ILock
    - IPerformanceMonitor
    - FileLogger
    - DatabaseLogger
    - EventLogLogger
    - ConsoleLogger
    - ErrorInfo
    - modError
    - modStackTrace
    - clsPerformanceMonitor
    - QueueImpl
  - `[メソッド一覧]` :
    - `Configure(settings As ILoggerSettings, queue As IQueue, lock As ILock, performanceMonitor As IPerformanceMonitor)` : ロガーの設定、キュー、ロック、パフォーマンスモニターを構成する。
    - `SetLogger(destination As String, logger As ILogger)` : 特定の出力先に対するロガーインスタンスを設定する。
    - `ILogger_Log(moduleName As String, message As String, Optional errorCode As ErrorCode)` : ログメッセージをキューに追加する。
    - `ILogger_LogLevel` : ログレベルの設定・取得
    - `ILogger_LogDestination` : ログ出力先の設定・取得
    - `InitializeLoggers()` : 設定に基づいてロガーを初期化する
    - `ProcessLogQueue()` : ログキューを処理し、各ロガーにログメッセージを渡す
    - `CreateLogger(destination As String)` : 指定された出力先に対応するロガーオブジェクトを作成する
    - `StartWorkerThread()` : ワーカースレッドを作成し、開始する
  - `[その他特記事項]` :
    - テスト用の内部メソッド（`GetQueueCount`, `ClearQueue`, `GetPerformanceMonitor`, `SimulateTimer`）は本番環境では使用しないこと。
    - `ProcessLogQueue` メソッドはタイマーによって定期的に呼び出され、ログキューからメッセージを取り出して処理する。
    - `EMERGENCY_LOG_PATH` 定数で指定されたパスに緊急時用のログファイルが出力される。
    - `Logged` イベントはログメッセージが処理された際に発生する。
    - `MAX_RETRY_COUNT` 定数でデータベースおよびイベントログへの書き込みリトライ回数を設定する。
    - `RETRY_DELAY_MS` 定数でデータベースおよびイベントログへの書き込みリトライ間隔を設定する。

- `clsPerformanceMonitor` : モジュール名
  - `[概要]` : パフォーマンス測定機能を提供するレガシークラス。時間計測、実行中の操作の記録、計測結果の取得、有効化・無効化などの機能を持つ。
  - `[依存関係]` :
    - IPerformanceMonitor
    - PerformanceMonitorImpl
  - `[メソッド一覧]` :
    - `StartMeasurement(description As String)` : パフォーマンス測定を開始する。
    - `EndMeasurement(description As String)` : パフォーマンス測定を終了する。
    - `GetMeasurement(description As String)` : 指定された測定対象の測定結果を取得する。
    - `GetAllMeasurements()` : 全ての測定結果を取得する。
    - `ClearMeasurements()` : 全ての測定結果をクリアする。
  - `[その他特記事項]` :
    - 新規実装では `IPerformanceMonitor` インターフェースを使用することを推奨。
    - `IsEnabled` プロパティでパフォーマンス監視の有効・無効を切り替え可能。
    - `CurrentOperation` プロパティで現在実行中の操作を取得可能。

- `ConfigImpl` : モジュール名
  - `[概要]` : アプリケーション設定とデータベース設定を管理するクラス。INIファイルから設定を読み込み、保存する機能を提供する。
  - `[依存関係]` :
    - IAppConfig
    - IDatabaseConfig
    - IIniFile
    - IniFileImpl
    - clsLock
    - IConnectionStringBuilder
    - ODBCConnectionStringBuilder
    - OLEDBConnectionStringBuilder
    - ICryptography
    - IPerformanceMonitor
    - IFileOperations
    - CryptographyImpl
    - FileSystemOperations
    - PerformanceMonitorImpl
    - ValidationResult
  - `[メソッド一覧]` :
    - `IAppConfig_GetSetting(settingName As String, Optional options As SettingOptions = soNone)` : 指定された設定名の値を取得する。
    - `IAppConfig_SetSetting(settingName As String, settingValue As Variant, Optional options As SettingOptions = soNone)` : 指定された設定名の値を設定する。
    - `IAppConfig_LoadSettings(filePath As String, Optional options As LoadOptions = loNone)` : 指定されたファイルから設定を読み込む。
    - `IAppConfig_SaveSettings(filePath As String, Optional options As SaveOptions = soNone)` : 指定されたファイルに設定を保存する。
    - `IAppConfig_SetEncryptionKey(encryptionKey As String, Optional options As EncryptionOptions = eoNone)` : 暗号化キーを設定する。
    - `IAppConfig_SetEnvironment(environment As String, Optional options As EnvironmentOptions = enoNone)` : 環境を設定する。
    - `IAppConfig_ValidateSettings(Optional options As ValidationOptions = voNone)` : 設定を検証する。
    - `IAppConfig_GetSettingHistory(Optional settingName As String, Optional options As HistoryOptions = hoNone)` : 設定の変更履歴を取得する。
    - `IAppConfig_CreateBackup(backupPath As String)` : 設定のバックアップを作成する。
    - `IAppConfig_RestoreFromBackup(backupPath As String, Optional options As RestoreOptions = roNone)` : バックアップから設定を復元する。
    - `IAppConfig_GetPerformanceMetrics()` : パフォーマンスメトリクスを取得する。
    - `IDatabaseConfig_GetConnectionString()` : データベース接続文字列を取得する。
    - `IDatabaseConfig_GetDatabaseSetting(settingName As String)` : 指定されたデータベース設定名の値を取得する。
    - `UpdateConnectionStringBuilder()` : データベースタイプに基づいてConnectionStringBuilderを更新する。
  - `[その他特記事項]` :
    - `AutoSave` プロパティが `True` の場合、設定変更時に自動的にファイルに保存される。
    - `mConnectionStringBuilder` は `IDatabaseConfig_DatabaseType` プロパティの値によって `ODBCConnectionStringBuilder` または `OLEDBConnectionStringBuilder` に切り替わる。
    - `AddSettingHistory`メソッドで設定の変更履歴を記録する。
    - `IsEncrypted`メソッドで設定値が暗号化されているかどうかを判定する。
    - `EncryptValue`、`DecryptValue`メソッドで設定値の暗号化と復号を行う。

- `ConsoleLogger` : モジュール名
  - `[概要]` : コンソールへのログ出力を担当するクラス。
  - `[依存関係]` :
    - ILogger
    - IPerformanceMonitor
    - ErrorInfo
    - modError
  - `[メソッド一覧]` :
    - `Initialize(performanceMonitor As IPerformanceMonitor, Optional useColors As Boolean = False)` : ConsoleLoggerを初期化する。
    - `ILogger_Log(logLevel As String, message As String, Optional stackTrace As String, Optional errorCode As Long)` : ログメッセージをコンソールに出力する。
  - `[その他特記事項]` :
    - `mUseColors` が `True` の場合、ログレベルに応じて色付きで出力される。
    - ANSIエスケープシーケンスを使用して色を表現しているが、現状では定数が空文字で定義されているため、この機能は実質的に無効化されている。

- `DatabaseLogger` : モジュール名
  - `[概要]` : データベースへのログ出力を担当するクラス。
  - `[依存関係]` :
    - ILogger
    - IConnectionPool
    - ILock
    - IPerformanceMonitor
    - IDatabaseConfig
    - ErrorInfo
    - modError
  - `[メソッド一覧]` :
    - `Initialize(connectionPool As IConnectionPool, lock As ILock, performanceMonitor As IPerformanceMonitor, databaseConfig As IDatabaseConfig, tableName As String)` : DatabaseLoggerを初期化する。
    - `ILogger_Log(logLevel As String, message As String, Optional stackTrace As String, Optional errorCode As Long)` : ログメッセージをデータベースに書き込む。
    - `WriteLogToDatabase(logLevel As String, message As String, stackTrace As String, errorCode As Long)` : ログメッセージをデータベースに書き込む。(プライベートメソッド)
  - `[その他特記事項]` :
    - `MAX_RETRY_COUNT` 定数でデータベースへの書き込みリトライ回数を設定する。
    - `RETRY_DELAY_MS` 定数でデータベースへの書き込みリトライ間隔を設定する。
    - ログメッセージは `ADODB.Command` オブジェクトを使用してデータベースに挿入される。

- `DefaultLoggerSettings` : モジュール名
  - `[概要]` : ロガーのデフォルト設定を提供するクラス。
  - `[依存関係]` :
    - ILoggerSettings
    - IAppConfig
    - IFileOperations
    - IConnectionPool
    - IEventLog
    - IDatabaseConfig
  - `[メソッド一覧]` :
    - `Initialize(appConfig As IAppConfig, fileOperations As IFileOperations, connectionPool As IConnectionPool, eventLog As IEventLog, databaseConfig As IDatabaseConfig)` : DefaultLoggerSettingsを初期化する。
    - `ILoggerSettings_LogLevel` : ログレベルの設定・取得
    - `ILoggerSettings_GetLogDestinations()` : 有効なログ出力先のリストを取得する。
    - `ILoggerSettings_LogFilePath` : ログファイルのパスを取得する。
    - `ILoggerSettings_LogTableName` : ログを保存するデータベースのテーブル名を取得する。
    - `ILoggerSettings_LogEventSource` : イベントログのソース名を取得する。
    - `ILoggerSettings_TimerInterval` : タイマーの間隔を取得する。
    - `ILoggerSettings_GetFileOperations()` : ファイル操作オブジェクトを取得する。
    - `ILoggerSettings_GetConnectionPool()` : データベース接続プールオブジェクトを取得する。
    - `ILoggerSettings_GetEventLog()` : イベントログオブジェクトを取得する。
    - `ILoggerSettings_GetDatabaseConfig()` : データベース設定オブジェクトを取得する。
    - `ILoggerSettings_ShouldLog(destination As String, level As LogLevel)` : 指定された出力先とログレベルでログを出力すべきかどうかを判定する。
    - `ILoggerSettings_FormatLogMessage(logLevel As String, message As String, stackTrace As String, errorCode As Long)` : ログメッセージをフォーマットする。
  - `[その他特記事項]` :
    - `LoadSettings` メソッドで `IAppConfig` から設定を読み込む。
    - `ParseDestinations` メソッドで有効なログ出力先をパースする。

- `DefaultMessageFormatter` : モジュール名
  - `[概要]` : デフォルトのエラーメッセージフォーマットを提供するクラス。
  - `[依存関係]` :
    - IMessageFormatter
    - ErrorInfo
    - modStackTrace
    - modError
  - `[メソッド一覧]` :
    - `IMessageFormatter_FormatMessage(message As String, ByRef errorInfo As ErrorInfo)` : エラーメッセージをフォーマットする。
  - `[その他特記事項]` :
    - `FormatMessage` メソッドは、エラー情報オブジェクトの内容に基づいてエラーメッセージを組み立てる。
    - スタックトレースが存在する場合は、メッセージに追加される。
    - メッセージのフォーマット中にエラーが発生した場合は、エラーハンドラに処理を委譲する。
    - テスト用の `TestFormatMessage` メソッドを持つ（DEBUG時のみ）。

- `EmergencyLogger` : モジュール名
  - `[概要]` : 緊急時（エラーハンドラでエラーが発生した場合など）にログを出力するクラス。
  - `[依存関係]` :
    - IEmergencyLogger
    - clsLock
    - modWindowsAPI
  - `[メソッド一覧]` :
    - `IEmergencyLogger_LogEmergencyError(message As String, Optional ByRef errorInfo As ErrorInfo)` : 緊急エラーをログに記録する。
    - `FormatErrorMessage(message As String, ByRef errorInfo As ErrorInfo)` : エラーメッセージをフォーマットする。（プライベートメソッド）
    - `WriteToEventLog(message As String)` : イベントログに書き込む。（プライベートメソッド）
    - `WriteToEmergencyFile(message As String)` : 緊急用ファイルに書き込む。（プライベートメソッド）
  - `[その他特記事項]` :
    - `EMERGENCY_LOG_FILE` 定数で緊急用ログファイルのパスを指定する。
    - `EVENT_SOURCE` 定数でイベントログのソース名を指定する。
    - テスト用の `GetEmergencyLogPath` および `ClearEmergencyLog` メソッドを持つ（DEBUG時のみ）。

- `ErrorHandlerManager` : モジュール名
  - `[概要]` : エラーハンドラーの登録・管理を行うクラス。
  - `[依存関係]` :
    - IErrorHandler
    - clsLock
    - DatabaseConnectionErrorHandler
    - FileNotFoundErrorHandler
    - InvalidInputErrorHandler
    - modErrorCodes
  - `[メソッド一覧]` :
    - `InitializeManager()` : エラーハンドラーマネージャーを初期化する。
    - `RegisterHandler(errorCode As ErrorCode, handler As IErrorHandler)` : エラーコードに対応するエラーハンドラーを登録する。
    - `UnregisterHandler(errorCode As ErrorCode)` : エラーコードに対応するエラーハンドラーの登録を解除する。
    - `GetErrorHandler(errorCode As ErrorCode)` : エラーコードに対応するエラーハンドラーを取得する。
    - `RegisterDefaultHandlers()` : デフォルトのエラーハンドラーを登録する。（プライベートメソッド）
    - `GetDefaultHandlerForCategory(category As ErrorCodeCategory)` : エラーカテゴリに対応するデフォルトのエラーハンドラーを取得する。（プライベートメソッド）
  - `[その他特記事項]` :
    - `mHandlers` コレクションにエラーコードとエラーハンドラーのペアを保持する。
    - テスト用の `GetHandlerCount` および `ClearHandlers` メソッドを持つ（DEBUG時のみ）。

- `ErrorImpl` : モジュール名
  - `[概要]` : エラー処理の主機能を担当するクラス。エラーハンドラーの呼び出し、エラー情報の補完、エラー回数のカウント、エラー履歴の管理、エラー分析、パフォーマンスメトリクスの収集などを行う。
  - `[依存関係]` :
    - IError
    - ErrorHandlerManager
    - IEmergencyLogger
    - EmergencyLogger
    - clsLock
    - clsPerformanceMonitor
    - ErrorInfo
    - modErrorCodes
    - modStackTrace
  - `[メソッド一覧]` :
    - `IError_HandleError(ByRef errorInfo As ErrorInfo, Optional ByVal options As HandlingOptions)` : エラーを処理する。
    - `IError_HandleBatchErrors(ByVal errors As Collection, Optional ByVal options As BatchOptions)` : 複数のエラーを一括で処理する。
    - `IError_AnalyzeErrors(Optional ByVal options As AnalysisOptions)` : エラーを分析する。
    - `IError_GetErrorHistory(Optional ByVal options As HistoryOptions)` : エラー履歴を取得する。
    - `IError_GetPerformanceMetrics()` : パフォーマンスメトリクスを取得する。
    - `CompleteErrorInfo(ByRef errorInfo As ErrorInfo)` : エラー情報オブジェクトの不足情報を補完する。（プライベートメソッド）
    - `AddToHistory(ByRef errorInfo As ErrorInfo)` : エラー情報を履歴に追加する。（プライベートメソッド）
    - `RemoveOldestHistoryEntry()` : エラー履歴から最も古いエントリを削除する。（プライベートメソッド）
    - `PerformCleanup()` : 定期的なクリーンアップ処理を行う。（プライベートメソッド）
    - `CleanupErrorCounts()` : エラーカウントのクリーンアップを行う。（プライベートメソッド）
    - `CleanupErrorHistory()` : エラー履歴のクリーンアップを行う。（プライベートメソッド）
    - `GetErrorSeverity(ByRef errorInfo As ErrorInfo)` : エラーの重要度を判定する。（プライベートメソッド）
    - `AnalyzeError(ByRef errorInfo As ErrorInfo)` : エラーの分析を行う。（プライベートメソッド）
    - `CategorizeError(ByRef errorInfo As ErrorInfo)` : エラーの分類を行う。（プライベートメソッド）
    - `AnalyzeErrorHistory(ByRef result As ErrorAnalysisResult)` : エラー履歴の分析を行う。（プライベートメソッド）
    - `DetectErrorPatterns(ByRef result As ErrorAnalysisResult)` : エラーパターンの検出を行う。（プライベートメソッド）
    - `AnalyzeErrorTrends(ByRef result As ErrorAnalysisResult)` : エラー傾向の分析を行う。（プライベートメソッド）
    - `GetErrorCountMetrics()` : エラーカウントのメトリクス取得を行う。（プライベートメソッド）
    - `GetHandlerPerformanceMetrics()` : ハンドラーのパフォーマンスメトリクス取得を行う。（プライベートメソッド）
    - `GetMemoryUsageMetrics()` : メモリ使用量のメトリクス取得を行う。（プライベートメソッド）
  - `[その他特記事項]` :
    - `MAX_ERROR_RECURSION` 定数でエラー処理の再帰呼び出しの上限回数を設定する。
    - `ERROR_COUNT_DICT_SIZE` 定数でエラーカウントを保持する辞書の最大サイズを設定する。
    - `MAX_HISTORY_SIZE` 定数でエラー履歴の最大サイズを設定する。
    - `MAX_BATCH_SIZE` 定数で一括処理するエラーの最大数を設定する。
    - `CLEANUP_INTERVAL_MS` 定数でクリーンアップ処理の間隔を設定する。
    - `MAX_RECOVERY_ATTEMPTS` 定数で回復試行の最大回数を設定する。
    - `ANALYSIS_INTERVAL_MS` 定数で分析処理の間隔を設定する。
    - `mErrorCounts` 辞書にエラーコードごとの発生回数を記録する。
    - `mErrorHistory` コレクションにエラー履歴を格納する。
    - `mErrorCategories` 辞書にエラーコードとエラーカテゴリの対応を記録する。
    - `ErrorOccurred` イベントでエラーの発生を通知する。
    - `ErrorHandled` イベントでエラーが処理されたことを通知する。
    - `ErrorAnalysisCompleted` イベントでエラー分析が完了したことを通知する。
    - `ErrorCategoryChanged` イベントでエラーカテゴリが変更されたことを通知する。
    - `ErrorThresholdExceeded` イベントでエラーの発生回数が閾値を超えたことを通知する。
    - `RecoveryAttempted` イベントで回復処理が試行されたことを通知する。
    - `BatchProcessed` イベントで一括処理が完了したことを通知する。
    - `PerformanceAlert` イベントでパフォーマンスの問題を通知する。
    - `ResourceExhausted` イベントでリソース不足を通知する。
    - テスト用の `GetErrorCount`, `ClearErrorCounts`, `GetPerformanceReport` メソッドを持つ（DEBUG時のみ）。

- `ErrorInfo` : モジュール名
  - `[概要]` : エラー情報を格納するクラス。エラーコード、説明、発生源、スタックトレースなどの情報を持つ。
  - `[依存関係]` : なし
  - `[メソッド一覧]` :
    - `AddAdditionalInfo(key As String, value As Variant)` : エラー情報に追加情報を追加する。
    - `GetAdditionalInfo(key As String)` : エラー情報から追加情報を取得する。
    - `HasAdditionalInfo(key As String)` : エラー情報に追加情報が存在するかどうかを確認する。
    - `Clone()` : エラー情報オブジェクトのコピーを作成する。
    - `ToString()` : エラー情報オブジェクトを文字列形式で返す。
  - `[その他特記事項]` :
    - `ErrorSeverity` 列挙型でエラーの重要度を定義する。
    - `mAdditionalInfo` コレクションに追加情報を格納する。
    - `Code`, `Description`, `Category`, `Source`, `ProcedureName`, `StackTrace`, `OccurredAt`, `Severity`, `InnerError`, `RecoveryAttempted`, `RecoverySuccessful` の各プロパティを持つ。

- `EventLogImpl` : モジュール名
  - `[概要]` : Windowsイベントログへの書き込みを担当するクラス。イベントソースの確認、メッセージの整形、バッチ処理、フィルター機能、バックアップ作成、メトリクス収集などの機能を持つ。
  - `[依存関係]` :
    - IEventLog
    - clsLock
    - clsPerformanceMonitor
    - modWindowsAPI
    - SecurityContext
  - `[メソッド一覧]` :
    - `IEventLog_WriteToEventLog(source As String, message As String, eventType As EventLogType, Optional options As WriteOptions)` : イベントログに書き込む。
    - `IEventLog_WriteBatch(entries As Collection, Optional options As BatchOptions)` : 複数のイベントログエントリを一括で書き込む。
    - `IEventLog_FilterEvents(criteria As String, Optional options As FilterOptions)` : イベントログをフィルタリングする。
    - `IEventLog_CreateBackup(Optional options As BackupOptions)` : イベントログのバックアップを作成する。
    - `IEventLog_GetPerformanceMetrics()` : パフォーマンスメトリクスを取得する。
    - `ValidateSecurityContext(source As String)` : セキュリティコンテキストを検証する。（プライベートメソッド）
    - `LogSecurityAlert(alertType As String, details As String)` : セキュリティアラートをログに記録する。（プライベートメソッド）
    - `VerifyEventSource(source As String)` : イベントソースの存在を確認する。（プライベートメソッド）
    - `ValidateAndFormatMessage(message As String)` : メッセージを検証・整形する。（プライベートメソッド）
    - `WriteEventLogWithRetry(source As String, message As String, eventType As EventLogType, options As WriteOptions)` : リトライ機能付きでイベントログに書き込む。（プライベートメソッド）
    - `PerformPeriodicTasks()` : 定期的なタスクを実行する。（プライベートメソッド）
    - `CollectMetrics()` : メトリクスを収集する。（プライベートメソッド）
    - `PerformCleanup()` : クリーンアップ処理を行う。（プライベートメソッド）
    - `CreateEventLogBackup(options As BackupOptions)` : イベントログのバックアップを作成する。（プライベートメソッド）
    - `ApplyEventFilters(results As Collection, criteria As String, options As FilterOptions)` : フィルターを適用する。（プライベートメソッド）
    - `GetEventLogMetrics()` : イベントログ関連のメトリクスを取得する。（プライベートメソッド）
    - `GetCacheMetrics()` : キャッシュ関連のメトリクスを取得する。（プライベートメソッド）
    - `GetSecurityMetrics()` : セキュリティ関連のメトリクスを取得する。（プライベートメソッド）
  - `[その他特記事項]` :
    - `MAX_MESSAGE_LENGTH` 定数でメッセージの最大長を設定する。
    - `MAX_BATCH_SIZE` 定数で一括処理するエントリの最大数を設定する。
    - `CACHE_DURATION_MS` 定数でキャッシュの有効期間を設定する。
    - `MAX_RETRY_COUNT` 定数でリトライの最大回数を設定する。
    - `BACKUP_INTERVAL_MS` 定数でバックアップの実行間隔を設定する。
    - `METRICS_INTERVAL_MS` 定数でメトリクス収集の間隔を設定する。
    - `CLEANUP_INTERVAL_MS` 定数でクリーンアップ処理の間隔を設定する。
    - `mSourceCache` 辞書にイベントソースの存在確認結果をキャッシュする。
    - `mEventFilters` コレクションにイベントログのフィルターを格納する。
    - `mSecurityContext` メンバー変数でセキュリティコンテキストを管理する。
    - `EventLogged` イベントでイベントログが書き込まれたことを通知する。
    - `BatchProcessed` イベントで一括処理が完了したことを通知する。
    - `SourceRegistered` イベントでイベントソースが登録されたことを通知する。
    - `BackupCreated` イベントでバックアップが作成されたことを通知する。
    - `FilterApplied` イベントでフィルターが適用されたことを通知する。
    - `SecurityAlert` イベントでセキュリティアラートを通知する。
    - `PerformanceAlert` イベントでパフォーマンスの問題を通知する。
    - `ResourceExhausted` イベントでリソース不足を通知する。
    - テスト用の `ValidateState`, `GetPerformanceMonitor`, `TestEventLogAccess` メソッドを持つ（DEBUG時のみ）。

- `EventLogLogger` : モジュール名
  - `[概要]` : イベントログへのログ出力を担当するクラス。
  - `[依存関係]` :
    - ILogger
    - IEventLog
    - ILock
    - IPerformanceMonitor
    - ErrorInfo
    - modError
  - `[メソッド一覧]` :
    - `Initialize(eventLog As IEventLog, lock As ILock, performanceMonitor As IPerformanceMonitor, eventSource As String)` : EventLogLoggerを初期化する。
    - `ILogger_Log(logLevel As String, message As String, Optional stackTrace As String, Optional errorCode As Long)` : ログメッセージをイベントログに書き込む。
    - `WriteLogToEventLog(logLevel As String, message As String, stackTrace As String, errorCode As Long)` : ログメッセージをイベントログに書き込む。（プライベートメソッド）
  - `[その他特記事項]` :
    - `MAX_RETRY_COUNT` 定数でイベントログへの書き込みリトライ回数を設定する。
    - `RETRY_DELAY_MS` 定数でイベントログへの書き込みリトライ間隔を設定する。
    - ログレベルに応じてイベントの種類（`EVENTLOG_SUCCESS`, `EVENTLOG_ERROR`, `EVENTLOG_WARNING`, `EVENTLOG_INFORMATION`）を決定する。

- `FileLogger` : モジュール名
  - `[概要]` : ファイルへのログ出力を担当するクラス。
  - `[依存関係]` :
    - ILogger
    - IFileOperations
    - ILock
    - IPerformanceMonitor
    - ErrorInfo
    - modError
  - `[メソッド一覧]` :
    - `Initialize(fileOperations As IFileOperations, lock As ILock, performanceMonitor As IPerformanceMonitor, logFilePath As String)` : FileLoggerを初期化する。
    - `ILogger_Log(logLevel As String, message As String, Optional stackTrace As String, Optional errorCode As Long)` : ログメッセージをファイルに書き込む。
    - `BuildLogMessage(logLevel As String, message As String, stackTrace As String, errorCode As Long)` : ログメッセージを組み立てる。（プライベートメソッド）
    - `WriteLogToFile(logMessage As String)` : ログメッセージをファイルに書き込む。（プライベートメソッド）
  - `[その他特記事項]` :
    - `MAX_RETRY_COUNT` 定数でファイルへの書き込みリトライ回数を設定する。
    - `RETRY_DELAY_MS` 定数でファイルへの書き込みリトライ間隔を設定する。
    - ログファイルが存在しない場合は `CreateFile` メソッドで作成される。

- `InvalidInputErrorHandler` : モジュール名
  - `[概要]` : 不正な入力エラーを処理するエラーハンドラー。入力値の自動補正とユーザーへの通知を行う。
  - `[依存関係]` :
    - IErrorHandler
    - ILock
    - ILogger
    - IEmergencyLogger
    - IUserNotifier
    - IValidator
    - ErrorInfo
    - modStackTrace
    - modError
  - `[メソッド一覧]` :
    - `Create(lock As ILock, logger As ILogger, emergencyLogger As IEmergencyLogger, userNotifier As IUserNotifier, validator As IValidator)` : InvalidInputErrorHandlerのインスタンスを作成する。
    - `IErrorHandler_HandleError(ByRef errorDetail As ErrorInfo)` : 不正な入力エラーを処理する。
    - `TryCorrectInput(value As Variant, inputType As String, ByRef correctedValue As Variant)` : 入力値の自動補正を試みる。（プライベートメソッド）
    - `TryCorrectNumber(value As Variant, ByRef correctedValue As Variant)` : 数値型の入力値の自動補正を試みる。（プライベートメソッド）
    - `TryCorrectDate(value As Variant, ByRef correctedValue As Variant)` : 日付型の入力値の自動補正を試みる。（プライベートメソッド）
    - `TryCorrectString(value As Variant, ByRef correctedValue As Variant)` : 文字列型の入力値の自動補正を試みる。（プライベートメソッド）
    - `LogError(ByRef errorDetail As ErrorInfo)` : エラーをログに記録する。（プライベートメソッド）
    - `NotifyUser(ByRef errorDetail As ErrorInfo, style As VbMsgBoxStyle)` : ユーザーにエラーを通知する。（プライベートメソッド）
  - `[その他特記事項]` :
    - コンストラクタ `Create` を使用して、依存オブジェクトを注入する。
    - 入力値の型に応じて `TryCorrectNumber`, `TryCorrectDate`, `TryCorrectString`を使い分けて自動補正を試みる。
    - 自動補正に成功した場合は、エラー情報オブジェクトに `CorrectedValue` として補正後の値を追加する。
    - エラーの重要度が `ESError` 以上の場合、緊急ログにも記録する。

- `modAppInitializer` : モジュール名
  - `[概要]` : アプリケーションの初期化処理を担当するモジュール。設定の読み込み、ロギングシステムの初期化、データベースの初期化、エラーハンドラーの初期化を順に行う。
  - `[依存関係]` :
    - modConfig
    - clsLogger
    - DefaultLoggerSettings
    - modDatabaseUtils
    - ConnectionPool
    - DatabaseConnectionErrorHandler
    - ErrorInfo
    - modError
    - modStackTrace
  - `[メソッド一覧]` :
    - `InitializeApplication()` : アプリケーションを初期化する。
    - `InitializeLogging()` : ロギングシステムを初期化する。（プライベートメソッド）
    - `InitializeDatabase()` : データベース関連の初期化を行う。（プライベートメソッド）
    - `InitializeErrorHandlers()` : エラーハンドラーを初期化する。（プライベートメソッド）
  - `[その他特記事項]` :
    - `InitializationState` 型変数 `mInitState` で各コンポーネントの初期化状態を管理する。
    - `InitializeApplication` メソッドは、設定、ロギング、データベース、エラーハンドラーの順に初期化を行う。

- `modConfig` : モジュール名
  - `[概要]` : アプリケーションの設定を管理するモジュール。設定の読み込み、保存、デフォルト値の提供、データベース接続文字列の提供などを行う。
  - `[依存関係]` :
    - IDatabaseConfig
    - modWindowsAPI
    - clsLock
    - clsPerformanceMonitor
    - ErrorInfo
    - modError
    - clsCallStack
  - `[メソッド一覧]` :
    - `InitializeModule()` : モジュールを初期化する。
    - `TerminateModule()` : モジュールを終了処理する。
    - `GetConfigValue(section As String, key As String, Optional defaultValue As String = "")` : 設定ファイルから値を取得する。
    - `SetConfigValue(section As String, key As String, Value As String)` : 設定ファイルに値を設定する。
    - `LoadDefaultSettings()` : デフォルトの設定値を読み込む。（プライベートメソッド）
    - `LoadConfigurationFromFile()` : 設定ファイルから設定を読み込む。（プライベートメソッド）
    - `SaveConfigurationToFile()` : 設定ファイルに設定を保存する。（プライベートメソッド）
    - `GetConfigFilePath()` : 設定ファイルのパスを取得する。（プライベートメソッド）
    - `SaveChanges()` : 変更された設定を保存する。
    - `GetCurrentCallStack()` : 現在のコールスタックを取得する。（プライベートメソッド）
    - `IDatabaseConfig_GetConnectionString()` : データベース接続文字列を取得する。
  - `[その他特記事項]` :
    - `CONFIG_FILE_PATH` 定数で設定ファイルのパスを指定する。
    - `MAX_BUFFER_SIZE` 定数で設定値の最大バッファサイズを指定する。
    - `DEFAULT_SECTION` 定数でデフォルトのセクション名を指定する。
    - `ConfigurationSettings` 型変数 `settings` に設定値を保持する。
    - `AutoSave` プロパティで設定変更時に自動的に保存するかどうかを制御する。
    - `HasUnsavedChanges` プロパティで未保存の変更があるかどうかを確認できる。
    - テスト用の `ResetModule` および `ValidateSettings` メソッドを持つ（DEBUG時のみ）。

- `modError` : モジュール名
  - `[概要]` : エラー処理機能を提供するモジュール。エラーハンドラーの呼び出し、エラー情報の補完、緊急ログ出力などを行う。
  - `[依存関係]` :
    - IErrorHandler
    - clsLock
    - clsPerformanceMonitor
    - DatabaseConnectionErrorHandler
    - FileNotFoundErrorHandler
    - InvalidInputErrorHandler
    - ErrorInfo
    - modErrorCodes
    - modStackTrace
    - modWindowsAPI
  - `[メソッド一覧]` :
    - `InitializeModule()` : モジュールを初期化する。
    - `TerminateModule()` : モジュールを終了処理する。
    - `TryHandleError(ByRef context As ErrorContext)` : エラー処理を試行する。（プライベートメソッド）
    - `HandleError(ByRef errInfo As ErrorInfo)` : エラーを処理する。
    - `RegisterDefaultHandlers()` : デフォルトのエラーハンドラーを登録する。（プライベートメソッド）
    - `GetErrorHandler(errorCode As ErrorCode)` : エラーコードに対応するエラーハンドラーを取得する。（プライベートメソッド）
    - `GetDefaultHandlerForCategory(category As ErrorCodeCategory)` : エラーカテゴリに対応するデフォルトのエラーハンドラーを取得する。（プライベートメソッド）
    - `RegisterErrorHandler(errorCode As ErrorCode, handler As IErrorHandler)` : エラーハンドラーを登録する。
    - `UnregisterErrorHandler(errorCode As ErrorCode)` : エラーハンドラーの登録を解除する。
    - `EmergencyErrorLog(message As String)` : 緊急エラーをログに記録する。（プライベートメソッド）
    - `WriteToEventLog(message As String)` : イベントログに書き込む。（プライベートメソッド）
    - `WriteToEmergencyFile(message As String)` : 緊急用ファイルに書き込む。（プライベートメソッド）
  - `[その他特記事項]` :
    - `MAX_ERROR_RECURSION` 定数でエラー処理の再帰呼び出しの上限回数を設定する。
    - `ErrorContext` 型変数 `context` にエラー処理に必要な情報を格納する。
    - テスト用の `TestErrorHandling`, `TestErrorRecursion`, `TestResourceManagement`, `GetPerformanceReport`, `GetRegisteredHandlerCount`, `ClearHandlers`, `ResetModule`, `GetActiveLockCount` メソッドを持つ（DEBUG時のみ）。

- `modErrorCodes` : モジュール名
  - `[概要]` : エラーコードとエラーカテゴリを定義するモジュール。
  - `[依存関係]` : なし
  - `[メソッド一覧]` :
    - `GetErrorCategory(errCode As ErrorCode)` : エラーコードに対応するエラーカテゴリを取得する。
  - `[その他特記事項]` :
    - `ErrorCodeCategory` 列挙型でエラーカテゴリを定義する。
    - `ErrorCode` 列挙型でエラーコードを定義する。

- `PerformanceCounterImpl` : モジュール名
  - `[概要]` : パフォーマンスカウンターへのアクセスを提供するクラス。
  - `[依存関係]` :
    - IPerformanceCounter
    - clsLock
    - modWindowsAPI
    - ErrorInfo
    - modError
    - modStackTrace
  - `[メソッド一覧]` :
    - `IPerformanceCounter_QueryPerformanceCounter(ByRef performanceCount As Currency)` : パフォーマンスカウンターの値を取得する。
    - `IPerformanceCounter_QueryPerformanceFrequency(ByRef frequency As Currency)` : パフォーマンスカウンターの周波数を取得する。
    - `CheckHighResolutionSupport()` : 高分解能タイマーがサポートされているかどうかを確認する。（プライベートメソッド）
    - `LogError(message As String)` : エラーをログに記録する。（プライベートメソッド）
    - `GetResolution()` : パフォーマンスカウンターの分解能を取得する。
  - `[その他特記事項]` :
    - 高分解能タイマーがサポートされている場合、`mIsHighResolutionSupported` が `True` に設定される。
    - パフォーマンスカウンターの周波数は `mFrequency` にキャッシュされる。
    - テスト用の `ValidateState`, `GetFrequency`, `IsHighResolutionSupported`, `TestTimerConsistency` メソッドを持つ（DEBUG時のみ）。

- `PerformanceMonitorImpl` : モジュール名
  - `[概要]` : パフォーマンス測定機能を提供するクラス。測定の開始・終了、経過時間の取得、メモリ使用量の取得などの機能を持つ。
  - `[依存関係]` :
    - IPerformanceMonitor
    - modWindowsAPI
    - clsLock
    - IAppConfig
    - modConfig
    - ErrorInfo
    - modError
    - modStackTrace
  - `[メソッド一覧]` :
    - `IPerformanceMonitor_Start(measurementName As String)` : パフォーマンス測定を開始する。
    - `IPerformanceMonitor_Stop(measurementName As String)` : パフォーマンス測定を終了する。
    - `IPerformanceMonitor_GetMeasurement(measurementName As String)` : 指定された測定の経過時間（ミリ秒）を取得する。
    - `IPerformanceMonitor_GetAllMeasurements()` : すべての測定の経過時間（ミリ秒）を取得する。
    - `IPerformanceMonitor_Clear()` : すべての測定をクリアする。
    - `IPerformanceMonitor_IsEnabled()` : パフォーマンス監視が有効かどうかを取得する。
    - `IPerformanceMonitor_Enable()` : パフォーマンス監視を有効にする。
    - `IPerformanceMonitor_Disable()` : パフォーマンス監視を無効にする。
    - `GetDetailedMeasurement(measurementName As String)` : 指定された測定の詳細なレポートを取得する。
  - `[その他特記事項]` :
    - `mFrequency` メンバー変数にパフォーマンスカウンターの周波数を保持する。
    - `mMeasurements` コレクションに測定データを格納する。
    - `mIsEnabled` メンバー変数でパフォーマンス監視の有効/無効を管理する。
    - `mCurrentOperation` メンバー変数に現在実行中の操作名を格納する。
    - `PROCESS_MEMORY_COUNTERS` 型を使用して、プロセスのメモリ使用量を取得する。
    - テスト用の `ValidatePerformanceCounter`, `GetMeasurementCount`, `SimulateMeasurement`, `CurrentOperation` メソッド/プロパティを持つ（DEBUG時のみ）。

- `ValidationResult` : モジュール名
  - `[概要]` : バリデーション結果を格納するクラス。検証結果、エラーメッセージ、警告メッセージなどを保持する。
  - `[依存関係]` : なし
  - `[メソッド一覧]` :
    - `AddError(errorMessage As String)` : エラーを追加する。
    - `AddWarning(warningMessage As String)` : 警告を追加する。
    - `AddValidatedRule(ruleName As String, ruleResult As Boolean)` : 検証ルールを追加する。
    - `GetSummary()` : バリデーション結果の要約を取得する。
    - `GetStateName(state As ValidationState)` : バリデーション状態名を取得する。（プライベートメソッド）
  - `[その他特記事項]` :
    - `mIsValid` メンバー変数でバリデーション結果が有効かどうかを保持する。
    - `mErrors` コレクションにエラーメッセージを格納する。
    - `mWarnings` コレクションに警告メッセージを格納する。
    - `mValidatedRules` コレクションに検証されたルールを格納する。
    - `mState` メンバー変数でバリデーションの状態を保持する。
    - `mStartTime` メンバー変数でバリデーションの開始時刻を保持する。
    - `mEndTime` メンバー変数でバリデーションの終了時刻を保持する。
    - `mValidatedSettingCount` メンバー変数で検証された設定の数を保持する。
