VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAsyncValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("非同期バリデーションの実行と管理を担当するインターフェース")

' 定数定義
Private Const DEFAULT_TIMEOUT_MS As Long = 30000        ' デフォルトタイムアウト時間（30秒）
Private Const MAX_CONCURRENT_TASKS As Long = 10         ' 最大同時実行タスク数
Private Const MIN_BATCH_SIZE As Long = 100              ' 最小バッチサイズ
Private Const MAX_QUEUE_SIZE As Long = 1000             ' 最大キューサイズ
Private Const CACHE_DURATION_MS As Long = 300000        ' キャッシュ有効期間（5分）
Private Const RETRY_INTERVAL_MS As Long = 1000          ' リトライ間隔（1秒）
Private Const MAX_RETRY_COUNT As Long = 3               ' 最大リトライ回数

' イベント定義
Public Event ValidationStarted(ByVal validationId As String, ByVal context As IValidationContext, ByVal executionPlan As ValidationExecutionPlan)
Public Event ValidationProgress(ByVal validationId As String, ByVal progressPercentage As Long, ByVal currentOperation As String, ByVal estimatedTimeRemaining As Long)
Public Event ValidationCompleted(ByVal validationId As String, ByVal success As Boolean, ByVal results As Collection, ByVal metrics As ValidationMetrics)
Public Event ValidationError(ByVal validationId As String, ByVal errorInfo As ErrorInfo, ByVal retryCount As Long)
Public Event ValidationCancelled(ByVal validationId As String, ByVal reason As String)
Public Event ResourceExhausted(ByVal validationId As String, ByVal resourceType As String, ByVal currentUsage As Long)
Public Event ValidationQueued(ByVal validationId As String, ByVal queuePosition As Long, ByVal estimatedWaitTime As Long)
Public Event BatchProcessingStarted(ByVal batchId As String, ByVal itemCount As Long)
Public Event BatchProcessingCompleted(ByVal batchId As String, ByVal successCount As Long, ByVal failureCount As Long)
Public Event PerformanceAlert(ByVal validationId As String, ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)
Public Event CacheUpdated(ByVal validationId As String, ByVal isCached As Boolean, ByVal cacheSize As Long)
Public Event StateChanged(ByVal validationId As String, ByVal oldState As ValidationState, ByVal newState As ValidationState)

''' <summary>
''' 非同期バリデーションを開始します
''' </summary>
''' <param name="value">検証する値</param>
''' <param name="context">バリデーションコンテキスト</param>
''' <param name="priority">実行優先度（オプション）</param>
''' <param name="options">実行オプション（オプション）</param>
''' <returns>バリデーションID（キャンセル時に使用）</returns>
''' <remarks>
''' エラー処理要件：
''' - パラメータの検証
''' - リソースの可用性確認
''' - 同時実行数の制限
''' - タイムアウト設定
''' - キューの管理
''' - メモリ使用量の監視
''' - キャッシュの活用
''' </remarks>
Public Function StartAsyncValidation(ByVal value As Variant, ByVal context As IValidationContext, Optional ByVal priority As ValidationPriority = vpNormal, Optional ByVal options As AsyncValidationOptions) As String
End Function

''' <summary>
''' バッチバリデーションを開始します
''' </summary>
''' <param name="items">検証するアイテムのコレクション</param>
''' <param name="context">バリデーションコンテキスト</param>
''' <param name="options">バッチオプション（オプション）</param>
''' <returns>バッチID</returns>
Public Function StartBatchValidation(ByVal items As Collection, ByVal context As IValidationContext, Optional ByVal options As BatchValidationOptions) As String
End Function

''' <summary>
''' 非同期バリデーションをキャンセルします
''' </summary>
''' <param name="validationId">バリデーションID</param>
''' <param name="gracefulShutdown">グレースフルシャットダウンを試みるかどうか</param>
''' <param name="options">キャンセルオプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - IDの有効性確認
''' - リソースの適切な解放
''' - 実行中の処理の安全な停止
''' - 依存タスクの処理
''' - キャッシュの更新
''' - メトリクスの記録
''' </remarks>
Public Sub CancelValidation(ByVal validationId As String, Optional ByVal gracefulShutdown As Boolean = True, Optional ByVal options As CancellationOptions)
End Sub

''' <summary>
''' バリデーションの状態を取得します
''' </summary>
''' <param name="validationId">バリデーションID</param>
''' <param name="includeDetails">詳細情報を含めるかどうか</param>
''' <returns>現在の状態情報</returns>
Public Function GetValidationStatus(ByVal validationId As String, Optional ByVal includeDetails As Boolean = False) As ValidationStatus
End Function

''' <summary>
''' タイムアウト時間を設定します
''' </summary>
''' <param name="milliseconds">タイムアウト時間（ミリ秒）</param>
''' <remarks>
''' エラー処理要件：
''' - 値の範囲チェック
''' - 実行中の処理への影響考慮
''' - タイムアウト処理の整合性確認
''' - デフォルト値の適用
''' </remarks>
Public Property Let TimeoutMilliseconds(ByVal milliseconds As Long)
End Property

''' <summary>
''' タイムアウト時間を取得します
''' </summary>
Public Property Get TimeoutMilliseconds() As Long
End Property

''' <summary>
''' 最大同時実行数を設定します
''' </summary>
''' <param name="count">最大同時実行数</param>
''' <remarks>
''' エラー処理要件：
''' - 値の範囲チェック
''' - システムリソースの考慮
''' - 実行中タスクへの影響評価
''' - スケーリングの適用
''' </remarks>
Public Property Let MaxConcurrentValidations(ByVal count As Long)
End Property

''' <summary>
''' 最大同時実行数を取得します
''' </summary>
Public Property Get MaxConcurrentValidations() As Long
End Property

''' <summary>
''' 進捗通知の間隔を設定します
''' </summary>
''' <param name="milliseconds">通知間隔（ミリ秒）</param>
''' <remarks>
''' エラー処理要件：
''' - 値の範囲チェック
''' - パフォーマンスへの影響考慮
''' - イベント発行の最適化
''' - バッチ処理の考慮
''' </remarks>
Public Property Let ProgressNotificationInterval(ByVal milliseconds As Long)
End Property

''' <summary>
''' 進捗通知の間隔を取得します
''' </summary>
Public Property Get ProgressNotificationInterval() As Long
End Property

''' <summary>
''' すべての非同期バリデーションを待機します
''' </summary>
''' <param name="timeoutMilliseconds">タイムアウト時間（ミリ秒）</param>
''' <param name="options">待機オプション（オプション）</param>
''' <returns>すべての処理が完了した場合True</returns>
''' <remarks>
''' エラー処理要件：
''' - タイムアウト処理
''' - リソースの監視
''' - デッドロック防止
''' - 部分的な完了の処理
''' - キャンセル処理
''' </remarks>
Public Function WaitForAll(Optional ByVal timeoutMilliseconds As Long = -1, Optional ByVal options As WaitOptions) As Boolean
End Function

''' <summary>
''' 実行中のバリデーションの一覧を取得します
''' </summary>
''' <param name="filter">フィルタ条件（オプション）</param>
''' <returns>実行中のバリデーションIDのコレクション</returns>
Public Function GetRunningValidations(Optional ByVal filter As ValidationFilter) As Collection
End Function

''' <summary>
''' バリデーション結果を非同期で取得します
''' </summary>
''' <param name="validationId">バリデーションID</param>
''' <param name="options">取得オプション（オプション）</param>
''' <returns>バリデーション結果（完了していない場合はNothing）</returns>
Public Function TryGetResults(ByVal validationId As String, Optional ByVal options As ResultOptions) As Variant
End Function

''' <summary>
''' バリデーションキューの状態を取得します
''' </summary>
''' <param name="includeMetrics">メトリクスを含めるかどうか</param>
''' <returns>キューの状態情報</returns>
Public Function GetQueueStatus(Optional ByVal includeMetrics As Boolean = False) As ValidationQueueStatus
End Function

''' <summary>
''' バリデーションの優先度を変更します
''' </summary>
''' <param name="validationId">バリデーションID</param>
''' <param name="newPriority">新しい優先度</param>
''' <param name="options">優先度変更オプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - IDの有効性確認
''' - 優先度の範囲チェック
''' - キューの再構築
''' - 実行中タスクへの影響考慮
''' </remarks>
Public Sub ChangePriority(ByVal validationId As String, ByVal newPriority As ValidationPriority, Optional ByVal options As PriorityOptions)
End Sub

''' <summary>
''' エラー発生時のリトライ戦略を設定します
''' </summary>
''' <param name="strategy">リトライ戦略</param>
Public Property Let RetryStrategy(ByVal strategy As IErrorStrategy)
End Property

''' <summary>
''' エラー発生時のリトライ戦略を取得します
''' </summary>
Public Property Get RetryStrategy() As IErrorStrategy
End Property

''' <summary>
''' バリデーション結果の永続化を設定します
''' </summary>
''' <param name="enabled">永続化を有効にするかどうか</param>
''' <param name="options">永続化オプション（オプション）</param>
Public Property Let PersistResults(ByVal enabled As Boolean, Optional ByVal options As PersistenceOptions)
End Property

''' <summary>
''' バリデーション結果の永続化状態を取得します
''' </summary>
Public Property Get PersistResults() As Boolean
End Property

''' <summary>
''' リソース使用量の制限を設定します
''' </summary>
''' <param name="resourceType">リソースの種類</param>
''' <param name="limit">制限値</param>
''' <param name="options">リソース制限オプション（オプション）</param>
Public Sub SetResourceLimit(ByVal resourceType As String, ByVal limit As Long, Optional ByVal options As ResourceLimitOptions)
End Sub

''' <summary>
''' リソース使用量の制限を取得します
''' </summary>
''' <param name="resourceType">リソースの種類</param>
Public Function GetResourceLimit(ByVal resourceType As String) As Long
End Function

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <returns>パフォーマンス情報のコレクション</returns>
Public Function GetPerformanceMetrics() As Collection
End Function

''' <summary>
''' キャッシュ戦略を設定します
''' </summary>
''' <param name="strategy">キャッシュ戦略</param>
Public Property Let CacheStrategy(ByVal strategy As IValidationCacheStrategy)
End Property

''' <summary>
''' スケジューリング戦略を設定します
''' </summary>
''' <param name="strategy">スケジューリング戦略</param>
Public Property Let SchedulingStrategy(ByVal strategy As IValidationSchedulingStrategy)
End Property