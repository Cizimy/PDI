VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IErrorNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("エラー通知の配信と管理を担当し、複数の通知チャネルとポリシーをサポートするインターフェース")

' 定数定義
Private Const MAX_NOTIFICATION_TARGETS As Long = 10     ' 最大通知先数
Private Const MAX_BATCH_SIZE As Long = 100              ' 最大バッチサイズ
Private Const CACHE_DURATION_MS As Long = 300000        ' キャッシュ有効期間（5分）
Private Const MIN_NOTIFICATION_INTERVAL_MS As Long = 100 ' 最小通知間隔（100ミリ秒）
Private Const MAX_RETRY_COUNT As Long = 3               ' 最大リトライ回数
Private Const MAX_TEMPLATE_SIZE As Long = 10240         ' 最大テンプレートサイズ（10KB）
Private Const DEFAULT_TIMEOUT_MS As Long = 5000         ' デフォルトタイムアウト時間（5秒）

' イベント定義
Public Event NotificationSent(ByVal errorInfo As ErrorInfo, ByVal level As ErrorNotificationLevel, ByVal metrics As NotificationMetrics)
Public Event NotificationDelivered(ByVal errorInfo As ErrorInfo, ByVal target As INotificationTarget, ByVal deliveryMetrics As DeliveryMetrics)
Public Event NotificationFailed(ByVal errorInfo As ErrorInfo, ByVal target As INotificationTarget, ByVal reason As String, ByVal retryCount As Long)
Public Event NotificationAcknowledged(ByVal errorInfo As ErrorInfo, ByVal target As INotificationTarget, ByVal acknowledgeTime As Date)
Public Event NotificationBatchProcessed(ByVal batchId As String, ByVal successCount As Long, ByVal failureCount As Long)
Public Event NotificationThrottled(ByVal errorInfo As ErrorInfo, ByVal target As INotificationTarget, ByVal nextAttemptTime As Date)
Public Event TemplateUpdated(ByVal templateName As String, ByVal version As Long)
Public Event DeliveryPolicyViolation(ByVal policyName As String, ByVal violationType As String)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)
Public Event ResourceExhausted(ByVal resourceType As String, ByVal currentUsage As Long)

''' <summary>
''' エラー通知レベルを設定します
''' </summary>
''' <param name="level">通知レベル</param>
''' <param name="options">レベルオプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - 無効な通知レベルの検証
''' - 現在の通知処理への影響考慮
''' - 設定変更のログ記録
''' - イベントの適切な発行
''' - 依存システムへの通知
''' </remarks>
Public Property Let NotificationLevel(ByVal level As ErrorNotificationLevel, Optional ByVal options As NotificationLevelOptions)
End Property

''' <summary>
''' 現在のエラー通知レベルを取得します
''' </summary>
Public Property Get NotificationLevel() As ErrorNotificationLevel
End Property

''' <summary>
''' エラーを通知します
''' </summary>
''' <param name="errorInfo">エラー情報</param>
''' <param name="options">通知オプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - 通知レベルに応じた処理
''' - 通知失敗時の代替処理
''' - リソースの適切な解放
''' - 通知の重複防止
''' - イベントの適切な発行
''' - メトリクスの収集
''' - パフォーマンスの監視
''' 
''' 実装時の注意点：
''' 1. 通知処理自体のエラーを適切に処理すること
''' 2. パフォーマンスに影響を与えないよう注意すること
''' 3. UI スレッドをブロックしないこと
''' 4. 機密情報の適切な処理を行うこと
''' 5. 通知の優先順位を考慮すること
''' 6. スロットリングを適用すること
''' </remarks>
Public Sub NotifyError(ByRef errorInfo As ErrorInfo, Optional ByVal options As NotificationOptions)
End Sub

''' <summary>
''' バッチ通知を実行します
''' </summary>
''' <param name="errors">エラー情報のコレクション</param>
''' <param name="options">バッチオプション（オプション）</param>
''' <returns>バッチ処理ID</returns>
Public Function NotifyBatch(ByVal errors As Collection, Optional ByVal options As BatchNotificationOptions) As String
End Function

''' <summary>
''' 通知先を追加します
''' </summary>
''' <param name="target">通知先オブジェクト</param>
''' <param name="priority">優先度（オプション）</param>
''' <param name="options">追加オプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - 通知先の有効性確認
''' - 重複登録の防止
''' - リソース制限の考慮
''' - 優先度の検証
''' - 依存関係の確認
''' </remarks>
Public Sub AddNotificationTarget(ByVal target As INotificationTarget, Optional ByVal priority As ValidationPriority = vpNormal, Optional ByVal options As TargetOptions)
End Sub

''' <summary>
''' 通知先を削除します
''' </summary>
''' <param name="target">通知先オブジェクト</param>
''' <param name="options">削除オプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - 存在しない通知先の処理
''' - 削除中の通知への影響考慮
''' - リソースの適切な解放
''' - イベントの適切な発行
''' - 依存システムの更新
''' </remarks>
Public Sub RemoveNotificationTarget(ByVal target As INotificationTarget, Optional ByVal options As TargetRemovalOptions)
End Sub

''' <summary>
''' すべての通知先をクリアします
''' </summary>
''' <param name="options">クリアオプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - 進行中の通知への影響考慮
''' - リソースの適切な解放
''' - 状態の整合性維持
''' - イベントの適切な発行
''' - キャッシュの更新
''' </remarks>
Public Sub ClearNotificationTargets(Optional ByVal options As ClearOptions)
End Sub

''' <summary>
''' 通知のフィルタリング条件を設定します
''' </summary>
''' <param name="filterExpression">フィルタリング条件の式</param>
''' <param name="options">フィルタオプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - 式の構文検証
''' - 無効な条件の検出
''' - パフォーマンスへの影響考慮
''' - 既存フィルタの更新
''' - 条件の最適化
''' </remarks>
Public Sub SetNotificationFilter(ByVal filterExpression As String, Optional ByVal options As FilterOptions)
End Sub

''' <summary>
''' 通知履歴を取得します
''' </summary>
''' <param name="options">取得オプション（オプション）</param>
''' <returns>通知履歴のコレクション</returns>
''' <remarks>
''' エラー処理要件：
''' - メモリ使用量の考慮
''' - 大量データの効率的な処理
''' - 履歴の整合性確認
''' - フィルタリングの適用
''' - キャッシュの活用
''' </remarks>
Public Function GetNotificationHistory(Optional ByVal options As NotificationHistoryOptions) As Collection
End Function

''' <summary>
''' 通知テンプレートを設定します
''' </summary>
''' <param name="templateName">テンプレート名</param>
''' <param name="template">テンプレート内容</param>
''' <param name="options">テンプレートオプション（オプション）</param>
Public Sub SetNotificationTemplate(ByVal templateName As String, ByVal template As String, Optional ByVal options As TemplateOptions)
End Sub

''' <summary>
''' 通知テンプレートを取得します
''' </summary>
''' <param name="templateName">テンプレート名</param>
''' <param name="version">バージョン（オプション）</param>
Public Function GetNotificationTemplate(ByVal templateName As String, Optional ByVal version As Long = -1) As String
End Function

''' <summary>
''' 通知のグループ化ルールを設定します
''' </summary>
''' <param name="groupingRule">グループ化ルール</param>
''' <param name="options">グループ化オプション（オプション）</param>
Public Sub SetGroupingRule(ByVal groupingRule As INotificationGroupingRule, Optional ByVal options As GroupingOptions)
End Sub

''' <summary>
''' 通知の重複排除ルールを設定します
''' </summary>
''' <param name="deduplicationRule">重複排除ルール</param>
''' <param name="options">重複排除オプション（オプション）</param>
Public Sub SetDeduplicationRule(ByVal deduplicationRule As INotificationDeduplicationRule, Optional ByVal options As DeduplicationOptions)
End Sub

''' <summary>
''' 通知の配信ポリシーを設定します
''' </summary>
''' <param name="policy">配信ポリシー</param>
''' <param name="options">ポリシーオプション（オプション）</param>
Public Sub SetDeliveryPolicy(ByVal policy As INotificationDeliveryPolicy, Optional ByVal options As PolicyOptions)
End Sub

''' <summary>
''' 通知のバッチ処理を有効/無効にします
''' </summary>
Public Property Let BatchProcessingEnabled(ByVal value As Boolean)
End Property

''' <summary>
''' 通知のバッチ処理が有効かどうかを取得します
''' </summary>
Public Property Get BatchProcessingEnabled() As Boolean
End Property

''' <summary>
''' 通知の進捗状況を取得します
''' </summary>
''' <param name="includeMetrics">メトリクスを含めるかどうか</param>
Public Function GetProgress(Optional ByVal includeMetrics As Boolean = False) As NotificationProgress
End Function

''' <summary>
''' キャンセル状態を設定/取得します
''' </summary>
Public Property Get IsCancellationRequested() As Boolean
End Property

Public Property Let IsCancellationRequested(ByVal value As Boolean)
End Property

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <returns>パフォーマンス情報のコレクション</returns>
Public Function GetPerformanceMetrics() As Collection
End Function

''' <summary>
''' 通知の配信状態を分析します
''' </summary>
''' <param name="timeWindow">分析期間（ミリ秒）</param>
''' <returns>配信分析結果のコレクション</returns>
Public Function AnalyzeDeliveryPatterns(ByVal timeWindow As Long) As Collection
End Function

''' <summary>
''' スロットリング戦略を設定します
''' </summary>
''' <param name="strategy">スロットリング戦略</param>
Public Property Let ThrottlingStrategy(ByVal strategy As INotificationThrottlingStrategy)
End Property

''' <summary>
''' スケジューリング戦略を設定します
''' </summary>
''' <param name="strategy">スケジューリング戦略</param>
Public Property Let SchedulingStrategy(ByVal strategy As INotificationSchedulingStrategy)
End Property