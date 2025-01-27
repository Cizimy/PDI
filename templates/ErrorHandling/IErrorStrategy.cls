VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IErrorStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("エラーからの回復戦略を定義し、段階的な回復処理とフォールバックメカニズムを提供するインターフェース")

' 定数定義
Private Const DEFAULT_RETRY_COUNT As Long = 3           ' デフォルトのリトライ回数
Private Const DEFAULT_RETRY_INTERVAL_MS As Long = 1000  ' デフォルトのリトライ間隔（1秒）
Private Const MAX_RETRY_COUNT As Long = 10              ' 最大リトライ回数
Private Const MAX_RECOVERY_CHAIN As Long = 5            ' 最大回復チェーン数
Private Const RECOVERY_TIMEOUT_MS As Long = 30000       ' 回復処理のタイムアウト時間（30秒）
Private Const MIN_BACKOFF_MS As Long = 100              ' 最小バックオフ時間（100ミリ秒）
Private Const MAX_BACKOFF_MS As Long = 30000           ' 最大バックオフ時間（30秒）

' フォールバック理由を定義する列挙型
Public Enum FallbackReason
    frRetryLimitExceeded = 1    ' リトライ回数超過
    frTimeoutOccurred = 2       ' タイムアウト発生
    frResourceUnavailable = 3   ' リソース利用不可
    frStrategyFailed = 4        ' 戦略実行失敗
    frSystemOverload = 5        ' システム過負荷
    frDataInconsistency = 6     ' データ不整合
    frUnexpectedError = 7       ' 予期せぬエラー
End Enum

' イベント定義
Public Event RecoveryChainStarted(ByVal chainId As String, ByVal strategies As Collection)
Public Event RecoveryChainCompleted(ByVal chainId As String, ByVal successCount As Long, ByVal failureCount As Long)
Public Event FallbackStrategyActivated(ByVal errorInfo As ErrorInfo, ByVal fallbackStrategy As IErrorStrategy, ByVal reason As FallbackReason)
Public Event StrategyExecutionStarted(ByVal strategyId As String, ByVal priority As Long)
Public Event StrategyExecutionCompleted(ByVal strategyId As String, ByVal success As Boolean)

''' <summary>
''' エラーからの回復を試みます
''' </summary>
''' <param name="errorInfo">エラー情報</param>
''' <param name="errorManager">エラーマネージャー</param>
''' <param name="context">回復コンテキスト（オプション）</param>
''' <param name="options">回復オプション（オプション）</param>
''' <returns>回復成功時True、失敗時False</returns>
''' <remarks>
''' エラー処理要件：
''' - エラーマネージャーと連携した回復処理
''' - タイムアウト処理
''' - リソースの適切な解放
''' - 回復不可能な状態の検出
''' - イベントの適切な発行
''' - メトリクスの収集
''' - パターンの分析
'''
''' 実装時の注意点：
''' 1. エラーマネージャーのコンテキストを利用すること
''' 2. 回復処理自体のエラーを適切に処理すること
''' 3. リソースリークを防ぐこと
''' 4. デッドロックを防ぐこと
''' 5. 回復の進捗を通知すること
''' 6. パフォーマンスを監視すること
''' </remarks>
Public Function RecoverFromError(ByRef errorInfo As ErrorInfo, ByVal errorManager As clsErrorManager, Optional ByVal context As RecoveryContext, Optional ByVal options As RecoveryOptions) As Boolean
End Function

''' <summary>
''' 回復チェーンを作成します
''' </summary>
''' <param name="strategies">チェーンに含める戦略のコレクション</param>
''' <param name="errorManager">エラーマネージャー</param>
''' <returns>回復戦略のコレクション</returns>
''' <remarks>
''' - 戦略は優先度順にソートされる
''' - 各戦略はエラーマネージャーのコンテキストを共有
''' - チェーンの実行状態はErrorManagerで管理
''' - 戦略間の依存関係を考慮
''' </remarks>
Public Function CreateRecoveryChain(ByVal strategies As Collection, ByVal errorManager As clsErrorManager) As RecoveryChain
End Function

''' <summary>
''' 最大リトライ回数を設定します
''' </summary>
''' <param name="value">リトライ回数</param>
''' <param name="options">リトライオプション（オプション）</param>
''' <remarks>
''' - 0以上の値を設定すること
''' - デフォルト値は3
''' - 実行中の回復処理への影響を考慮
''' - バックオフ戦略の適用
''' </remarks>
Public Property Let MaxRetryCount(ByVal value As Long, Optional ByVal options As RetryOptions)
End Property

''' <summary>
''' 最大リトライ回数を取得します
''' </summary>
Public Property Get MaxRetryCount() As Long
End Property

''' <summary>
''' 回復戦略の優先度を取得します
''' </summary>
''' <remarks>
''' - 値が大きいほど優先度が高い
''' - 複数の回復戦略がある場合、優先度の高い順に実行される
''' - 同じ優先度の場合、登録順で実行
''' - チェーン内での順序も考慮
''' </remarks>
Public Property Get Priority() As Long
End Property

''' <summary>
''' 回復戦略の優先度を設定します
''' </summary>
''' <param name="value">優先度</param>
''' <param name="options">優先度オプション（オプション）</param>
Public Property Let Priority(ByVal value As Long, Optional ByVal options As PriorityOptions)
End Property

''' <summary>
''' リトライ間隔（ミリ秒）を設定します
''' </summary>
''' <param name="value">間隔（ミリ秒）</param>
''' <param name="options">間隔オプション（オプション）</param>
''' <remarks>
''' - 0以上の値を設定すること
''' - デフォルト値は1000（1秒）
''' - 指数バックオフを考慮
''' - システム負荷を考慮
''' </remarks>
Public Property Let RetryInterval(ByVal value As Long, Optional ByVal options As IntervalOptions)
End Property

''' <summary>
''' リトライ間隔（ミリ秒）を取得します
''' </summary>
Public Property Get RetryInterval() As Long
End Property

''' <summary>
''' 回復戦略の説明を取得します
''' </summary>
''' <param name="locale">ロケール（オプション）</param>
''' <returns>戦略の説明</returns>
''' <remarks>
''' - ログ出力やデバッグ用
''' - 具体的な回復方法を説明する文字列を返すこと
''' - 多言語対応を考慮
''' - メトリクス情報を含める
''' </remarks>
Public Function GetDescription(Optional ByVal locale As String) As String
End Function

''' <summary>
''' フォールバック戦略を設定します
''' </summary>
''' <param name="strategy">フォールバック戦略</param>
''' <param name="options">フォールバックオプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - 循環参照の検出
''' - 戦略の有効性確認
''' - リソースの確保
''' - チェーンの整合性確認
''' </remarks>
Public Sub SetFallbackStrategy(ByVal strategy As IErrorStrategy, Optional ByVal options As FallbackOptions)
End Sub

''' <summary>
''' フォールバック戦略を取得します
''' </summary>
Public Property Get FallbackStrategy() As IErrorStrategy
End Property

''' <summary>
''' 回復処理の進捗状況を取得します
''' </summary>
''' <param name="includeMetrics">メトリクスを含めるかどうか</param>
Public Function GetProgress(Optional ByVal includeMetrics As Boolean = False) As RecoveryProgress
End Function

''' <summary>
''' エラー種別ごとの回復ハンドラーを設定します
''' </summary>
''' <param name="errorCode">エラーコード</param>
''' <param name="handler">エラーハンドラー</param>
''' <param name="options">ハンドラーオプション（オプション）</param>
''' <remarks>
''' - ハンドラーと戦略を紐付け
''' - 既存のハンドラーは上書きされる
''' - 優先度は自動的に計算
''' - エラーマネージャーのコンテキストを利用
''' </remarks>
Public Sub SetRecoveryHandler(ByVal errorCode As ErrorCode, _
                           ByVal handler As IErrorHandler, _
                           Optional ByVal options As StrategyOptions)
End Sub

''' <summary>
''' 回復履歴を取得します
''' </summary>
''' <param name="filter">フィルタ条件（オプション）</param>
''' <returns>回復試行の履歴コレクション</returns>
Public Function GetRecoveryHistory(Optional ByVal filter As RecoveryHistoryFilter) As Collection
End Function

''' <summary>
''' 回復戦略の状態を検証します
''' </summary>
''' <param name="validationLevel">検証レベル（オプション）</param>
''' <returns>有効な場合True</returns>
Public Function Validate(Optional ByVal validationLevel As ValidationLevel = vlNormal) As Boolean
End Function

''' <summary>
''' リソースを解放します
''' </summary>
''' <param name="options">クリーンアップオプション（オプション）</param>
Public Sub Cleanup(Optional ByVal options As CleanupOptions)
End Sub

''' <summary>
''' 回復処理をキャンセルします
''' </summary>
''' <param name="options">キャンセルオプション（オプション）</param>
Public Sub CancelRecovery(Optional ByVal options As CancellationOptions)
End Sub

''' <summary>
''' キャンセル状態を取得します
''' </summary>
Public Property Get IsCancellationRequested() As Boolean
End Property

''' <summary>
''' 非同期回復をサポートしているかどうかを取得します
''' </summary>
Public Property Get SupportsAsyncRecovery() As Boolean
End Property

''' <summary>
''' トランザクションをサポートしているかどうかを取得します
''' </summary>
Public Property Get SupportsTransactions() As Boolean
End Property

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <returns>パフォーマンス情報のコレクション</returns>
Public Function GetPerformanceMetrics() As Collection
End Function

''' <summary>
''' エラーパターンを分析します
''' </summary>
''' <param name="timeWindow">分析期間（ミリ秒）</param>
''' <returns>パターン分析結果のコレクション</returns>
Public Function AnalyzeErrorPatterns(ByVal timeWindow As Long) As Collection
End Function

''' <summary>
''' バックオフ戦略を設定します
''' </summary>
''' <param name="strategy">バックオフ戦略</param>
Public Property Let BackoffStrategy(ByVal strategy As IBackoffStrategy)
End Property