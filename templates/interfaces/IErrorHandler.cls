VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

'@Description("エラー処理とリカバリを提供するインターフェース")

' 定数定義
Private Const MAX_RETRY_COUNT As Long = 3           ' 最大リトライ回数
Private Const RETRY_DELAY_MS As Long = 1000         ' リトライ間隔（1秒）
Private Const MAX_ERROR_QUEUE As Long = 1000        ' エラーキューの最大サイズ
Private Const DEFAULT_TIMEOUT_MS As Long = 30000    ' デフォルトタイムアウト（30秒）
Private Const MAX_RECOVERY_ATTEMPTS As Long = 5     ' 最大回復試行回数
Private Const CLEANUP_INTERVAL_MS As Long = 300000  ' クリーンアップ間隔（5分）

' イベント定義
Public Event ErrorHandlingStarted(ByVal errorInfo As ErrorInfo, ByVal priority As ErrorPriority)
Public Event ErrorHandlingCompleted(ByVal errorInfo As ErrorInfo, ByVal success As Boolean)
Public Event RecoveryAttempted(ByVal errorInfo As ErrorInfo, ByVal attempt As Long, ByVal success As Boolean)
Public Event RetryAttempted(ByVal errorInfo As ErrorInfo, ByVal attempt As Long)
Public Event ResourceReleased(ByVal resourceType As String, ByVal success As Boolean)
Public Event ErrorQueued(ByVal errorInfo As ErrorInfo, ByVal queuePosition As Long)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)

''' <summary>
''' エラーを処理します
''' </summary>
''' <param name="errorInfo">エラー詳細情報</param>
''' <param name="options">処理オプション（オプション）</param>
''' <returns>処理を継続する場合はTrue、中断する場合はFalse</returns>
''' <remarks>
''' エラー処理要件：
''' - エラー情報の検証
''' - 優先度の判定
''' - リカバリ処理
''' - リソース解放
''' - ログ記録
''' - メトリクス収集
''' </remarks>
Public Function HandleError(ByRef errorInfo As ErrorInfo, _
                          Optional ByVal options As HandlingOptions) As Boolean
End Function

''' <summary>
''' エラーからの回復を試みます
''' </summary>
''' <param name="errorInfo">エラー情報</param>
''' <param name="options">回復オプション（オプション）</param>
''' <returns>回復成功時True</returns>
Public Function AttemptRecovery(ByRef errorInfo As ErrorInfo, _
                              Optional ByVal options As RecoveryOptions) As Boolean
End Function

''' <summary>
''' リトライ処理を実行します
''' </summary>
''' <param name="operation">実行する操作</param>
''' <param name="options">リトライオプション（オプション）</param>
''' <returns>実行結果</returns>
Public Function RetryOperation(ByVal operation As IRetryableOperation, _
                             Optional ByVal options As RetryOptions) As OperationResult
End Function

''' <summary>
''' エラーの優先度を設定します
''' </summary>
''' <param name="errorType">エラータイプ</param>
''' <param name="priority">優先度</param>
Public Sub SetErrorPriority(ByVal errorType As String, ByVal priority As ErrorPriority)
End Sub

''' <summary>
''' エラー処理戦略を設定します
''' </summary>
''' <param name="errorType">エラータイプ</param>
''' <param name="strategy">処理戦略</param>
Public Sub SetHandlingStrategy(ByVal errorType As String, ByVal strategy As IErrorStrategy)
End Sub

''' <summary>
''' リソースを解放します
''' </summary>
''' <param name="resource">解放するリソース</param>
''' <param name="options">解放オプション（オプション）</param>
''' <returns>解放成功時True</returns>
Public Function ReleaseResource(ByVal resource As Object, _
                              Optional ByVal options As ReleaseOptions) As Boolean
End Function

''' <summary>
''' エラーをキューに追加します
''' </summary>
''' <param name="errorInfo">エラー情報</param>
''' <param name="options">キューオプション（オプション）</param>
Public Sub QueueError(ByRef errorInfo As ErrorInfo, Optional ByVal options As QueueOptions)
End Sub

''' <summary>
''' キューのエラーを処理します
''' </summary>
''' <param name="options">処理オプション（オプション）</param>
Public Sub ProcessErrorQueue(Optional ByVal options As ProcessOptions)
End Sub

''' <summary>
''' エラー処理の統計を取得します
''' </summary>
''' <returns>統計情報</returns>
Public Function GetHandlingStatistics() As HandlingStatistics
End Function

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <returns>パフォーマンス情報のコレクション</returns>
Public Function GetPerformanceMetrics() As Collection
End Function

''' <summary>
''' エラーハンドラーの状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function

''' <summary>
''' リソースを解放します
''' </summary>
Public Sub Cleanup()
End Sub