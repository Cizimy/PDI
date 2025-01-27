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
Private Const DEFAULT_MAX_RETRY_COUNT As Long = 3    ' デフォルトの最大リトライ回数
Private Const DEFAULT_RETRY_INTERVAL As Long = 1000  ' デフォルトのリトライ間隔（1秒）
Private Const MIN_RETRY_INTERVAL As Long = 100       ' 最小リトライ間隔（100ミリ秒）
Private Const MAX_RETRY_INTERVAL As Long = 30000     ' 最大リトライ間隔（30秒）
Private Const DEFAULT_TIMEOUT_MS As Long = 30000     ' デフォルトタイムアウト（30秒）
Private Const MAX_RECOVERY_ATTEMPTS As Long = 5      ' 最大回復試行回数
Private Const CLEANUP_INTERVAL_MS As Long = 300000   ' クリーンアップ間隔（5分）

' リトライ設定用の列挙型
Public Enum RetryBehavior
    rbLinear = 0        ' 線形バックオフ
    rbExponential = 1   ' 指数バックオフ
    rbFibonacci = 2     ' フィボナッチバックオフ
    rbCustom = 3        ' カスタムバックオフ
End Enum

' イベント定義
Public Event ErrorHandlingStarted(ByVal errorInfo As ErrorInfo, ByVal priority As ErrorPriority)
Public Event ErrorHandlingCompleted(ByVal errorInfo As ErrorInfo, ByVal success As Boolean)
Public Event RetryAttempted(ByVal errorInfo As ErrorInfo, ByVal attempt As Long, ByVal success As Boolean)
Public Event ResourceReleased(ByVal resourceType As String, ByVal success As Boolean)

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
''' リトライ処理を実行します
''' </summary>
''' <param name="operation">実行する操作</param>
''' <param name="options">リトライオプション（オプション）</param>
''' <returns>実行結果</returns>
''' <remarks>
''' - リトライ回数と間隔は構成済みの設定を使用
''' - 指数バックオフを適用
''' - リトライ結果は統計情報に自動的に記録
''' </remarks>
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
''' リトライ設定を構成します
''' </summary>
''' <param name="maxRetryCount">最大リトライ回数</param>
''' <param name="retryInterval">リトライ間隔（ミリ秒）</param>
''' <param name="behavior">リトライの挙動</param>
''' <param name="options">リトライオプション（オプション）</param>
Public Sub ConfigureRetry(ByVal maxRetryCount As Long, ByVal retryInterval As Long, _
                        Optional ByVal behavior As RetryBehavior = rbLinear, _
                        Optional ByVal options As RetryOptions)
End Sub

''' <summary>
''' 現在のリトライ設定を取得します
''' </summary>
''' <returns>リトライ設定のディクショナリ</returns>
Public Function GetRetryConfiguration() As Dictionary
End Function

''' <summary>
''' リトライ間隔を計算します
''' </summary>
''' <param name="attempt">現在の試行回数</param>
''' <returns>次のリトライまでの待機時間（ミリ秒）</returns>
Public Function CalculateRetryInterval(ByVal attempt As Long) As Long
End Function

''' <summary>
''' カスタムバックオフ戦略を設定します
''' </summary>
''' <param name="strategy">バックオフ戦略を計算する関数</param>
Public Sub SetCustomBackoffStrategy(ByVal strategy As IBackoffStrategy)
End Sub

''' <summary>
''' エラー処理の統計を取得します
''' </summary>
''' <param name="errorManager">統計情報を提供するErrorManager</param>
''' <returns>統計情報</returns>
''' <remarks>
''' - 統計情報はclsErrorManagerのErrorStatisticsオブジェクトから取得
''' - ハンドラー固有の統計のみを返す
''' - グローバルな統計はErrorManagerから直接取得すること
''' </remarks>
Public Function GetHandlingStatistics(ByVal errorManager As clsErrorManager) As HandlingStatistics
End Function

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <param name="errorManager">メトリクスを提供するErrorManager</param>
''' <returns>パフォーマンス情報のコレクション</returns>
''' <remarks>
''' - メトリクスはclsErrorManagerのErrorStatisticsオブジェクトから取得
''' - ハンドラー固有のメトリクスのみを返す
''' - システム全体のメトリクスはErrorManagerから直接取得すること
''' </remarks>
Public Function GetPerformanceMetrics(ByVal errorManager As clsErrorManager) As Collection
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