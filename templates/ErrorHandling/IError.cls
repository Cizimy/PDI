VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("エラー処理と分析の基本機能を提供するインターフェース")

' 定数定義
Private Const MAX_ERROR_HISTORY As Long = 1000     ' エラー履歴の最大数
Private Const MAX_HANDLERS As Long = 100           ' ハンドラーの最大数
Private Const DEFAULT_THRESHOLD As Long = 10       ' デフォルトの閾値
Private Const CACHE_DURATION_MS As Long = 300000   ' キャッシュの有効期間（5分）
Private Const MAX_RETRY_COUNT As Long = 3          ' 最大リトライ回数
Private Const CLEANUP_INTERVAL_MS As Long = 3600000 ' クリーンアップ間隔（1時間）
Private Const MAX_ERROR_MESSAGE_LENGTH As Long = 8192 ' エラーメッセージの最大長

' イベント定義
Public Event ErrorOccurred(ByVal errorInfo As ErrorInfo, ByVal severity As ErrorSeverity)
Public Event ErrorHandled(ByVal errorInfo As ErrorInfo, ByVal handler As IErrorHandler, ByVal success As Boolean)
Public Event ThresholdExceeded(ByVal errorCode As ErrorCode, ByVal count As Long, ByVal threshold As Long)
Public Event HandlerRegistered(ByVal errorCode As ErrorCode, ByVal handler As IErrorHandler)
Public Event HandlerUnregistered(ByVal errorCode As ErrorCode)
Public Event ErrorAnalysisCompleted(ByVal analysisResult As ErrorAnalysisResult)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)

'@Description("エラー処理オプションを定義する列挙型")
Public Enum ErrorHandlingOptions
    ehoNone = 0             ' オプションなし
    ehoRetry = 1            ' エラー発生時にリトライを試みる
    ehoLogOnly = 2          ' エラーをログに記録するのみで処理は続行
    ehoNotifyUser = 4       ' エラーをユーザーに通知する
    ehoRollback = 8         ' エラー発生時にトランザクションをロールバック
    ehoIgnore = 16          ' エラーを無視する（非推奨）
    ehoEmergencyLog = 32    ' エマージェンシーロガーに記録
End Enum

'@Description("エラーハンドラー登録オプションを定義する列挙型")
Public Enum HandlerOptions
    hoNone = 0             ' オプションなし
    hoReplace = 1          ' 既存のハンドラーを置き換える
    hoAppend = 2           ' 既存のハンドラーリストに追加する
    hoHighPriority = 4     ' 高優先度のハンドラーとして登録
    hoLowPriority = 8      ' 低優先度のハンドラーとして登録
End Enum

'@Description("エラーハンドラー登録解除オプションを定義する列挙型")
Public Enum UnregisterOptions
    uoNone = 0             ' オプションなし
    uoSpecificHandler = 1   ' 特定のハンドラーのみを解除
    uoAllHandlers = 2       ' エラーコードに関連付けられたすべてのハンドラーを解除
End Enum

'@Description("エラーカウントオプションを定義する列挙型")
Public Enum CountOptions
    coNone = 0             ' オプションなし
    coIncrement = 1        ' エラーカウントをインクリメント
    coReset = 2            ' エラーカウントをリセット
    coDeduct = 4           ' エラーカウントをデクリメント
End Enum

'@Description("エラー閾値チェックオプションを定義する列挙型")
Public Enum ThresholdOptions
    toNone = 0             ' オプションなし
    toCheckGlobal = 1      ' グローバル閾値をチェック
    toCheckSpecific = 2    ' 特定のエラーコードの閾値をチェック
    toResetAfterCheck = 4  ' チェック後に閾値カウントをリセット
End Enum

'@Description("エラー履歴取得オプションを定義する列挙型")
Public Enum HistoryOptions
    hoNone = 0             ' オプションなし
    hoLimit = 1            ' 取得する履歴の数を制限
    hoSortAscending = 2    ' 昇順でソート
    hoSortDescending = 4   ' 降順でソート
    hoFilterByCode = 8     ' エラーコードでフィルタリング
    hoFilterBySeverity = 16 ' 重大度でフィルタリング
    hoFilterByTime = 32    ' 発生時刻でフィルタリング
End Enum

'@Description("エラー分析オプションを定義する列挙型")
Public Enum AnalysisOptions
    aoNone = 0              ' オプションなし
    aoBasicAnalysis = 1     ' 基本的な分析を実行
    aoDetailedAnalysis = 2  ' 詳細な分析を実行
    aoRootCause = 4         ' 根本原因分析を含む
    aoImpactAnalysis = 8    ' 影響範囲分析を含む
    aoRecommendations = 16  ' 推奨対策を含む
End Enum

'@Description("エラー分類オプションを定義する列挙型")
Public Enum CategoryOptions
    coNone = 0             ' オプションなし
    coManual = 1           ' 手動で分類
    coAutomatic = 2        ' 自動で分類
    coUseDefault = 4       ' デフォルトの分類を使用
End Enum

'@Description("エラー統計情報取得オプションを定義する列挙型")
Public Enum StatisticsOptions
    soNone = 0             ' オプションなし
    soByErrorCode = 1      ' エラーコード別の統計
    soBySeverity = 2       ' 重大度別の統計
    soByTimePeriod = 4     ' 期間別の統計
    soIncludeDetails = 8   ' 詳細情報を含む
End Enum

''' <summary>
''' エラーを処理します
''' </summary>
''' <param name="errorInfo">エラー詳細情報</param>
''' <param name="options">処理オプション（オプション）</param>
''' <returns>処理を継続する場合はTrue、中断する場合はFalse</returns>
''' <remarks>
''' エラー処理要件：
''' - エラー情報の検証
''' - ハンドラーの選択
''' - リトライ処理
''' - ログ記録
''' - メトリクス収集
''' </remarks>
Public Function HandleError(ByRef errorInfo As ErrorInfo, _
                          Optional ByVal options As ErrorHandlingOptions = ehoNone) As Boolean
End Function

''' <summary>
''' エラーコードに対応するエラーハンドラーを登録します
''' </summary>
''' <param name="errorCode">エラーコード</param>
''' <param name="handler">エラーハンドラー</param>
''' <param name="options">登録オプション（オプション）</param>
Public Sub RegisterHandler(ByVal errorCode As ErrorCode, _
                         ByVal handler As IErrorHandler, _
                         Optional ByVal options As HandlerOptions = hoNone)
End Sub

''' <summary>
''' エラーコードに対応するエラーハンドラーの登録を解除します
''' </summary>
''' <param name="errorCode">エラーコード</param>
''' <param name="options">解除オプション（オプション）</param>
Public Sub UnregisterHandler(ByVal errorCode As ErrorCode, _
                           Optional ByVal options As UnregisterOptions = uoNone)
End Sub

''' <summary>
''' エラーコードごとの発生回数をカウントアップします
''' </summary>
''' <param name="errorCode">エラーコード</param>
''' <param name="options">カウントオプション（オプション）</param>
Public Sub IncrementErrorCount(ByVal errorCode As ErrorCode, _
                             Optional ByVal options As CountOptions = coNone)
End Sub

''' <summary>
''' エラーコードの発生回数が閾値を超えているかチェックします
''' </summary>
''' <param name="errorCode">エラーコード</param>
''' <param name="threshold">閾値</param>
''' <param name="options">チェックオプション（オプション）</param>
''' <returns>閾値を超えている場合はTrue</returns>
Public Function CheckErrorThreshold(ByVal errorCode As ErrorCode, _
                                  ByVal threshold As Long, _
                                  Optional ByVal options As ThresholdOptions = toNone) As Boolean
End Function

''' <summary>
''' エラー履歴を取得します
''' </summary>
''' <param name="options">履歴オプション（オプション）</param>
''' <returns>エラー履歴のコレクション</returns>
Public Function GetErrorHistory(Optional ByVal options As HistoryOptions = hoNone) As Collection
End Function

''' <summary>
''' エラーを分析します
''' </summary>
''' <param name="options">分析オプション（オプション）</param>
''' <returns>分析結果</returns>
Public Function AnalyzeErrors(Optional ByVal options As AnalysisOptions = aoNone) As ErrorAnalysisResult
End Function

''' <summary>
''' エラーを分類します
''' </summary>
''' <param name="errorInfo">エラー情報</param>
''' <param name="options">分類オプション（オプション）</param>
''' <returns>エラー分類結果</returns>
Public Function CategorizeError(ByRef errorInfo As ErrorInfo, _
                              Optional ByVal options As CategoryOptions = coNone) As ErrorCategory
End Function

''' <summary>
''' エラー処理の統計情報を取得します
''' </summary>
''' <param name="options">統計オプション（オプション）</param>
''' <returns>統計情報</returns>
Public Function GetErrorStatistics(Optional ByVal options As StatisticsOptions = soNone) As ErrorStatistics
End Function

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <returns>パフォーマンス情報のコレクション</returns>
Public Function GetPerformanceMetrics() As Collection
End Function

''' <summary>
''' エラー処理の状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function

''' <summary>
''' リソースを解放します
''' </summary>
Public Sub Cleanup()
End Sub