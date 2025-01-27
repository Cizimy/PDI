VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

Implements IError

' 定数定義
Private Const MODULE_NAME As String = "ErrorImpl"
Private Const MAX_ERROR_RECURSION As Long = 3
Private Const ERROR_COUNT_DICT_SIZE As Long = 100
Private Const MAX_HISTORY_SIZE As Long = 1000
Private Const MAX_BATCH_SIZE As Long = 100
Private Const CLEANUP_INTERVAL_MS As Long = 300000
Private Const MAX_RECOVERY_ATTEMPTS As Long = 3
Private Const ANALYSIS_INTERVAL_MS As Long = 3600000

' イベント定義
Public Event ErrorOccurred(ByVal errorInfo As ErrorInfo, ByVal severity As ErrorSeverity)
Public Event ErrorHandled(ByVal errorInfo As ErrorInfo, ByVal handler As IErrorHandler, ByVal success As Boolean)
Public Event ErrorAnalysisCompleted(ByVal analysisResult As ErrorAnalysisResult)
Public Event ErrorCategoryChanged(ByVal errorCode As ErrorCode, ByVal oldCategory As ErrorCategory, ByVal newCategory As ErrorCategory)
Public Event ErrorThresholdExceeded(ByVal errorCode As ErrorCode, ByVal count As Long, ByVal threshold As Long)
Public Event RecoveryAttempted(ByVal errorInfo As ErrorInfo, ByVal attempt As Long, ByVal success As Boolean)
Public Event BatchProcessed(ByVal batchId As String, ByVal successCount As Long, ByVal failureCount As Long)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)

' メンバ変数
Private mHandlerManager As ErrorHandlerManager
Private mEmergencyLogger As IEmergencyLogger
Private mErrorCounts As Object ' Scripting.Dictionary
Private mErrorHistory As Collection
Private mErrorCategories As Object ' Scripting.Dictionary
Private mLock As clsLock
Private mIsInitialized As Boolean
Private mErrorRecursionCount As Long
Private mPerformanceMonitor As clsPerformanceMonitor
Private mLastCleanupTime As Date
Private mLastAnalysisTime As Date

' 初期化・終了処理
Private Sub Class_Initialize()
    Set mHandlerManager = New ErrorHandlerManager
    Set mEmergencyLogger = New EmergencyLogger
    Set mErrorCounts = CreateObject("Scripting.Dictionary")
    Set mErrorHistory = New Collection
    Set mErrorCategories = CreateObject("Scripting.Dictionary")
    Set mLock = New clsLock
    Set mPerformanceMonitor = New clsPerformanceMonitor
    mErrorRecursionCount = 0
    mLastCleanupTime = Now
    mLastAnalysisTime = Now
    mIsInitialized = True
End Sub

Private Sub Class_Terminate()
    Set mHandlerManager = Nothing
    Set mEmergencyLogger = Nothing
    Set mErrorCounts = Nothing
    Set mErrorHistory = Nothing
    Set mErrorCategories = Nothing
    Set mLock = Nothing
    Set mPerformanceMonitor = Nothing
    mIsInitialized = False
End Sub

' IError 実装
Private Function IError_HandleError(ByRef errorInfo As ErrorInfo, Optional ByVal options As HandlingOptions) As Boolean
    If Not mIsInitialized Then Exit Function
    
    ' パフォーマンス計測開始
    mPerformanceMonitor.StartMeasurement "HandleError"
    
    ' エラーの再帰制御
    mErrorRecursionCount = mErrorRecursionCount + 1
    If mErrorRecursionCount > MAX_ERROR_RECURSION Then
        mEmergencyLogger.LogEmergencyError "エラー処理の再帰回数が上限を超えました。処理を中断します。", errorInfo
        IError_HandleError = False
        GoTo Cleanup
    End If
    
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    ' エラー情報の補完と履歴の追加
    CompleteErrorInfo errorInfo
    AddToHistory errorInfo
    
    ' エラー発生回数をカウント
    IError_IncrementErrorCount errorInfo.Code
    
    ' エラーの分析と分類
    AnalyzeError errorInfo
    CategorizeError errorInfo
    
    ' イベント通知
    RaiseEvent ErrorOccurred(errorInfo, GetErrorSeverity(errorInfo))
    
    ' エラーハンドラーの取得と実行
    Dim handler As IErrorHandler
    Set handler = mHandlerManager.GetErrorHandler(errorInfo.Code)
    
    If Not handler Is Nothing Then
        IError_HandleError = handler.HandleError(errorInfo)
        RaiseEvent ErrorHandled(errorInfo, handler, IError_HandleError)
    End If
    
    ' 定期的なクリーンアップとメトリクス収集
    If DateDiff("s", mLastCleanupTime, Now) > CLEANUP_INTERVAL_MS / 1000 Then
        PerformCleanup
    End If
    
    mLock.ReleaseLock
    GoTo Cleanup
    
ErrorHandler:
    mLock.ReleaseLock
    mEmergencyLogger.LogEmergencyError "HandleError中にエラーが発生しました: " & Err.Description, errorInfo
    IError_HandleError = False
    
Cleanup:
    mErrorRecursionCount = mErrorRecursionCount - 1
    mPerformanceMonitor.EndMeasurement "HandleError"
End Function

Private Function IError_HandleBatchErrors(ByVal errors As Collection, Optional ByVal options As BatchOptions) As BatchResult
    Dim result As New BatchResult
    Dim batchId As String
    batchId = "BATCH_" & Format(Now, "yyyymmddhhnnss")
    
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    Dim errorInfo As ErrorInfo
    For Each errorInfo In errors
        If IError_HandleError(errorInfo) Then
            result.SuccessCount = result.SuccessCount + 1
        Else
            result.FailureCount = result.FailureCount + 1
        End If
    Next
    
    RaiseEvent BatchProcessed(batchId, result.SuccessCount, result.FailureCount)
    mLock.ReleaseLock
    Set IError_HandleBatchErrors = result
    Exit Function
    
ErrorHandler:
    mLock.ReleaseLock
    mEmergencyLogger.LogEmergencyError "HandleBatchErrors中にエラーが発生しました: " & Err.Description, Nothing
    Set IError_HandleBatchErrors = result
End Function

Private Function IError_AnalyzeErrors(Optional ByVal options As AnalysisOptions) As ErrorAnalysisResult
    Dim result As New ErrorAnalysisResult
    
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    ' エラー履歴の分析
    AnalyzeErrorHistory result
    
    ' エラーパターンの検出
    DetectErrorPatterns result
    
    ' エラー傾向の分析
    AnalyzeErrorTrends result
    
    RaiseEvent ErrorAnalysisCompleted(result)
    mLock.ReleaseLock
    Set IError_AnalyzeErrors = result
    Exit Function
    
ErrorHandler:
    mLock.ReleaseLock
    mEmergencyLogger.LogEmergencyError "AnalyzeErrors中にエラーが発生しました: " & Err.Description, Nothing
    Set IError_AnalyzeErrors = result
End Function

Private Function IError_GetErrorHistory(Optional ByVal options As HistoryOptions) As Collection
    Set IError_GetErrorHistory = mErrorHistory
End Function

Private Function IError_GetPerformanceMetrics() As Collection
    Dim metrics As New Collection
    
    With metrics
        .Add mPerformanceMonitor.GetAllMeasurements()
        .Add GetErrorCountMetrics()
        .Add GetHandlerPerformanceMetrics()
        .Add GetMemoryUsageMetrics()
    End With
    
    Set IError_GetPerformanceMetrics = metrics
End Function

' プライベートメソッド
Private Sub CompleteErrorInfo(ByRef errorInfo As ErrorInfo)
    With errorInfo
        If .OccurredAt = #12:00:00 AM# Then .OccurredAt = Now
        If .Category = 0 Then .Category = modErrorCodes.GetErrorCategory(.Code)
        If Len(.StackTrace) = 0 Then .StackTrace = modStackTrace.GetStackTrace()
    End With
End Sub

Private Sub AddToHistory(ByRef errorInfo As ErrorInfo)
    mErrorHistory.Add errorInfo
    If mErrorHistory.Count > MAX_HISTORY_SIZE Then
        RemoveOldestHistoryEntry
    End If
End Sub

Private Sub RemoveOldestHistoryEntry()
    If mErrorHistory.Count > 0 Then
        mErrorHistory.Remove 1
    End If
End Sub

Private Sub PerformCleanup()
    mLastCleanupTime = Now
    CleanupErrorCounts
    CleanupErrorHistory
    CollectGarbage
End Sub

Private Sub CleanupErrorCounts()
    While mErrorCounts.Count > ERROR_COUNT_DICT_SIZE
        RemoveOldestErrorCount
    Wend
End Sub

Private Sub CleanupErrorHistory()
    While mErrorHistory.Count > MAX_HISTORY_SIZE
        RemoveOldestHistoryEntry
    Wend
End Sub

Private Function GetErrorSeverity(ByRef errorInfo As ErrorInfo) As ErrorSeverity
    ' エラーの重要度を判定するロジック
    ' 実装は省略
End Function

Private Sub AnalyzeError(ByRef errorInfo As ErrorInfo)
    ' エラーの分析ロジック
    ' 実装は省略
End Sub

Private Sub CategorizeError(ByRef errorInfo As ErrorInfo)
    ' エラーの分類ロジック
    ' 実装は省略
End Sub

Private Sub AnalyzeErrorHistory(ByRef result As ErrorAnalysisResult)
    ' エラー履歴の分析ロジック
    ' 実装は省略
End Sub

Private Sub DetectErrorPatterns(ByRef result As ErrorAnalysisResult)
    ' エラーパターンの検出ロジック
    ' 実装は省略
End Sub

Private Sub AnalyzeErrorTrends(ByRef result As ErrorAnalysisResult)
    ' エラー傾向の分析ロジック
    ' 実装は省略
End Sub

Private Function GetErrorCountMetrics() As Collection
    ' エラーカウントのメトリクス取得ロジック
    ' 実装は省略
End Function

Private Function GetHandlerPerformanceMetrics() As Collection
    ' ハンドラーのパフォーマンスメトリクス取得ロジック
    ' 実装は省略
End Function

Private Function GetMemoryUsageMetrics() As Collection
    ' メモリ使用量のメトリクス取得ロジック
    ' 実装は省略
End Function

' テストサポート（開発環境用）
#If DEBUG Then
    Public Function GetErrorCount(ByVal errorCode As ErrorCode) As Long
        mLock.AcquireLock
        Dim key As String
        key = CStr(errorCode)
        GetErrorCount = 0
        If mErrorCounts.Exists(key) Then
            GetErrorCount = mErrorCounts(key)
        End If
        mLock.ReleaseLock
    End Function
    
    Public Sub ClearErrorCounts()
        mLock.AcquireLock
        mErrorCounts.RemoveAll
        mLock.ReleaseLock
    End Sub
    
    Public Function GetPerformanceReport() As String
        GetPerformanceReport = mPerformanceMonitor.GetAllMeasurements()
    End Function
#End If