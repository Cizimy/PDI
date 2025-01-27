VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsErrorManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

Implements IError

' プライベート変数
Private mErrorHandlers As Object ' Dictionary of Collections
Private mErrorHistory As Collection
Private mErrorStatistics As ErrorStatistics
Private mLastError As ErrorInfo
Private mIsInitialized As Boolean

' 定数
Private Const MAX_RETRY_ATTEMPTS As Long = 3
Private Const RETRY_DELAY_MS As Long = 1000

Private Sub Class_Initialize()
    Set mErrorHandlers = CreateObject("Scripting.Dictionary")
    Set mErrorHistory = New Collection
    Set mErrorStatistics = New ErrorStatistics
    mIsInitialized = True
End Sub

Private Sub Class_Terminate()
    Set mErrorHandlers = Nothing
    Set mErrorHistory = Nothing
    Set mErrorStatistics = Nothing
    Set mLastError = Nothing
End Sub

Private Function IError_HandleError(ByRef errorInfo As ErrorInfo, Optional ByVal options As ErrorHandlingOptions = ehoNone) As Boolean
    On Error GoTo ErrorHandler
    
    ' エラー情報の検証
    If errorInfo Is Nothing Then
        Debug.Print "Error: errorInfo is Nothing"
        Exit Function
    End If
    
    ' エラー情報の記録
    Set mLastError = errorInfo
    mErrorHistory.Add errorInfo
    mErrorStatistics.UpdateStatistics errorInfo
    
    ' イベントの発行
    RaiseEvent ErrorOccurred(errorInfo, errorInfo.Severity)
    
    ' オプションに基づいたログ記録
    If (options And ehoLogOnly) = ehoLogOnly Then
        LogError errorInfo
    End If
    
    ' エマージェンシーログへの記録
    If (options And ehoEmergencyLog) = ehoEmergencyLog Then
        LogEmergencyError errorInfo
    End If
    
    ' ハンドラーの呼び出し
    If mErrorHandlers.Exists(errorInfo.ErrorCode) Then
        Dim handler As IErrorHandler
        For Each handler In mErrorHandlers(errorInfo.ErrorCode)
            If handler.HandleError(errorInfo, options) Then
                RaiseEvent ErrorHandled(errorInfo, handler, True)
                IError_HandleError = True
                Exit Function
            End If
        Next handler
    End If
    
    ' リトライオプションの処理
    If (options And ehoRetry) = ehoRetry Then
        If AttemptRetry(errorInfo) Then
            IError_HandleError = True
            Exit Function
        End If
    End If
    
    ' ユーザー通知
    If (options And ehoNotifyUser) = ehoNotifyUser Then
        NotifyUser errorInfo
    End If
    
    IError_HandleError = False
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in HandleError: " & Err.Description
    IError_HandleError = False
End Function

Private Sub IError_RegisterHandler(ByVal errorCode As ErrorCode, ByVal handler As IErrorHandler, Optional ByVal options As HandlerOptions = hoNone)
    On Error GoTo ErrorHandler
    
    If Not mErrorHandlers.Exists(errorCode) Then
        Set mErrorHandlers(errorCode) = New Collection
    End If
    
    If (options And hoReplace) = hoReplace Then
        ' 既存のハンドラーをすべて削除
        Set mErrorHandlers(errorCode) = New Collection
        mErrorHandlers(errorCode).Add handler
    ElseIf (options And hoAppend) = hoAppend Then
        ' ハンドラーをリストの末尾に追加
        mErrorHandlers(errorCode).Add handler
    ElseIf (options And hoHighPriority) = hoHighPriority Then
        ' ハンドラーをリストの先頭に追加
        mErrorHandlers(errorCode).Add handler, Before:=1
    Else
        ' デフォルトではハンドラーをリストの末尾に追加
        mErrorHandlers(errorCode).Add handler
    End If
    
    RaiseEvent HandlerRegistered(errorCode, handler)
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in RegisterHandler: " & Err.Description
End Sub

Private Sub IError_UnregisterHandler(ByVal errorCode As ErrorCode, Optional ByVal options As UnregisterOptions = uoNone)
    On Error GoTo ErrorHandler
    
    If mErrorHandlers.Exists(errorCode) Then
        If (options And uoAllHandlers) = uoAllHandlers Then
            ' すべてのハンドラーを削除
            mErrorHandlers.Remove errorCode
        Else
            ' 最初のハンドラーを削除
            mErrorHandlers(errorCode).Remove 1
        End If
        
        RaiseEvent HandlerUnregistered(errorCode)
    End If
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in UnregisterHandler: " & Err.Description
End Sub

Private Sub IError_IncrementErrorCount(ByVal errorCode As ErrorCode, Optional ByVal options As CountOptions = coNone)
    On Error GoTo ErrorHandler
    
    If (options And coIncrement) = coIncrement Then
        mErrorStatistics.IncrementErrorCount errorCode
    ElseIf (options And coReset) = coReset Then
        ' カウントをリセット
        If mErrorStatistics.ErrorCountByCode.Exists(errorCode) Then
            mErrorStatistics.ErrorCountByCode(errorCode) = 0
        End If
    ElseIf (options And coDeduct) = coDeduct Then
        ' カウントを減少
        If mErrorStatistics.ErrorCountByCode.Exists(errorCode) Then
            If mErrorStatistics.ErrorCountByCode(errorCode) > 0 Then
                mErrorStatistics.ErrorCountByCode(errorCode) = mErrorStatistics.ErrorCountByCode(errorCode) - 1
            End If
        End If
    End If
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in IncrementErrorCount: " & Err.Description
End Sub

Private Function IError_CheckErrorThreshold(ByVal errorCode As ErrorCode, ByVal threshold As Long, Optional ByVal options As ThresholdOptions = toNone) As Boolean
    On Error GoTo ErrorHandler
    
    Dim count As Long
    
    If (options And toCheckSpecific) = toCheckSpecific Then
        ' 特定のエラーコードの閾値をチェック
        If mErrorStatistics.ErrorCountByCode.Exists(errorCode) Then
            count = mErrorStatistics.ErrorCountByCode(errorCode)
        End If
    Else
        ' グローバル閾値をチェック
        count = mErrorStatistics.TotalErrorCount
    End If
    
    If count >= threshold Then
        RaiseEvent ThresholdExceeded(errorCode, count, threshold)
        
        If (options And toResetAfterCheck) = toResetAfterCheck Then
            ' カウントをリセット
            If (options And toCheckSpecific) = toCheckSpecific Then
                If mErrorStatistics.ErrorCountByCode.Exists(errorCode) Then
                    mErrorStatistics.ErrorCountByCode(errorCode) = 0
                End If
            Else
                mErrorStatistics.Reset
            End If
        End If
        
        IError_CheckErrorThreshold = True
    Else
        IError_CheckErrorThreshold = False
    End If
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in CheckErrorThreshold: " & Err.Description
    IError_CheckErrorThreshold = False
End Function

Private Function IError_GetErrorHistory(Optional ByVal options As HistoryOptions = hoNone) As Collection
    On Error GoTo ErrorHandler
    
    Dim result As New Collection
    Dim item As ErrorInfo
    Dim i As Long
    
    ' フィルタリング
    For Each item In mErrorHistory
        If ShouldIncludeInHistory(item, options) Then
            result.Add item
        End If
    Next
    
    ' ソート（必要に応じて実装）
    
    ' 制限
    If (options And hoLimit) = hoLimit Then
        Dim limit As Long
        limit = 100 ' デフォルト制限
        
        If result.Count > limit Then
            Dim temp As New Collection
            For i = 1 To limit
                temp.Add result(i)
            Next i
            Set result = temp
        End If
    End If
    
    Set IError_GetErrorHistory = result
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in GetErrorHistory: " & Err.Description
    Set IError_GetErrorHistory = New Collection
End Function

Private Function IError_AnalyzeErrors(Optional ByVal options As AnalysisOptions = aoNone) As ErrorAnalysisResult
    On Error GoTo ErrorHandler
    
    Dim result As New ErrorAnalysisResult
    result.AnalysisStartTime = Now
    
    ' 基本分析
    If (options And aoBasicAnalysis) = aoBasicAnalysis Then
        PerformBasicAnalysis result
    End If
    
    ' 詳細分析
    If (options And aoDetailedAnalysis) = aoDetailedAnalysis Then
        PerformDetailedAnalysis result
    End If
    
    ' 根本原因分析
    If (options And aoRootCause) = aoRootCause Then
        AnalyzeRootCause result
    End If
    
    ' 影響範囲分析
    If (options And aoImpactAnalysis) = aoImpactAnalysis Then
        AnalyzeImpact result
    End If
    
    ' 推奨対策
    If (options And aoRecommendations) = aoRecommendations Then
        GenerateRecommendations result
    End If
    
    result.AnalysisEndTime = Now
    result.AnalysisDuration = DateDiff("s", result.AnalysisStartTime, result.AnalysisEndTime)
    
    RaiseEvent ErrorAnalysisCompleted(result)
    Set IError_AnalyzeErrors = result
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in AnalyzeErrors: " & Err.Description
    Set IError_AnalyzeErrors = New ErrorAnalysisResult
End Function

Private Function IError_CategorizeError(ByRef errorInfo As ErrorInfo, Optional ByVal options As CategoryOptions = coNone) As ErrorCategory
    On Error GoTo ErrorHandler
    
    ' デフォルトでは自動分類を試みる
    Select Case errorInfo.ErrorCode
        Case ecInvalidInput
            IError_CategorizeError = ecInputError
        Case ecDatabaseError
            IError_CategorizeError = ecDatabaseError
        Case ecNetworkError
            IError_CategorizeError = ecNetworkError
        Case ecFileNotFound, ecFileAccessDenied
            IError_CategorizeError = ecFileIOError
        Case ecSecurityError
            IError_CategorizeError = ecSecurityError
        Case ecConfigurationError
            IError_CategorizeError = ecConfigurationError
        Case Else
            IError_CategorizeError = ecUnknownCategory
    End Select
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in CategorizeError: " & Err.Description
    IError_CategorizeError = ecUnknownCategory
End Function

Private Function IError_GetErrorStatistics(Optional ByVal options As StatisticsOptions = soNone) As ErrorStatistics
    Set IError_GetErrorStatistics = mErrorStatistics
End Function

Private Function IError_GetPerformanceMetrics() As Collection
    Dim metrics As New Collection
    
    ' 基本的なメトリクスを収集
    metrics.Add Array("TotalErrors", mErrorStatistics.TotalErrorCount)
    metrics.Add Array("ErrorRate", mErrorStatistics.ErrorRate)
    metrics.Add Array("AverageHandlingTime", mErrorStatistics.AverageHandlingTime)
    metrics.Add Array("ResolutionRate", mErrorStatistics.ResolutionRate)
    
    Set IError_GetPerformanceMetrics = metrics
End Function

Private Function IError_Validate() As Boolean
    IError_Validate = mIsInitialized
End Function

Private Sub IError_Cleanup()
    Set mErrorHandlers = Nothing
    Set mErrorHistory = Nothing
    Set mErrorStatistics = Nothing
    Set mLastError = Nothing
    mIsInitialized = False
End Sub

' プライベートヘルパーメソッド

Private Sub LogError(ByVal errorInfo As ErrorInfo)
    ' ログ記録の実装
    Debug.Print "Error logged: " & errorInfo.Message
End Sub

Private Sub LogEmergencyError(ByVal errorInfo As ErrorInfo)
    ' エマージェンシーログの実装
    Debug.Print "Emergency error logged: " & errorInfo.Message
End Sub

Private Function AttemptRetry(ByVal errorInfo As ErrorInfo) As Boolean
    Dim attempt As Long
    
    For attempt = 1 To MAX_RETRY_ATTEMPTS
        Sleep RETRY_DELAY_MS
        
        ' リトライ処理の実装
        ' ここでは単純な例として、3回目で成功すると仮定
        If attempt = 3 Then
            AttemptRetry = True
            Exit Function
        End If
    Next attempt
    
    AttemptRetry = False
End Function

Private Sub NotifyUser(ByVal errorInfo As ErrorInfo)
    ' ユーザー通知の実装
    MsgBox errorInfo.Message, vbExclamation, "エラー"
End Sub

Private Function ShouldIncludeInHistory(ByVal errorInfo As ErrorInfo, ByVal options As HistoryOptions) As Boolean
    ' フィルタリング条件の実装
    ShouldIncludeInHistory = True
End Function

Private Sub PerformBasicAnalysis(ByRef result As ErrorAnalysisResult)
    ' 基本分析の実装
End Sub

Private Sub PerformDetailedAnalysis(ByRef result As ErrorAnalysisResult)
    ' 詳細分析の実装
End Sub

Private Sub AnalyzeRootCause(ByRef result As ErrorAnalysisResult)
    ' 根本原因分析の実装
End Sub

Private Sub AnalyzeImpact(ByRef result As ErrorAnalysisResult)
    ' 影響範囲分析の実装
End Sub

Private Sub GenerateRecommendations(ByRef result As ErrorAnalysisResult)
    ' 推奨対策の生成
End Sub

Private Sub Sleep(ByVal milliseconds As Long)
    ' スリープ処理の実装
    Application.Wait Now + TimeSerial(0, 0, milliseconds / 1000)
End Sub