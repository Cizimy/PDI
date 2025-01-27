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
Private mMessageResource As ErrorMessageResource
Private mRecoveryStrategies As Collection

' 定数
Private Const MAX_RETRY_ATTEMPTS As Long = 3
Private Const RETRY_DELAY_MS As Long = 1000
Private Const MAX_EMERGENCY_LOG_SIZE As Long = 1000
Private Const EMERGENCY_LOG_FLUSH_INTERVAL As Long = 1000
Private Const MAX_RECOVERY_CHAIN_LENGTH As Long = 5

' イベント定義
Public Event EmergencyLogWritten(ByVal message As String, ByVal severity As EmergencyLevel)
Public Event RecoveryProgressChanged(ByVal errorInfo As ErrorInfo, ByVal progressPercentage As Long)
Public Event RecoveryChainExecuted(ByVal chainId As String, ByVal successCount As Long, ByVal failureCount As Long)

Private Sub Class_Initialize()
    Set mErrorHandlers = CreateObject("Scripting.Dictionary")
    Set mErrorHistory = New Collection
    Set mErrorStatistics = New ErrorStatistics
    Set mMessageResource = New ErrorMessageResource
    Set mRecoveryStrategies = New Collection
    mIsInitialized = True
End Sub

Private Sub Class_Terminate()
    Set mErrorHandlers = Nothing
    Set mErrorHistory = Nothing
    Set mErrorStatistics = Nothing
    Set mLastError = Nothing
End Sub

Private Function IError_HandleError(ByRef errorInfo As ErrorInfo, Optional ByVal options As ErrorHandlingOptions = ehoNone) As HandlingResult
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
            Dim handlingResult As HandlingResult
            handlingResult = handler.HandleError(errorInfo, options)
            
            Select Case handlingResult
                Case hrSuccess
                    RaiseEvent ErrorHandled(errorInfo, handler, True)
                    IError_HandleError = hrSuccess
                    Exit Function
                    
                Case hrPartialSuccess
                    RaiseEvent ErrorHandled(errorInfo, handler, True)
                    IError_HandleError = hrPartialSuccess
                    Exit Function
                    
                Case hrRetryNeeded
                    ' リトライが必要な場合は次のハンドラーに進む
                    
                Case hrFallbackNeeded
                    ' フォールバックが必要な場合は次のハンドラーに進む
                    
                Case hrFailed, hrUnhandled
                    ' 失敗または未処理の場合は次のハンドラーに進む
            End Select
        Next handler
    End If
    
    ' リトライオプションの処理
    If (options And ehoRetry) = ehoRetry Then
        Dim retryResult As HandlingResult
        retryResult = AttemptRetry(errorInfo)
        
        Select Case retryResult
            Case hrSuccess, hrPartialSuccess
                IError_HandleError = retryResult
                Exit Function
                
            Case hrRetryNeeded, hrFallbackNeeded
                ' リトライまたはフォールバックが必要な場合は回復チェーンを試行
                Dim recoveryChain As Collection
                Set recoveryChain = CreateRecoveryChain(errorInfo)
                
                If recoveryChain.Count > 0 Then
                    If ExecuteRecoveryChain(recoveryChain, errorInfo) Then
                        IError_HandleError = hrSuccess
                        Exit Function
                    End If
                End If
                
                ' 回復チェーンも失敗した場合
                IError_HandleError = hrFailed
                
            Case hrFailed, hrUnhandled
                ' リトライ失敗時は回復チェーンを試行
                Set recoveryChain = CreateRecoveryChain(errorInfo)
                
                If recoveryChain.Count > 0 Then
                    If ExecuteRecoveryChain(recoveryChain, errorInfo) Then
                        IError_HandleError = hrSuccess
                        Exit Function
                    End If
                End If
        End Select
    End If
    
    ' ユーザー通知
    If (options And ehoNotifyUser) = ehoNotifyUser Then
        NotifyUser errorInfo
    End If
    
    IError_HandleError = hrUnhandled
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
    On Error GoTo ErrorHandler
    
    ' 多言語対応メッセージの生成
    Dim message As String
    message = mMessageResource.GetMessage(errorInfo.ErrorCode)
    
    ' エラーメッセージのフォーマット
    Dim formattedMessage As String
    formattedMessage = FormatErrorMessage(message, errorInfo)
    
    ' イベントログへの出力を試みる
    Dim eventLogSuccess As Boolean
    eventLogSuccess = False
    
    On Error Resume Next
    WriteToEventLog formattedMessage
    If Err.Number = 0 Then
        eventLogSuccess = True
    End If
    On Error GoTo ErrorHandler
    
    ' イベントログ書き込み失敗時はファイルへの出力を試みる
    If Not eventLogSuccess Then
        Dim fileSuccess As Boolean
        fileSuccess = False
        
        On Error Resume Next
        WriteToEmergencyFile formattedMessage
        If Err.Number = 0 Then
            fileSuccess = True
        End If
        On Error GoTo ErrorHandler
        
        ' 両方の出力が失敗した場合はエラーを発生させる
        If Not fileSuccess Then
            Err.Raise vbObjectError + 1, "LogEmergencyError", _
                     "Failed to write to both event log and emergency file"
        End If
    End If
    
    ' イベントの発行
    RaiseEvent EmergencyLogWritten(message, errorInfo.Severity)
    
    Exit Sub

ErrorHandler:
    ' エラーを上位に伝播させる
    Err.Raise Err.Number, "LogEmergencyError", _
              "Emergency logging failed: " & Err.Description
End Sub

Private Function FormatErrorMessage(ByVal message As String, ByRef errorInfo As ErrorInfo) As String
    Dim builder As String
    builder = Now & " [EMERGENCY] " & message
    
    With errorInfo
        builder = builder & vbCrLf & _
                 "コード: " & .ErrorCode & vbCrLf & _
                 "重大度: " & .Severity & vbCrLf & _
                 "説明: " & .Description & vbCrLf & _
                 "ソース: " & .SourceModule & "." & .SourceProcedure & vbCrLf & _
                 "発生時刻: " & .OccurredAt & vbCrLf & _
                 "スタックトレース: " & vbCrLf & .StackTrace
    End With
    
    FormatErrorMessage = builder
End Function

Private Sub WriteToEventLog(ByVal message As String)
    On Error Resume Next
    modWindowsAPI.WriteToEventLog "PDI Error", message, EVENTLOG_ERROR_TYPE
End Sub

Private Sub WriteToEmergencyFile(ByVal message As String)
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim logPath As String
    logPath = GetEmergencyLogPath()
    
    Dim logFile As Object
    Set logFile = fso.OpenTextFile(logPath, 8, True) ' 8: ForAppending, Create if doesn't exist
    logFile.WriteLine message
    logFile.Close
    
    Set logFile = Nothing
    Set fso = Nothing
End Sub

Private Function GetEmergencyLogPath() As String
    ' エマージェンシーログのパスを取得
    GetEmergencyLogPath = App.Path & "\emergency.log"
End Function

Private Function AttemptRetry(ByVal errorInfo As ErrorInfo) As HandlingResult
    On Error GoTo ErrorHandler
    
    Dim attempt As Long
    Dim handler As IErrorHandler
    Dim backoffInterval As Long
    Dim retryCount As Long
    Dim retryInterval As Long
    Dim retryOptions As RetryOptions
    
    ' エラーコードに対応するハンドラーを取得
    If Not mErrorHandlers.Exists(errorInfo.ErrorCode) Then
        AttemptRetry = hrUnhandled
        Exit Function
    End If
    
    ' ハンドラーからリトライ設定を取得
    Set handler = mErrorHandlers(errorInfo.ErrorCode)(1)
    retryCount = handler.MaxRetryCount
    retryInterval = handler.RetryInterval
    
    ' リトライオプションの設定
    Set retryOptions = New RetryOptions
    retryOptions.UseExponentialBackoff = True
    retryOptions.MaxAttempts = retryCount
    retryOptions.BaseInterval = retryInterval
    
    For attempt = 1 To retryCount
        ' 指数バックオフの計算
        backoffInterval = CalculateBackoffInterval(attempt, retryInterval)
        Sleep backoffInterval
        
        ' 進捗の通知
        RaiseEvent RecoveryProgressChanged(errorInfo, (attempt / retryCount) * 100)
        
        ' 各ハンドラーでリトライを試行
        For Each handler In mErrorHandlers(errorInfo.ErrorCode)
            ' リトライ操作の作成
            Dim operation As IRetryableOperation
            Set operation = CreateRetryOperation(errorInfo, handler)
            
            ' リトライの実行
            Dim result As OperationResult
            Set result = handler.RetryOperation(operation, retryOptions)
            
            ' リトライ試行のイベント発行（バックオフ間隔を含む）
            RaiseEvent RetryAttempted(errorInfo, attempt, backoffInterval, IIf(result.Success, hrSuccess, hrFailed))
            
            If result.Success Then
                ' 統計情報の更新
                mErrorStatistics.IncrementRetrySuccess
                AttemptRetry = hrSuccess
                Exit Function
            End If
        Next handler
        
        ' 統計情報の更新
        errorInfo.RetryCount = attempt
        errorInfo.LastRetryTimestamp = Now
        
        ' エラーキューの処理（必要な場合のみ）
        If ShouldProcessQueue(errorInfo) Then
            ProcessErrorQueue errorInfo
            ' キュー処理後も失敗している場合はフォールバックが必要
            AttemptRetry = hrFallbackNeeded
            Exit Function
        End If
    Next attempt
    
    ' すべてのリトライが失敗した場合
    AttemptRetry = hrFailed
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in AttemptRetry: " & Err.Description
    AttemptRetry = hrFailed
End Function

Private Function CalculateBackoffInterval(ByVal attempt As Long, ByVal baseInterval As Long) As Long
    ' 指数バックオフの計算
    Dim interval As Long
    interval = baseInterval * (2 ^ (attempt - 1))
    
    ' 最大値を制限
    If interval > MAX_BACKOFF_INTERVAL Then
        interval = MAX_BACKOFF_INTERVAL
    End If
    
    CalculateBackoffInterval = interval
End Function

Private Function CreateRetryOperation(ByVal errorInfo As ErrorInfo, ByVal handler As IErrorHandler) As IRetryableOperation
    ' リトライ可能な操作を作成
    Dim operation As New RetryableOperation
    operation.Initialize errorInfo, handler
    Set CreateRetryOperation = operation
End Function

Private Function ShouldProcessQueue(ByVal errorInfo As ErrorInfo) As Boolean
    ' キュー処理の必要性を判断
    
    ' 以下の条件でキューを処理する：
    ' 1. リトライ回数が一定回数（MAX_RETRY_ATTEMPTS/2）を超えた場合
    ' 2. エラーの重大度が高い（Critical以上）場合
    ' 3. 同じエラーが短時間に複数回発生している場合
    
    ' リトライ回数による判断
    If errorInfo.RetryCount > MAX_RETRY_ATTEMPTS \ 2 Then
        ShouldProcessQueue = True
        Exit Function
    End If
    
    ' エラーの重大度による判断
    If errorInfo.Severity >= esCritical Then
        ShouldProcessQueue = True
        Exit Function
    End If
    
    ' エラーの発生頻度による判断
    Dim errorCount As Long
    If mErrorStatistics.ErrorCountByCode.Exists(errorInfo.ErrorCode) Then
        errorCount = mErrorStatistics.ErrorCountByCode(errorInfo.ErrorCode)
        
        ' 直近のエラー発生時刻との間隔を確認
        If errorCount > 0 And Not errorInfo.LastRetryTimestamp = #12:00:00 AM# Then
            Dim timeSinceLastError As Long
            timeSinceLastError = DateDiff("s", errorInfo.LastRetryTimestamp, Now)
            
            ' 5分以内に3回以上発生している場合
            If timeSinceLastError <= 300 And errorCount >= 3 Then
                ShouldProcessQueue = True
                Exit Function
            End If
        End If
    End If
    
    ShouldProcessQueue = False
End Function

Private Sub ProcessErrorQueue(ByRef errorInfo As ErrorInfo)
    On Error GoTo ErrorHandler
    
    ' エラーキューの処理をIErrorStrategyに委譲
    Dim strategy As IErrorStrategy
    For Each strategy In mRecoveryStrategies
        If strategy.SupportsErrorCode(errorInfo.ErrorCode) Then
            ' RecoverFromErrorメソッドを使用してエラー回復を試みる
            If strategy.RecoverFromError(errorInfo) Then
                ' 回復成功時の処理
                mErrorStatistics.IncrementRecoverySuccess
                Exit Sub
            End If
        End If
    Next strategy
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ProcessErrorQueue: " & Err.Description
End Sub

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

'#Region "Recovery Chain Management"

Private Function CreateRecoveryChain(ByVal errorInfo As ErrorInfo) As RecoveryChain
    On Error GoTo ErrorHandler
    
    Dim chain As New RecoveryChain
    Dim strategy As IErrorStrategy
    Dim metadata As New Dictionary
    
    ' チェーンのメタデータを設定
    metadata.Add "ErrorCode", errorInfo.ErrorCode
    metadata.Add "Severity", errorInfo.Severity
    metadata.Add "CreatedAt", Now
    metadata.Add "MaxLength", MAX_RECOVERY_CHAIN_LENGTH
    
    chain.Initialize metadata
    
    ' エラーコードに基づいて適切な回復戦略を選択
    For Each strategy In mRecoveryStrategies
        If strategy.SupportsErrorCode(errorInfo.ErrorCode) Then
            ' 戦略の詳細情報を追加（必要最小限の情報のみ）
            Dim strategyInfo As New Dictionary
            strategyInfo.Add "Priority", strategy.Priority
            strategyInfo.Add "Description", strategy.GetDescription()
            
            ' 戦略をチェーンに追加（プロパティは直接アクセス）
            chain.AddStrategy strategy, strategyInfo
        End If
        
        ' チェーンの長さを制限
        If chain.StrategyCount >= MAX_RECOVERY_CHAIN_LENGTH Then
            Exit For
        End If
    Next strategy
    
    ' 優先度でソート
    chain.SortByPriority
    
    ' チェーンの検証
    If Not chain.Validate Then
        Err.Raise vbObjectError + 1, "CreateRecoveryChain", "Invalid recovery chain"
    End If
    
    Set CreateRecoveryChain = chain
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in CreateRecoveryChain: " & Err.Description
    Set CreateRecoveryChain = New RecoveryChain
End Function

Private Sub SortRecoveryChain(ByRef chain As Collection)
    ' 優先度に基づいてチェーンをソート
    Dim i As Long, j As Long
    Dim tempStrategy As IErrorStrategy
    
    For i = 1 To chain.Count - 1
        For j = i + 1 To chain.Count
            If chain(i).Priority < chain(j).Priority Then
                Set tempStrategy = chain(i)
                Set chain(i) = chain(j)
                Set chain(j) = tempStrategy
            End If
        Next j
    Next i
End Sub

Private Function ExecuteRecoveryChain(ByVal chain As RecoveryChain, ByRef errorInfo As ErrorInfo) As Boolean
    On Error GoTo ErrorHandler
    
    Dim strategy As IErrorStrategy
    Dim strategyInfo As Dictionary
    Dim executionContext As New ExecutionContext
    Dim progress As Long
    
    ' 実行コンテキストの初期化
    executionContext.Initialize chain.ChainId, errorInfo
    
    ' チェーンの実行前検証
    On Error Resume Next
    Dim isValid As Boolean
    isValid = chain.Validate
    If Err.Number <> 0 Then
        ' Validateメソッドでエラーが発生した場合
        Dim validateError As Long
        validateError = Err.Number
        Dim validateDesc As String
        validateDesc = Err.Description
        On Error GoTo ErrorHandler
        Err.Raise validateError, "ExecuteRecoveryChain", _
                "Chain validation failed: " & validateDesc
    End If
    On Error GoTo ErrorHandler
    
    ' 検証結果の確認
    If Not isValid Then
        Err.Raise vbObjectError + 1, "ExecuteRecoveryChain", _
                "Recovery chain validation failed: Chain is in invalid state"
    End If
    
    ' 各戦略を実行
    For Each strategy In chain.Strategies
        ' 戦略の情報を取得
        Set strategyInfo = chain.GetStrategyInfo(strategy)
        
        ' 進捗状況の更新と通知
        progress = chain.CalculateProgress()
        RaiseEvent RecoveryProgressChanged(errorInfo, progress)
        
        ' 実行コンテキストの更新
        executionContext.CurrentStrategy = strategyInfo("Description")
        executionContext.CurrentPriority = CLng(strategyInfo("Priority"))
        
        ' トランザクションのサポートを確認
        If strategyInfo("SupportsTransactions") Then
            ' トランザクション内で実行
            If ExecuteStrategyWithTransaction(strategy, errorInfo, executionContext) Then
                chain.IncrementSuccess
                ExecuteRecoveryChain = True
                Exit For
            End If
        Else
            ' 通常実行
            If strategy.RecoverFromError(errorInfo) Then
                chain.IncrementSuccess
                ExecuteRecoveryChain = True
                Exit For
            End If
        End If
        
        ' 失敗をカウント
        chain.IncrementFailure
        
        ' フォールバック戦略の確認と実行
        If strategy.FallbackStrategy IsNot Nothing Then
            executionContext.UsingFallback = True
            If ExecuteStrategyWithFallback(strategy.FallbackStrategy, errorInfo, executionContext) Then
                chain.IncrementSuccessCount
                ExecuteRecoveryChain = True
                Exit For
            End If
        End If
    Next strategy
    
    ' 実行結果の記録
    executionContext.UpdateDuration
    chain.RecordExecution executionContext.ToDictionary
    
    ' チェーン実行結果の通知
    RaiseEvent RecoveryChainExecuted(chain.ChainId, chain.SuccessCount, chain.FailureCount)
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in ExecuteRecoveryChain: " & Err.Description
    If Not chain Is Nothing Then
        chain.RecordError Err.Description
    End If
    ExecuteRecoveryChain = False
End Function

Private Function ExecuteStrategyWithTransaction(ByVal strategy As IErrorStrategy, ByRef errorInfo As ErrorInfo, ByVal context As Dictionary) As Boolean
    On Error GoTo ErrorHandler
    
    ' トランザクションの開始
    BeginTransaction
    
    If strategy.RecoverFromError(errorInfo) Then
        CommitTransaction
        ExecuteStrategyWithTransaction = True
    Else
        RollbackTransaction
        ExecuteStrategyWithTransaction = False
    End If
    Exit Function
    
ErrorHandler:
    If IsInTransaction Then
        RollbackTransaction
    End If
    ' エラーを上位に伝播
    Err.Raise Err.Number, "ExecuteStrategyWithTransaction", _
              "Transaction execution failed: " & Err.Description
End Function

Private Function ExecuteStrategyWithFallback(ByVal fallbackStrategy As IErrorStrategy, ByRef errorInfo As ErrorInfo, ByVal context As ExecutionContext) As Boolean
    On Error GoTo ErrorHandler
    
    ' フォールバック実行前の状態を保存
    Dim originalContext As ExecutionContext
    Set originalContext = context.Clone
    
    If fallbackStrategy.RecoverFromError(errorInfo) Then
        ExecuteStrategyWithFallback = True
    Else
        ' 状態を復元
        RestoreContext context, originalState
        ExecuteStrategyWithFallback = False
    End If
    Exit Function
    
ErrorHandler:
    ' エラーを上位に伝播
    Err.Raise Err.Number, "ExecuteStrategyWithFallback", _
              "Fallback strategy execution failed: " & Err.Description
End Function

Private Function CopyDictionary(ByVal source As Dictionary) As Dictionary
    Dim result As New Dictionary
    Dim key As Variant
    
    For Each key In source.Keys
        result.Add key, source(key)
    Next
    
    Set CopyDictionary = result
End Function

Private Sub RestoreContext(ByRef target As Dictionary, ByVal source As Dictionary)
    Dim key As Variant
    
    target.RemoveAll
    For Each key In source.Keys
        target.Add key, source(key)
    Next
End Sub

'#End Region

Private Sub Sleep(ByVal milliseconds As Long)
    ' スリープ処理の実装
    Application.Wait Now + TimeSerial(0, 0, milliseconds / 1000)
End Sub