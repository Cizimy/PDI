VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDefaultErrorStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

' インターフェース実装
Implements IErrorStrategy

' プライベート変数
Private mMaxRetryCount As Long           ' 最大リトライ回数
Private mRetryInterval As Long           ' リトライ間隔（ミリ秒）
Private mPriority As Long                ' 戦略の優先度
Private mFallbackStrategy As IErrorStrategy ' フォールバック戦略
Private mMetrics As Object               ' パフォーマンスメトリクス
Private mProgress As RecoveryProgress    ' 回復の進捗状況
Private mLogger As Object                ' ロギング用オブジェクト
Private mIsInitialized As Boolean        ' 初期化フラグ
Private mIsCancellationRequested As Boolean ' キャンセル要求フラグ

' 定数定義
Private Const DEFAULT_MAX_RETRY_COUNT As Long = 3
Private Const DEFAULT_RETRY_INTERVAL_MS As Long = 1000
Private Const MAX_RECOVERY_ATTEMPTS As Long = 5
Private Const RECOVERY_TIMEOUT_MS As Long = 30000
Private Const LOG_FILE_PATH As String = "error_strategy.log"

' イベント定義
Public Event RecoveryChainStarted(ByVal chainId As String, ByVal strategies As Collection)
Public Event RecoveryChainCompleted(ByVal chainId As String, ByVal successCount As Long, ByVal failureCount As Long)
Public Event FallbackStrategyActivated(ByVal errorInfo As ErrorInfo, ByVal fallbackStrategy As IErrorStrategy, ByVal reason As FallbackReason)
Public Event StrategyExecutionStarted(ByVal strategyId As String, ByVal priority As Long)
Public Event StrategyExecutionCompleted(ByVal strategyId As String, ByVal success As Boolean)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 初期化・終了処理
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()
    ' デフォルト値の設定
    mMaxRetryCount = DEFAULT_MAX_RETRY_COUNT
    mRetryInterval = DEFAULT_RETRY_INTERVAL_MS
    mPriority = 0
    Set mMetrics = CreateObject("Scripting.Dictionary")
    Set mProgress = New RecoveryProgress
    Set mLogger = CreateObject("Scripting.FileSystemObject")
    mIsInitialized = True
    mIsCancellationRequested = False
    
    ' メトリクスの初期化
    InitializeMetrics
End Sub

Private Sub Class_Terminate()
    ' リソースの解放
    Set mFallbackStrategy = Nothing
    Set mMetrics = Nothing
    Set mProgress = Nothing
    Set mLogger = Nothing
    mIsInitialized = False
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IErrorStrategy インターフェース実装
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function IErrorStrategy_RecoverFromError(ByRef errorInfo As ErrorInfo, _
                                              ByVal errorManager As clsErrorManager, _
                                              Optional ByVal context As RecoveryContext, _
                                              Optional ByVal options As RecoveryOptions) As Boolean
    On Error GoTo ErrorHandler
    
    ' 初期化チェック
    If Not mIsInitialized Then
        LogError "RecoverFromError", "Strategy not initialized"
        Exit Function
    End If
    
    ' エラー情報の検証
    If Not ValidateErrorInfo(errorInfo) Then
        LogError "RecoverFromError", "Invalid error info"
        Exit Function
    End If
    
    ' 進捗状況の初期化
    mProgress.TotalSteps = MAX_RECOVERY_ATTEMPTS
    mProgress.CurrentStep = 0
    mProgress.Status = rsInProgress
    
    ' 回復処理の開始
    Dim startTime As Date
    startTime = Now
    
    ' メトリクスの更新
    UpdateMetrics "RecoveryAttempts", 0
    UpdateMetrics "StartTime", startTime
    
    ' 回復チェーンの作成と実行
    Dim chain As RecoveryChain
    Set chain = CreateRecoveryChainInternal(errorInfo, errorManager)
    
    If chain Is Nothing Then
        LogError "RecoverFromError", "Failed to create recovery chain"
        Exit Function
    End If
    
    ' イベント通知
    RaiseEvent RecoveryChainStarted(chain.ChainId, chain.Strategies)
    
    ' 回復処理の実行
    Dim success As Boolean
    success = ExecuteRecoveryChain(chain, errorInfo, errorManager)
    
    ' 結果の記録
    UpdateMetrics "EndTime", Now
    UpdateMetrics "Success", success
    
    ' イベント通知
    RaiseEvent RecoveryChainCompleted(chain.ChainId, chain.SuccessCount, chain.FailureCount)
    
    IErrorStrategy_RecoverFromError = success
    Exit Function
    
ErrorHandler:
    LogError "RecoverFromError", "Error: " & Err.Description
    IErrorStrategy_RecoverFromError = False
End Function

Private Function IErrorStrategy_CreateRecoveryChain(ByVal strategies As Collection, _
                                                 ByVal errorManager As clsErrorManager) As RecoveryChain
    On Error GoTo ErrorHandler
    
    ' 新しい回復チェーンの作成
    Dim chain As New RecoveryChain
    Dim metadata As New Dictionary
    
    ' メタデータの設定
    metadata.Add "CreatedAt", Now
    metadata.Add "MaxLength", MAX_RECOVERY_CHAIN_LENGTH
    
    chain.Initialize metadata
    
    ' 戦略の追加
    Dim strategy As IErrorStrategy
    For Each strategy In strategies
        ' 戦略の情報を取得
        Dim strategyInfo As New Dictionary
        strategyInfo.Add "Priority", strategy.Priority
        strategyInfo.Add "Description", strategy.GetDescription()
        
        ' チェーンに追加
        chain.AddStrategy strategy, strategyInfo
        
        ' チェーンの長さを制限
        If chain.StrategyCount >= MAX_RECOVERY_CHAIN_LENGTH Then
            Exit For
        End If
    Next strategy
    
    ' 優先度でソート
    chain.SortByPriority
    
    Set IErrorStrategy_CreateRecoveryChain = chain
    Exit Function
    
ErrorHandler:
    LogError "CreateRecoveryChain", "Error: " & Err.Description
    Set IErrorStrategy_CreateRecoveryChain = Nothing
End Function

Private Property Let IErrorStrategy_MaxRetryCount(ByVal value As Long)
    If value >= 0 Then
        mMaxRetryCount = value
    End If
End Property

Private Property Get IErrorStrategy_MaxRetryCount() As Long
    IErrorStrategy_MaxRetryCount = mMaxRetryCount
End Property

Private Property Let IErrorStrategy_RetryInterval(ByVal value As Long)
    If value >= 0 Then
        mRetryInterval = value
    End If
End Property

Private Property Get IErrorStrategy_RetryInterval() As Long
    IErrorStrategy_RetryInterval = mRetryInterval
End Property

Private Property Let IErrorStrategy_Priority(ByVal value As Long)
    mPriority = value
End Property

Private Property Get IErrorStrategy_Priority() As Long
    IErrorStrategy_Priority = mPriority
End Property

Private Function IErrorStrategy_GetDescription(Optional ByVal locale As String) As String
    IErrorStrategy_GetDescription = "Default Error Recovery Strategy" & vbCrLf & _
                                  "Max Retry Count: " & mMaxRetryCount & vbCrLf & _
                                  "Retry Interval: " & mRetryInterval & "ms" & vbCrLf & _
                                  "Priority: " & mPriority
End Function

Private Property Set IErrorStrategy_FallbackStrategy(ByVal strategy As IErrorStrategy)
    Set mFallbackStrategy = strategy
End Property

Private Property Get IErrorStrategy_FallbackStrategy() As IErrorStrategy
    Set IErrorStrategy_FallbackStrategy = mFallbackStrategy
End Property

Private Function IErrorStrategy_GetProgress(Optional ByVal includeMetrics As Boolean = False) As RecoveryProgress
    If includeMetrics Then
        UpdateProgressMetrics
    End If
    Set IErrorStrategy_GetProgress = mProgress
End Function

Private Sub IErrorStrategy_SetRecoveryHandler(ByVal errorCode As ErrorCode, _
                                           ByVal handler As IErrorHandler, _
                                           Optional ByVal options As StrategyOptions)
    ' ハンドラーの登録処理
    On Error GoTo ErrorHandler
    
    If handler Is Nothing Then Exit Sub
    
    ' メトリクスの更新
    UpdateMetrics "RegisteredHandlers", mMetrics("RegisteredHandlers") + 1
    Exit Sub
    
ErrorHandler:
    LogError "SetRecoveryHandler", "Error: " & Err.Description
End Sub

Private Function IErrorStrategy_GetRecoveryHistory(Optional ByVal filter As RecoveryHistoryFilter) As Collection
    ' 履歴情報の取得
    Dim history As New Collection
    ' TODO: 履歴情報の実装
    Set IErrorStrategy_GetRecoveryHistory = history
End Function

Private Function IErrorStrategy_Validate(Optional ByVal validationLevel As ValidationLevel = vlNormal) As Boolean
    On Error GoTo ErrorHandler
    
    ' 基本的な検証
    If Not mIsInitialized Then Exit Function
    If mMaxRetryCount < 0 Then Exit Function
    If mRetryInterval < 0 Then Exit Function
    
    ' 詳細な検証（validationLevelに応じて）
    Select Case validationLevel
        Case vlStrict
            ' より厳密な検証
            If mMetrics Is Nothing Then Exit Function
            If mProgress Is Nothing Then Exit Function
            If mLogger Is Nothing Then Exit Function
    End Select
    
    IErrorStrategy_Validate = True
    Exit Function
    
ErrorHandler:
    LogError "Validate", "Error: " & Err.Description
    IErrorStrategy_Validate = False
End Function

Private Sub IErrorStrategy_Cleanup(Optional ByVal options As CleanupOptions)
    On Error Resume Next
    
    ' リソースの解放
    Set mFallbackStrategy = Nothing
    Set mMetrics = Nothing
    Set mProgress = Nothing
    
    ' メトリクスの保存
    SaveMetrics
    
    ' ログファイルのクローズ
    CloseLogFile
End Sub

Private Sub IErrorStrategy_CancelRecovery(Optional ByVal options As CancellationOptions)
    mIsCancellationRequested = True
End Sub

Private Property Get IErrorStrategy_IsCancellationRequested() As Boolean
    IErrorStrategy_IsCancellationRequested = mIsCancellationRequested
End Property

Private Property Get IErrorStrategy_SupportsAsyncRecovery() As Boolean
    IErrorStrategy_SupportsAsyncRecovery = False ' 現在は非同期回復をサポートしていない
End Property

Private Property Get IErrorStrategy_SupportsTransactions() As Boolean
    IErrorStrategy_SupportsTransactions = True ' トランザクションをサポート
End Property

Private Function IErrorStrategy_GetPerformanceMetrics() As Collection
    Dim metrics As New Collection
    
    ' 基本メトリクスの収集
    With mMetrics
        metrics.Add Array("TotalAttempts", .Item("RecoveryAttempts"))
        metrics.Add Array("SuccessRate", CalculateSuccessRate())
        metrics.Add Array("AverageRecoveryTime", CalculateAverageRecoveryTime())
    End With
    
    Set IErrorStrategy_GetPerformanceMetrics = metrics
End Function

Private Function IErrorStrategy_AnalyzeErrorPatterns(ByVal timeWindow As Long) As Collection
    Dim patterns As New Collection
    ' TODO: エラーパターンの分析実装
    Set IErrorStrategy_AnalyzeErrorPatterns = patterns
End Function

Private Property Let IErrorStrategy_BackoffStrategy(ByVal strategy As IBackoffStrategy)
    ' バックオフ戦略の設定
    ' TODO: バックオフ戦略の実装
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' プライベートヘルパーメソッド
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub InitializeMetrics()
    With mMetrics
        .Add "RecoveryAttempts", 0
        .Add "SuccessCount", 0
        .Add "FailureCount", 0
        .Add "TotalTime", 0
        .Add "RegisteredHandlers", 0
    End With
End Sub

Private Sub UpdateMetrics(ByVal metricName As String, ByVal value As Variant)
    If mMetrics.Exists(metricName) Then
        mMetrics(metricName) = value
    Else
        mMetrics.Add metricName, value
    End If
End Sub

Private Function CalculateSuccessRate() As Double
    Dim totalAttempts As Long
    totalAttempts = mMetrics("RecoveryAttempts")
    
    If totalAttempts = 0 Then
        CalculateSuccessRate = 0
        Exit Function
    End If
    
    CalculateSuccessRate = (mMetrics("SuccessCount") / totalAttempts) * 100
End Function

Private Function CalculateAverageRecoveryTime() As Double
    Dim totalAttempts As Long
    totalAttempts = mMetrics("RecoveryAttempts")
    
    If totalAttempts = 0 Then
        CalculateAverageRecoveryTime = 0
        Exit Function
    End If
    
    CalculateAverageRecoveryTime = mMetrics("TotalTime") / totalAttempts
End Function

Private Sub UpdateProgressMetrics()
    With mProgress
        .AddMetric "SuccessRate", CalculateSuccessRate()
        .AddMetric "AverageTime", CalculateAverageRecoveryTime()
        .AddMetric "TotalAttempts", mMetrics("RecoveryAttempts")
    End With
End Sub

Private Function ValidateErrorInfo(ByRef errorInfo As ErrorInfo) As Boolean
    If errorInfo Is Nothing Then Exit Function
    If Len(errorInfo.Message) = 0 Then Exit Function
    If errorInfo.ErrorCode = 0 Then Exit Function
    
    ValidateErrorInfo = True
End Function

Private Function CreateRecoveryChainInternal(ByRef errorInfo As ErrorInfo, _
                                          ByVal errorManager As clsErrorManager) As RecoveryChain
    On Error GoTo ErrorHandler
    
    Dim chain As New RecoveryChain
    Dim metadata As New Dictionary
    
    ' メタデータの設定
    With metadata
        .Add "ErrorCode", errorInfo.ErrorCode
        .Add "Severity", errorInfo.Severity
        .Add "CreatedAt", Now
        .Add "MaxLength", MAX_RECOVERY_CHAIN_LENGTH
    End With
    
    chain.Initialize metadata
    
    Set CreateRecoveryChainInternal = chain
    Exit Function
    
ErrorHandler:
    LogError "CreateRecoveryChainInternal", "Error: " & Err.Description
    Set CreateRecoveryChainInternal = Nothing
End Function

Private Function ExecuteRecoveryChain(ByVal chain As RecoveryChain, _
                                   ByRef errorInfo As ErrorInfo, _
                                   ByVal errorManager As clsErrorManager) As Boolean
    On Error GoTo ErrorHandler
    
    Dim strategy As IErrorStrategy
    Dim attempt As Long
    Dim success As Boolean
    
    For attempt = 1 To mMaxRetryCount
        ' キャンセル要求のチェック
        If mIsCancellationRequested Then
            LogError "ExecuteRecoveryChain", "Recovery cancelled by request"
            Exit Function
        End If
        
        ' タイムアウトチェック
        If DateDiff("s", chain.Metadata("CreatedAt"), Now) > RECOVERY_TIMEOUT_MS / 1000 Then
            LogError "ExecuteRecoveryChain", "Recovery timeout"
            Exit Function
        End If
        
        ' 進捗の更新
        mProgress.CurrentStep = attempt
        mProgress.CompletedSteps = attempt - 1
        
        For Each strategy In chain.Strategies
            ' 戦略の実行
            RaiseEvent StrategyExecutionStarted(strategy.GetDescription(), strategy.Priority)
            
            If strategy.RecoverFromError(errorInfo, errorManager) Then
                success = True
                chain.IncrementSuccess
                RaiseEvent StrategyExecutionCompleted(strategy.GetDescription(), True)
                ExecuteRecoveryChain = True
                Exit Function
            End If
            
            chain.IncrementFailure
            RaiseEvent StrategyExecutionCompleted(strategy.GetDescription(), False)
            
            ' フォールバック戦略の実行
            If Not strategy.FallbackStrategy Is Nothing Then
                RaiseEvent FallbackStrategyActivated(errorInfo, strategy.FallbackStrategy, frStrategyFailed)
                If strategy.FallbackStrategy.RecoverFromError(errorInfo, errorManager) Then
                    success = True
                    chain.IncrementSuccess
                    ExecuteRecoveryChain = True
                    Exit Function
                End If
            End If
        Next strategy
        
        ' リトライ間隔の待機
        Sleep mRetryInterval
    Next attempt
    
    ExecuteRecoveryChain = success
    Exit Function
    
ErrorHandler:
    LogError "ExecuteRecoveryChain", "Error: " & Err.Description
    ExecuteRecoveryChain = False
End Function

Private Sub LogError(ByVal source As String, ByVal message As String, _
                    Optional ByVal details As String = "")
    On Error Resume Next
    
    Dim logFile As Object
    Set logFile = mLogger.OpenTextFile(LOG_FILE_PATH, 8, True) ' 8 = ForAppending
    
    Dim logEntry As String
    logEntry = Now & "," & source & "," & Replace(message, ",", ";") & "," & details
    
    logFile.WriteLine logEntry
    logFile.Close
End Sub

Private Sub SaveMetrics()
    ' メトリクスの永続化処理
    ' TODO: メトリクスの保存実装
End Sub

Private Sub CloseLogFile()
    ' ログファイルのクローズ処理
    ' TODO: ログファイルのクローズ実装
End Sub

Private Sub Sleep(ByVal milliseconds As Long)
    ' スリープ処理の実装
    Application.Wait Now + TimeSerial(0, 0, milliseconds / 1000)
End Sub