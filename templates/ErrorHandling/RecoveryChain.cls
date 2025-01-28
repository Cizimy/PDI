VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecoveryChain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

' プライベート変数
Private mChainId As String
Private mStrategies As Collection
Private mStrategyInfo As Object ' Dictionary
Private mSuccessCount As Long
Private mFailureCount As Long
Private mMetadata As Dictionary
Private mIsInitialized As Boolean
Private mInTransaction As Boolean
Private mTransactionLevel As Long
Private mExecutionLog As Collection
Private mPerformanceMetrics As Dictionary

' 定数
Private Const MAX_RETRY_ATTEMPTS As Long = 3
Private Const TRANSACTION_TIMEOUT_MS As Long = 30000

' イベント
Public Event StrategyExecuted(ByVal strategyId As String, ByVal success As Boolean)
Public Event ChainCompleted(ByVal chainId As String, ByVal successCount As Long, ByVal failureCount As Long)
Public Event TransactionStarted(ByVal chainId As String, ByVal level As Long)
Public Event TransactionCommitted(ByVal chainId As String, ByVal level As Long)
Public Event TransactionRolledBack(ByVal chainId As String, ByVal level As Long, ByVal reason As String)
Public Event PerformanceThresholdExceeded(ByVal metricName As String, ByVal currentValue As Double, ByVal threshold As Double)

Private Sub Class_Initialize()
    Set mStrategies = New Collection
    Set mStrategyInfo = CreateObject("Scripting.Dictionary")
    Set mMetadata = New Dictionary
    Set mExecutionLog = New Collection
    Set mPerformanceMetrics = CreateObject("Scripting.Dictionary")
    
    mSuccessCount = 0
    mFailureCount = 0
    mIsInitialized = False
    mInTransaction = False
    mTransactionLevel = 0
    
    InitializePerformanceMetrics
End Sub

Private Sub InitializePerformanceMetrics()
    With mPerformanceMetrics
        .Add "TotalExecutionTime", 0
        .Add "AverageStrategyTime", 0
        .Add "MaxStrategyTime", 0
        .Add "TotalTransactions", 0
        .Add "SuccessfulTransactions", 0
        .Add "RolledBackTransactions", 0
    End With
End Sub

Public Sub Initialize(ByVal metadata As Dictionary)
    Set mMetadata = metadata
    mChainId = "CHAIN_" & Format(Now, "yyyymmddhhnnss") & "_" & CStr(Int(Rnd * 1000))
    mIsInitialized = True
End Sub

Public Property Get ChainId() As String
    ChainId = mChainId
End Property

Public Property Get Strategies() As Collection
    Set Strategies = mStrategies
End Property

Public Property Get StrategyCount() As Long
    StrategyCount = mStrategies.Count
End Property

Public Property Get SuccessCount() As Long
    SuccessCount = mSuccessCount
End Property

Public Property Get FailureCount() As Long
    FailureCount = mFailureCount
End Property

Public Sub AddStrategy(ByVal strategy As IErrorStrategy, ByVal info As Dictionary)
    If Not mIsInitialized Then
        Err.Raise vbObjectError + 1, "RecoveryChain.AddStrategy", "Chain not initialized"
    End If
    
    ValidateStrategy strategy
    
    mStrategies.Add strategy
    mStrategyInfo.Add strategy.GetDescription(), info
End Sub

Private Sub ValidateStrategy(ByVal strategy As IErrorStrategy)
    If strategy Is Nothing Then
        Err.Raise vbObjectError + 2, "RecoveryChain.ValidateStrategy", "Strategy cannot be null"
    End If
    
    If Not strategy.Validate Then
        Err.Raise vbObjectError + 3, "RecoveryChain.ValidateStrategy", _
                 "Strategy validation failed: " & strategy.GetDescription()
    End If
End Sub

Public Function GetStrategyInfo(ByVal strategy As IErrorStrategy) As Dictionary
    If mStrategyInfo.Exists(strategy.GetDescription()) Then
        Set GetStrategyInfo = mStrategyInfo(strategy.GetDescription())
    Else
        Set GetStrategyInfo = Nothing
    End If
End Function

Public Sub SortByPriority()
    Dim i As Long, j As Long
    Dim tempStrategy As IErrorStrategy
    Dim tempInfo As Dictionary
    
    For i = 1 To mStrategies.Count - 1
        For j = i + 1 To mStrategies.Count
            If mStrategies(i).Priority < mStrategies(j).Priority Then
                ' 戦略の入れ替え
                Set tempStrategy = mStrategies(i)
                Set mStrategies(i) = mStrategies(j)
                Set mStrategies(j) = tempStrategy
                
                ' 情報の入れ替え
                Set tempInfo = mStrategyInfo(tempStrategy.GetDescription())
                mStrategyInfo.Remove tempStrategy.GetDescription()
                mStrategyInfo.Add tempStrategy.GetDescription(), tempInfo
            End If
        Next j
    Next i
End Sub

Public Function ExecuteWithTransaction(ByVal strategy As IErrorStrategy, ByVal errorInfo As ErrorInfo) As Boolean
    On Error GoTo ErrorHandler
    
    BeginTransaction
    
    Dim startTime As Date
    startTime = Now
    
    If strategy.RecoverFromError(errorInfo) Then
        CommitTransaction
        
        ' パフォーマンス計測
        UpdatePerformanceMetrics strategy, startTime
        
        ExecuteWithTransaction = True
        Exit Function
    End If
    
ErrorHandler:
    RollbackTransaction "Strategy execution failed: " & Err.Description
    ExecuteWithTransaction = False
End Function

Private Sub BeginTransaction()
    mTransactionLevel = mTransactionLevel + 1
    mInTransaction = True
    
    RaiseEvent TransactionStarted(mChainId, mTransactionLevel)
    
    mPerformanceMetrics("TotalTransactions") = mPerformanceMetrics("TotalTransactions") + 1
End Sub

Private Sub CommitTransaction()
    If Not mInTransaction Then Exit Sub
    
    RaiseEvent TransactionCommitted(mChainId, mTransactionLevel)
    
    mTransactionLevel = mTransactionLevel - 1
    If mTransactionLevel = 0 Then
        mInTransaction = False
    End If
    
    mPerformanceMetrics("SuccessfulTransactions") = mPerformanceMetrics("SuccessfulTransactions") + 1
End Sub

Private Sub RollbackTransaction(ByVal reason As String)
    If Not mInTransaction Then Exit Sub
    
    RaiseEvent TransactionRolledBack(mChainId, mTransactionLevel, reason)
    
    mTransactionLevel = 0
    mInTransaction = False
    
    mPerformanceMetrics("RolledBackTransactions") = mPerformanceMetrics("RolledBackTransactions") + 1
End Sub

Private Sub UpdatePerformanceMetrics(ByVal strategy As IErrorStrategy, ByVal startTime As Date)
    Dim executionTime As Double
    executionTime = DateDiff("s", startTime, Now)
    
    With mPerformanceMetrics
        .Item("TotalExecutionTime") = .Item("TotalExecutionTime") + executionTime
        
        If executionTime > .Item("MaxStrategyTime") Then
            .Item("MaxStrategyTime") = executionTime
        End If
        
        ' 平均実行時間の更新
        If mSuccessCount > 0 Then
            .Item("AverageStrategyTime") = .Item("TotalExecutionTime") / mSuccessCount
        End If
    End With
    
    ' パフォーマンスしきい値チェック
    CheckPerformanceThresholds
End Sub

Private Sub CheckPerformanceThresholds()
    ' 実行時間のしきい値チェック
    If mPerformanceMetrics("AverageStrategyTime") > 5 Then ' 5秒を超える場合
        RaiseEvent PerformanceThresholdExceeded("AverageStrategyTime", _
                                              mPerformanceMetrics("AverageStrategyTime"), 5)
    End If
    
    ' ロールバック率のチェック
    Dim rollbackRate As Double
    If mPerformanceMetrics("TotalTransactions") > 0 Then
        rollbackRate = mPerformanceMetrics("RolledBackTransactions") / mPerformanceMetrics("TotalTransactions") * 100
        If rollbackRate > 20 Then ' 20%を超える場合
            RaiseEvent PerformanceThresholdExceeded("RollbackRate", rollbackRate, 20)
        End If
    End If
End Sub

Public Function Validate() As Boolean
    If Not mIsInitialized Then
        Err.Raise vbObjectError + 1, "RecoveryChain.Validate", "Chain not initialized"
    End If
    
    If mStrategies.Count = 0 Then
        Err.Raise vbObjectError + 2, "RecoveryChain.Validate", "No strategies defined"
    End If
    
    ' 各戦略の検証
    Dim strategy As IErrorStrategy
    For Each strategy In mStrategies
        ValidateStrategy strategy
    Next strategy
    
    Validate = True
End Function

Public Sub IncrementSuccess()
    mSuccessCount = mSuccessCount + 1
    RaiseEvent StrategyExecuted(mChainId, True)
    
    If mSuccessCount + mFailureCount = mStrategies.Count Then
        RaiseEvent ChainCompleted(mChainId, mSuccessCount, mFailureCount)
    End If
End Sub

Public Sub IncrementFailure()
    mFailureCount = mFailureCount + 1
    RaiseEvent StrategyExecuted(mChainId, False)
    
    If mSuccessCount + mFailureCount = mStrategies.Count Then
        RaiseEvent ChainCompleted(mChainId, mSuccessCount, mFailureCount)
    End If
End Sub

Public Function CalculateProgress() As Long
    If mStrategies.Count = 0 Then
        CalculateProgress = 0
        Exit Function
    End If
    
    CalculateProgress = ((mSuccessCount + mFailureCount) / mStrategies.Count) * 100
End Function

Public Sub RecordExecution(ByVal context As Dictionary)
    ' 実行結果の記録
    mMetadata.Add "ExecutionContext", context
    mMetadata.Add "CompletedAt", Now
    mMetadata.Add "TotalStrategies", mStrategies.Count
    mMetadata.Add "SuccessCount", mSuccessCount
    mMetadata.Add "FailureCount", mFailureCount
    
    ' パフォーマンス情報の記録
    mMetadata.Add "PerformanceMetrics", mPerformanceMetrics
    
    ' 実行ログに追加
    mExecutionLog.Add context
End Sub

Public Sub RecordError(ByVal errorDescription As String)
    mMetadata.Add "LastError", errorDescription
    mMetadata.Add "ErrorTimestamp", Now
    
    ' エラーログに追加
    Dim errorLog As Dictionary
    Set errorLog = New Dictionary
    errorLog.Add "Description", errorDescription
    errorLog.Add "Timestamp", Now
    errorLog.Add "ChainState", GetChainState()
    
    mExecutionLog.Add errorLog
End Sub

Private Function GetChainState() As Dictionary
    Dim state As New Dictionary
    
    With state
        .Add "InTransaction", mInTransaction
        .Add "TransactionLevel", mTransactionLevel
        .Add "SuccessCount", mSuccessCount
        .Add "FailureCount", mFailureCount
        .Add "Progress", CalculateProgress()
    End With
    
    Set GetChainState = state
End Function

Public Function GetPerformanceMetrics() As Dictionary
    Set GetPerformanceMetrics = mPerformanceMetrics
End Function

Public Function GetExecutionLog() As Collection
    Set GetExecutionLog = mExecutionLog
End Function