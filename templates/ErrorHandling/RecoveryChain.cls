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

' イベント
Public Event StrategyExecuted(ByVal strategyId As String, ByVal success As Boolean)
Public Event ChainCompleted(ByVal chainId As String, ByVal successCount As Long, ByVal failureCount As Long)

Private Sub Class_Initialize()
    Set mStrategies = New Collection
    Set mStrategyInfo = CreateObject("Scripting.Dictionary")
    Set mMetadata = New Dictionary
    mSuccessCount = 0
    mFailureCount = 0
    mIsInitialized = False
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
    
    mStrategies.Add strategy
    mStrategyInfo.Add strategy.GetDescription(), info
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
        If Not strategy.Validate Then
            Err.Raise vbObjectError + 3, "RecoveryChain.Validate", _
                     "Strategy validation failed: " & strategy.GetDescription()
        End If
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
End Sub

Public Sub RecordError(ByVal errorDescription As String)
    mMetadata.Add "LastError", errorDescription
    mMetadata.Add "ErrorTimestamp", Now
End Sub