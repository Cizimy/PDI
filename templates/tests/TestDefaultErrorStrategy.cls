VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestDefaultErrorStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private strategy As clsDefaultErrorStrategy
Private errorManager As clsErrorManager
Private mockErrorInfo As ErrorInfo

'@ModuleInitialize
Private Sub ModuleInitialize()
    'このメソッドは、モジュール内の各テストの前に1回実行されます
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    
    'テスト対象のオブジェクトを初期化
    Set strategy = New clsDefaultErrorStrategy
    Set errorManager = New clsErrorManager
    Set mockErrorInfo = New ErrorInfo
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'このメソッドは、モジュール内のすべてのテストの後に1回実行されます
    Set Assert = Nothing
    Set Fakes = Nothing
    Set strategy = Nothing
    Set errorManager = Nothing
    Set mockErrorInfo = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'このメソッドは、各テストの前に実行されます
    InitializeMockErrorInfo
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'このメソッドは、各テストの後に実行されます
End Sub

Private Sub InitializeMockErrorInfo()
    With mockErrorInfo
        .ErrorCode = ecDatabaseError
        .Message = "Test Error Message"
        .Severity = esError
        .SourceModule = "TestModule"
        .SourceProcedure = "TestProcedure"
        .Timestamp = Now
    End With
End Sub

'@TestMethod("RecoverFromError")
Private Sub TestRecoverFromError_BasicRecovery()
    On Error GoTo TestFail
    
    'Arrange
    Dim result As Boolean
    
    'Act
    result = strategy.RecoverFromError(mockErrorInfo, errorManager)
    
    'Assert
    Assert.IsTrue result, "Basic recovery should succeed"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("RetryMechanism")
Private Sub TestRetryMechanism()
    On Error GoTo TestFail
    
    'Arrange
    strategy.MaxRetryCount = 3
    strategy.RetryInterval = 100 ' 100ms
    
    'Act
    Dim result As Boolean
    result = strategy.RecoverFromError(mockErrorInfo, errorManager)
    
    'Assert
    Assert.IsTrue result, "Retry mechanism should eventually succeed"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("FallbackStrategy")
Private Sub TestFallbackStrategy()
    On Error GoTo TestFail
    
    'Arrange
    Dim fallback As New clsDefaultErrorStrategy
    Set strategy.FallbackStrategy = fallback
    
    'Act
    Dim result As Boolean
    result = strategy.RecoverFromError(mockErrorInfo, errorManager)
    
    'Assert
    Assert.IsTrue result, "Fallback strategy should be executed"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("MetricsCollection")
Private Sub TestMetricsCollection()
    On Error GoTo TestFail
    
    'Arrange
    strategy.RecoverFromError mockErrorInfo, errorManager
    
    'Act
    Dim metrics As Collection
    Set metrics = strategy.GetPerformanceMetrics
    
    'Assert
    Assert.IsNotNothing metrics, "Metrics collection should not be nothing"
    Assert.IsTrue metrics.Count > 0, "Metrics should contain at least one entry"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("EventNotification")
Private Sub TestEventNotification()
    On Error GoTo TestFail
    
    'Arrange
    Dim eventRaised As Boolean
    eventRaised = False
    
    'イベントハンドラの設定
    Dim WithEvents testStrategy As clsDefaultErrorStrategy
    Set testStrategy = strategy
    
    'Act
    testStrategy.RecoverFromError mockErrorInfo, errorManager
    
    'Assert
    Assert.IsTrue eventRaised, "Event should be raised during recovery"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ValidationChecks")
Private Sub TestValidation()
    On Error GoTo TestFail
    
    'Arrange
    Dim isValid As Boolean
    
    'Act
    isValid = strategy.Validate()
    
    'Assert
    Assert.IsTrue isValid, "Strategy should be valid after initialization"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("RecoveryChainCreation")
Private Sub TestRecoveryChainCreation()
    On Error GoTo TestFail
    
    'Arrange
    Dim strategies As New Collection
    strategies.Add New clsDefaultErrorStrategy
    strategies.Add New clsDefaultErrorStrategy
    
    'Act
    Dim chain As RecoveryChain
    Set chain = strategy.CreateRecoveryChain(strategies, errorManager)
    
    'Assert
    Assert.IsNotNothing chain, "Recovery chain should be created"
    Assert.AreEqual 2, chain.StrategyCount, "Chain should contain 2 strategies"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("ProgressTracking")
Private Sub TestProgressTracking()
    On Error GoTo TestFail
    
    'Arrange
    strategy.RecoverFromError mockErrorInfo, errorManager
    
    'Act
    Dim progress As RecoveryProgress
    Set progress = strategy.GetProgress(True)
    
    'Assert
    Assert.IsNotNothing progress, "Progress object should not be nothing"
    Assert.IsTrue progress.CompletedSteps >= 0, "Completed steps should be tracked"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("CancellationRequest")
Private Sub TestCancellation()
    On Error GoTo TestFail
    
    'Arrange
    strategy.CancelRecovery
    
    'Act
    Dim result As Boolean
    result = strategy.RecoverFromError(mockErrorInfo, errorManager)
    
    'Assert
    Assert.IsFalse result, "Recovery should be cancelled"
    Assert.IsTrue strategy.IsCancellationRequested, "Cancellation flag should be set"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'イベントハンドラ
Private Sub testStrategy_RecoveryChainStarted(ByVal chainId As String, ByVal strategies As Collection)
    eventRaised = True
End Sub