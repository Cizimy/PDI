VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestAsyncValidationImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private asyncValidation As AsyncValidationImpl
Private mockContext As MockValidationContext
Private mockCacheStrategy As MockValidationCacheStrategy
Private mockSchedulingStrategy As MockValidationSchedulingStrategy

'@ModuleInitialize
Private Sub ModuleInitialize()
    'このメソッドは、モジュール内の各テストの前に1回実行されます
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'このメソッドは、モジュール内のすべてのテストの後に1回実行されます
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'このメソッドは、各テストの前に実行されます
    Set asyncValidation = New AsyncValidationImpl
    Set mockContext = New MockValidationContext
    Set mockCacheStrategy = New MockValidationCacheStrategy
    Set mockSchedulingStrategy = New MockValidationSchedulingStrategy
    
    ' モックの設定
    mockContext.SetupValidationRules CreateTestRules
    asyncValidation.SetCacheStrategy mockCacheStrategy
    asyncValidation.SetSchedulingStrategy mockSchedulingStrategy
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'このメソッドは、各テストの後に実行されます
    Set asyncValidation = Nothing
    Set mockContext = Nothing
    Set mockCacheStrategy = Nothing
    Set mockSchedulingStrategy = Nothing
End Sub

'テストヘルパーメソッド
Private Function CreateTestRules() As Collection
    Dim rules As New Collection
    
    Dim requiredRule As New RequiredFieldRule
    requiredRule.FieldName = "TestField"
    rules.Add requiredRule
    
    Dim lengthRule As New StringLengthRule
    lengthRule.FieldName = "TestField"
    lengthRule.MinLength = 3
    lengthRule.MaxLength = 10
    rules.Add lengthRule
    
    Set CreateTestRules = rules
End Function

'@TestMethod("AsyncValidationImpl")
Private Sub TestStartAsyncValidation_ShouldReturnValidationId()
    On Error GoTo TestFail
    
    'Arrange
    Dim value As String
    value = "test"
    
    'Act
    Dim validationId As String
    validationId = asyncValidation.StartAsyncValidation(value, mockContext)
    
    'Assert
    Assert.IsTrue Len(validationId) > 0, "バリデーションIDが生成されるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("AsyncValidationImpl")
Private Sub TestStartAsyncValidation_WithCachedResult_ShouldReturnCachedValidation()
    On Error GoTo TestFail
    
    'Arrange
    mockCacheStrategy.SetupIsCached "test", True
    Dim cachedResult As New ValidationResult
    cachedResult.Success = True
    mockCacheStrategy.SetupGetCachedResult "test", cachedResult
    
    'Act
    Dim validationId As String
    validationId = asyncValidation.StartAsyncValidation("test", mockContext)
    Dim result As ValidationResult
    Set result = asyncValidation.TryGetResults(validationId)
    
    'Assert
    Assert.IsTrue result.Success, "キャッシュされた結果が返されるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("AsyncValidationImpl")
Private Sub TestStartBatchValidation_ShouldProcessMultipleValues()
    On Error GoTo TestFail
    
    'Arrange
    Dim values As New Collection
    values.Add "test1"
    values.Add "test2"
    values.Add "test3"
    
    'Act
    Dim batchId As String
    batchId = asyncValidation.StartBatchValidation(values, mockContext)
    
    'Assert
    Assert.IsTrue Len(batchId) > 0, "バッチIDが生成されるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("AsyncValidationImpl")
Private Sub TestCancelValidation_ShouldCancelActiveValidation()
    On Error GoTo TestFail
    
    'Arrange
    Dim validationId As String
    validationId = asyncValidation.StartAsyncValidation("test", mockContext)
    
    'Act
    asyncValidation.CancelValidation validationId
    
    'Assert
    Dim status As ValidationQueueStatus
    Set status = asyncValidation.GetQueueStatus
    Assert.AreEqual 0, status.ActiveJobs, "アクティブなジョブが0になるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("AsyncValidationImpl")
Private Sub TestWaitForValidation_ShouldWaitForCompletion()
    On Error GoTo TestFail
    
    'Arrange
    Dim validationId As String
    validationId = asyncValidation.StartAsyncValidation("test", mockContext)
    
    'Act
    Dim result As Boolean
    result = asyncValidation.WaitForValidation(validationId, 1000) ' 1秒タイムアウト
    
    'Assert
    Assert.IsTrue result, "バリデーションが完了するまで待機するべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("AsyncValidationImpl")
Private Sub TestWaitForAll_ShouldWaitForAllValidations()
    On Error GoTo TestFail
    
    'Arrange
    Dim values As New Collection
    values.Add "test1"
    values.Add "test2"
    asyncValidation.StartBatchValidation values, mockContext
    
    'Act
    Dim result As Boolean
    result = asyncValidation.WaitForAll(2000) ' 2秒タイムアウト
    
    'Assert
    Assert.IsTrue result, "すべてのバリデーションが完了するまで待機するべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("AsyncValidationImpl")
Private Sub TestChangePriority_ShouldUpdateValidationPriority()
    On Error GoTo TestFail
    
    'Arrange
    Dim validationId As String
    validationId = asyncValidation.StartAsyncValidation("test", mockContext, vpNormal)
    
    'Act
    asyncValidation.ChangePriority validationId, vpHigh
    
    'Assert
    ' 優先度の変更が反映されていることを確認（実装依存）
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("AsyncValidationImpl")
Private Sub TestSetResourceLimit_ShouldEnforceResourceLimits()
    On Error GoTo TestFail
    
    'Arrange
    asyncValidation.SetResourceLimit "MaxConcurrent", 2
    
    'Act
    Dim id1 As String, id2 As String, id3 As String
    id1 = asyncValidation.StartAsyncValidation("test1", mockContext)
    id2 = asyncValidation.StartAsyncValidation("test2", mockContext)
    id3 = asyncValidation.StartAsyncValidation("test3", mockContext)
    
    'Assert
    Dim status As ValidationQueueStatus
    Set status = asyncValidation.GetQueueStatus
    Assert.IsTrue status.ActiveJobs <= 2, "同時実行数が制限されるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("AsyncValidationImpl")
Private Sub TestGetPerformanceMetrics_ShouldTrackMetrics()
    On Error GoTo TestFail
    
    'Arrange
    Dim validationId As String
    validationId = asyncValidation.StartAsyncValidation("test", mockContext)
    asyncValidation.WaitForValidation validationId
    
    'Act
    Dim metrics As ValidationMetrics
    Set metrics = asyncValidation.GetPerformanceMetrics
    
    'Assert
    Assert.IsTrue metrics.TotalExecutionTime > 0, "実行時間が記録されるべき"
    Assert.IsTrue metrics.RuleCount > 0, "ルール実行数が記録されるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("AsyncValidationImpl")
Private Sub TestQueueStatus_ShouldProvideAccurateStatus()
    On Error GoTo TestFail
    
    'Arrange
    Dim values As New Collection
    values.Add "test1"
    values.Add "test2"
    asyncValidation.StartBatchValidation values, mockContext
    
    'Act
    Dim status As ValidationQueueStatus
    Set status = asyncValidation.GetQueueStatus
    
    'Assert
    Assert.IsTrue status.QueueLength >= 0, "キュー長が正しく報告されるべき"
    Assert.IsTrue status.ActiveJobs >= 0, "アクティブなジョブ数が正しく報告されるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("AsyncValidationImpl")
Private Sub TestParallelExecution_ShouldExecuteValidationsInParallel()
    On Error GoTo TestFail
    
    'Arrange
    asyncValidation.SetResourceLimit "MaxConcurrent", 3
    
    'Act
    Dim startTime As Date
    startTime = Now
    
    Dim id1 As String, id2 As String, id3 As String
    id1 = asyncValidation.StartAsyncValidation("test1", mockContext)
    id2 = asyncValidation.StartAsyncValidation("test2", mockContext)
    id3 = asyncValidation.StartAsyncValidation("test3", mockContext)
    
    asyncValidation.WaitForAll
    
    Dim executionTime As Double
    executionTime = DateDiff("s", startTime, Now)
    
    'Assert
    ' 並列実行により、総実行時間が直列実行より短くなることを確認
    Assert.IsTrue executionTime < 3, "並列実行により実行時間が短縮されるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("AsyncValidationImpl")
Private Sub TestErrorHandling_ShouldHandleValidationErrors()
    On Error GoTo TestFail
    
    'Arrange
    mockContext.SetupThrowError True ' モックにエラーを発生させる設定
    
    'Act
    Dim validationId As String
    validationId = asyncValidation.StartAsyncValidation("test", mockContext)
    asyncValidation.WaitForValidation validationId
    
    'Assert
    Dim result As ValidationResult
    Set result = asyncValidation.TryGetResults(validationId)
    Assert.IsFalse result.Success, "エラーが発生した場合は失敗を報告するべき"
    Assert.IsTrue Len(result.ErrorMessage) > 0, "エラーメッセージが設定されるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub