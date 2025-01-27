VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCompositeValidationRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private composite As CompositeValidationRule
Private requiredRule As RequiredFieldRule
Private lengthRule As StringLengthRule
Private numericRule As NumericRangeRule
Private regexRule As RegexPatternRule

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
    Set composite = New CompositeValidationRule
    Set requiredRule = New RequiredFieldRule
    Set lengthRule = New StringLengthRule
    Set numericRule = New NumericRangeRule
    Set regexRule = New RegexPatternRule
    
    ' 基本的な設定
    requiredRule.FieldName = "TestField"
    lengthRule.FieldName = "TestField"
    lengthRule.MinLength = 3
    lengthRule.MaxLength = 10
    numericRule.FieldName = "TestField"
    numericRule.MinValue = 1
    numericRule.MaxValue = 100
    regexRule.FieldName = "TestField"
    regexRule.Pattern = "^[A-Za-z0-9]+$"
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'このメソッドは、各テストの後に実行されます
    Set composite = Nothing
    Set requiredRule = Nothing
    Set lengthRule = Nothing
    Set numericRule = Nothing
    Set regexRule = Nothing
End Sub

'@TestMethod("CompositeValidationRule")
Private Sub TestValidate_WhenNoRules_ShouldReturnTrue()
    On Error GoTo TestFail
    
    'Arrange
    Dim value As String
    value = "test"
    
    'Act
    Dim result As Boolean
    result = composite.Validate(value)
    
    'Assert
    Assert.IsTrue result, "ルールが設定されていない場合はTrueを返すべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("CompositeValidationRule")
Private Sub TestValidate_WhenAllRulesPass_ShouldReturnTrue()
    On Error GoTo TestFail
    
    'Arrange
    composite.AddRule requiredRule
    composite.AddRule lengthRule
    Dim value As String
    value = "test123"
    
    'Act
    Dim result As Boolean
    result = composite.Validate(value)
    
    'Assert
    Assert.IsTrue result, "すべてのルールが成功する場合はTrueを返すべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("CompositeValidationRule")
Private Sub TestValidate_WhenAnyRuleFails_ShouldReturnFalse()
    On Error GoTo TestFail
    
    'Arrange
    composite.AddRule requiredRule
    composite.AddRule lengthRule
    Dim value As String
    value = "ab" ' 最小長未満
    
    'Act
    Dim result As Boolean
    result = composite.Validate(value)
    
    'Assert
    Assert.IsFalse result, "いずれかのルールが失敗する場合はFalseを返すべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("CompositeValidationRule")
Private Sub TestValidate_WithDependencies_ShouldExecuteInCorrectOrder()
    On Error GoTo TestFail
    
    'Arrange
    lengthRule.AddDependency requiredRule
    composite.AddRule lengthRule
    composite.AddRule requiredRule
    
    Dim value As String
    value = "" ' requiredRuleが失敗するケース
    
    'Act
    Dim result As Boolean
    result = composite.Validate(value)
    
    'Assert
    Assert.IsFalse result, "依存関係のあるルールが正しい順序で実行されるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("CompositeValidationRule")
Private Sub TestValidate_WhenStopOnFirstFailureIsTrue_ShouldStopOnFirstFailure()
    On Error GoTo TestFail
    
    'Arrange
    composite.StopOnFirstFailure = True
    composite.AddRule requiredRule
    composite.AddRule lengthRule
    composite.AddRule regexRule
    
    Dim value As String
    value = "" ' 最初のルール（required）で失敗
    
    'Act
    Dim result As Boolean
    result = composite.Validate(value)
    
    'Assert
    Assert.IsFalse result, "最初の失敗で検証を停止するべき"
    Assert.AreEqual 1, composite.ErrorDetails.Count, "エラーは1つだけ記録されるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("CompositeValidationRule")
Private Sub TestValidate_WhenStopOnFirstFailureIsFalse_ShouldContinueAfterFailure()
    On Error GoTo TestFail
    
    'Arrange
    composite.StopOnFirstFailure = False
    composite.AddRule requiredRule
    composite.AddRule lengthRule
    composite.AddRule regexRule
    
    Dim value As String
    value = "" ' すべてのルールで失敗
    
    'Act
    Dim result As Boolean
    result = composite.Validate(value)
    
    'Assert
    Assert.IsFalse result, "すべてのルールを実行し、失敗を記録するべき"
    Assert.IsTrue composite.ErrorDetails.Count > 1, "複数のエラーが記録されるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("CompositeValidationRule")
Private Sub TestValidate_WithParallelExecution_ShouldExecuteRulesInParallel()
    On Error GoTo TestFail
    
    'Arrange
    composite.ParallelExecution = True
    composite.AddRule requiredRule
    composite.AddRule regexRule
    
    Dim value As String
    value = "test123"
    
    'Act
    Dim result As Boolean
    result = composite.Validate(value)
    
    'Assert
    Assert.IsTrue result, "並列実行でも正しく検証されるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("CompositeValidationRule")
Private Sub TestValidationGroup_ShouldExecuteGroupRulesCorrectly()
    On Error GoTo TestFail
    
    'Arrange
    Dim groupRules As New Collection
    groupRules.Add requiredRule
    groupRules.Add lengthRule
    composite.CreateValidationGroup "TestGroup", groupRules
    
    Dim value As String
    value = "test123"
    
    'Act
    Dim result As Boolean
    result = composite.Validate(value)
    
    'Assert
    Assert.IsTrue result, "グループ内のルールが正しく実行されるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("CompositeValidationRule")
Private Sub TestExecutionPlan_ShouldOptimizeRuleExecution()
    On Error GoTo TestFail
    
    'Arrange
    composite.AddRule requiredRule
    composite.AddRule lengthRule
    composite.AddRule regexRule
    
    'Act
    Dim plan As ValidationExecutionPlan
    Set plan = composite.ExecutionPlan
    
    'Assert
    Assert.IsTrue plan.ExecutionOrder.Count = 3, "実行計画には3つのルールが含まれるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("CompositeValidationRule")
Private Sub TestErrorDetails_ShouldProvideDetailedErrorInformation()
    On Error GoTo TestFail
    
    'Arrange
    composite.StopOnFirstFailure = False
    composite.AddRule requiredRule
    composite.AddRule lengthRule
    
    Dim value As String
    value = "" ' 両方のルールで失敗
    
    'Act
    Dim result As Boolean
    result = composite.Validate(value)
    
    'Assert
    Assert.IsFalse result, "検証は失敗するべき"
    Assert.IsTrue composite.ErrorDetails.Count = 2, "両方のルールのエラーが記録されるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("CompositeValidationRule")
Private Sub TestPerformanceMetrics_ShouldTrackExecutionTime()
    On Error GoTo TestFail
    
    'Arrange
    composite.AddRule requiredRule
    composite.AddRule lengthRule
    
    Dim value As String
    value = "test123"
    
    'Act
    Dim result As Boolean
    result = composite.Validate(value)
    
    'Assert
    Assert.IsTrue composite.ExecutionTime > 0, "実行時間が記録されるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub