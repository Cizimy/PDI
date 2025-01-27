VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestNumericRangeRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private rule As NumericRangeRule

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
    Set rule = New NumericRangeRule
    rule.FieldName = "TestField"
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'このメソッドは、各テストの後に実行されます
    Set rule = Nothing
End Sub

'@TestMethod("NumericRangeRule")
Private Sub TestValidate_WhenValueIsNotNumeric_ShouldReturnFalse()
    On Error GoTo TestFail
    
    'Arrange
    Dim value As String
    value = "abc"
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsFalse result, "数値以外の値は無効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) > 0, "エラーメッセージが設定されているべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("NumericRangeRule")
Private Sub TestValidate_WhenValueIsWithinRange_ShouldReturnTrue()
    On Error GoTo TestFail
    
    'Arrange
    rule.MinValue = 1
    rule.MaxValue = 10
    Dim value As Double
    value = 5
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsTrue result, "範囲内の値は有効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) = 0, "エラーメッセージは設定されていないべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("NumericRangeRule")
Private Sub TestValidate_WhenValueIsTooSmall_ShouldReturnFalse()
    On Error GoTo TestFail
    
    'Arrange
    rule.MinValue = 1
    rule.IncludeMin = True
    Dim value As Double
    value = 0
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsFalse result, "最小値未満は無効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) > 0, "エラーメッセージが設定されているべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("NumericRangeRule")
Private Sub TestValidate_WhenValueIsTooLarge_ShouldReturnFalse()
    On Error GoTo TestFail
    
    'Arrange
    rule.MaxValue = 10
    rule.IncludeMax = True
    Dim value As Double
    value = 11
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsFalse result, "最大値超過は無効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) > 0, "エラーメッセージが設定されているべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("NumericRangeRule")
Private Sub TestValidate_WhenValueEqualsMinAndIncludeMinIsTrue_ShouldReturnTrue()
    On Error GoTo TestFail
    
    'Arrange
    rule.MinValue = 1
    rule.IncludeMin = True
    Dim value As Double
    value = 1
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsTrue result, "最小値と等しい値は有効であるべき（IncludeMin = True）"
    Assert.IsTrue Len(rule.ErrorMessage) = 0, "エラーメッセージは設定されていないべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("NumericRangeRule")
Private Sub TestValidate_WhenValueEqualsMinAndIncludeMinIsFalse_ShouldReturnFalse()
    On Error GoTo TestFail
    
    'Arrange
    rule.MinValue = 1
    rule.IncludeMin = False
    Dim value As Double
    value = 1
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsFalse result, "最小値と等しい値は無効であるべき（IncludeMin = False）"
    Assert.IsTrue Len(rule.ErrorMessage) > 0, "エラーメッセージが設定されているべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("NumericRangeRule")
Private Sub TestValidate_WhenValueEqualsMaxAndIncludeMaxIsTrue_ShouldReturnTrue()
    On Error GoTo TestFail
    
    'Arrange
    rule.MaxValue = 10
    rule.IncludeMax = True
    Dim value As Double
    value = 10
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsTrue result, "最大値と等しい値は有効であるべき（IncludeMax = True）"
    Assert.IsTrue Len(rule.ErrorMessage) = 0, "エラーメッセージは設定されていないべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("NumericRangeRule")
Private Sub TestValidate_WhenValueEqualsMaxAndIncludeMaxIsFalse_ShouldReturnFalse()
    On Error GoTo TestFail
    
    'Arrange
    rule.MaxValue = 10
    rule.IncludeMax = False
    Dim value As Double
    value = 10
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsFalse result, "最大値と等しい値は無効であるべき（IncludeMax = False）"
    Assert.IsTrue Len(rule.ErrorMessage) > 0, "エラーメッセージが設定されているべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("NumericRangeRule")
Private Sub TestValidate_WhenValueIsNullAndAllowNullIsTrue_ShouldReturnTrue()
    On Error GoTo TestFail
    
    'Arrange
    rule.AllowNull = True
    Dim value As Variant
    value = Null
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsTrue result, "Null値は有効であるべき（AllowNull = True）"
    Assert.IsTrue Len(rule.ErrorMessage) = 0, "エラーメッセージは設定されていないべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("NumericRangeRule")
Private Sub TestValidate_WhenValueHasMoreDecimalPlacesThanPrecision_ShouldReturnFalse()
    On Error GoTo TestFail
    
    'Arrange
    rule.Precision = 2
    Dim value As Double
    value = 1.234
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsFalse result, "指定された精度を超える値は無効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) > 0, "エラーメッセージが設定されているべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("NumericRangeRule")
Private Sub TestValidate_WhenValueHasExactPrecision_ShouldReturnTrue()
    On Error GoTo TestFail
    
    'Arrange
    rule.Precision = 2
    Dim value As Double
    value = 1.23
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsTrue result, "指定された精度と等しい値は有効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) = 0, "エラーメッセージは設定されていないべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("NumericRangeRule")
Private Sub TestErrorMessage_WhenFieldNameIsSet_ShouldIncludeFieldName()
    On Error GoTo TestFail
    
    'Arrange
    rule.FieldName = "テストフィールド"
    rule.MinValue = 1
    Dim value As Double
    value = 0
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsTrue InStr(rule.ErrorMessage, "テストフィールド") > 0, "エラーメッセージにフィールド名が含まれるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("NumericRangeRule")
Private Sub TestValidate_WhenRuleIsDisabled_ShouldReturnTrue()
    On Error GoTo TestFail
    
    'Arrange
    rule.IsEnabled = False
    rule.MinValue = 1
    Dim value As Double
    value = 0
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsTrue result, "ルールが無効の場合、常にTrueを返すべき"
    Assert.IsTrue Len(rule.ErrorMessage) = 0, "エラーメッセージは設定されていないべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("NumericRangeRule")
Private Sub TestExecutionTime_ShouldBeGreaterThanZero()
    On Error GoTo TestFail
    
    'Arrange
    rule.MinValue = 1
    rule.MaxValue = 10
    Dim value As Double
    value = 5
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsTrue rule.ExecutionTime > 0, "実行時間は0より大きいべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub