VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestRequiredFieldRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private rule As RequiredFieldRule

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
    Set rule = New RequiredFieldRule
    rule.FieldName = "TestField"
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'このメソッドは、各テストの後に実行されます
    Set rule = Nothing
End Sub

'@TestMethod("RequiredFieldRule")
Private Sub TestValidate_WhenValueIsNothing_ShouldReturnFalse()
    On Error GoTo TestFail
    
    'Arrange
    Dim value As Variant
    Set value = Nothing
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsFalse result, "Nothingの値は無効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) > 0, "エラーメッセージが設定されているべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("RequiredFieldRule")
Private Sub TestValidate_WhenValueIsNull_ShouldReturnFalse()
    On Error GoTo TestFail
    
    'Arrange
    Dim value As Variant
    value = Null
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsFalse result, "Nullの値は無効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) > 0, "エラーメッセージが設定されているべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("RequiredFieldRule")
Private Sub TestValidate_WhenValueIsEmptyString_ShouldReturnFalse()
    On Error GoTo TestFail
    
    'Arrange
    Dim value As String
    value = ""
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsFalse result, "空文字列は無効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) > 0, "エラーメッセージが設定されているべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("RequiredFieldRule")
Private Sub TestValidate_WhenValueIsWhitespace_ShouldReturnFalse()
    On Error GoTo TestFail
    
    'Arrange
    Dim value As String
    value = "   "
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsFalse result, "空白文字のみは無効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) > 0, "エラーメッセージが設定されているべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("RequiredFieldRule")
Private Sub TestValidate_WhenValueIsValid_ShouldReturnTrue()
    On Error GoTo TestFail
    
    'Arrange
    Dim value As String
    value = "有効な値"
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsTrue result, "有効な値は正しく検証されるべき"
    Assert.IsTrue Len(rule.ErrorMessage) = 0, "エラーメッセージは設定されていないべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("RequiredFieldRule")
Private Sub TestValidate_WhenAllowEmptyStringIsTrue_ShouldReturnTrue()
    On Error GoTo TestFail
    
    'Arrange
    rule.AllowEmptyString = True
    Dim value As String
    value = ""
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsTrue result, "AllowEmptyStringがTrueの場合、空文字列は有効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) = 0, "エラーメッセージは設定されていないべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("RequiredFieldRule")
Private Sub TestValidate_WhenTrimValueIsFalse_ShouldNotTrimWhitespace()
    On Error GoTo TestFail
    
    'Arrange
    rule.TrimValue = False
    Dim value As String
    value = "  有効な値  "
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsTrue result, "TrimValueがFalseの場合、前後の空白は保持されるべき"
    Assert.IsTrue Len(rule.ErrorMessage) = 0, "エラーメッセージは設定されていないべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("RequiredFieldRule")
Private Sub TestErrorMessage_WhenFieldNameIsSet_ShouldIncludeFieldName()
    On Error GoTo TestFail
    
    'Arrange
    rule.FieldName = "テストフィールド"
    Dim value As Variant
    value = Null
    
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

'@TestMethod("RequiredFieldRule")
Private Sub TestValidate_WhenRuleIsDisabled_ShouldReturnTrue()
    On Error GoTo TestFail
    
    'Arrange
    rule.IsEnabled = False
    Dim value As Variant
    value = Null
    
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

'@TestMethod("RequiredFieldRule")
Private Sub TestExecutionTime_ShouldBeGreaterThanZero()
    On Error GoTo TestFail
    
    'Arrange
    Dim value As String
    value = "有効な値"
    
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