VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestStringLengthRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private rule As StringLengthRule

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
    Set rule = New StringLengthRule
    rule.FieldName = "TestField"
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'このメソッドは、各テストの後に実行されます
    Set rule = Nothing
End Sub

'@TestMethod("StringLengthRule")
Private Sub TestValidate_WhenValueIsNotString_ShouldReturnFalse()
    On Error GoTo TestFail
    
    'Arrange
    Dim value As Long
    value = 123
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsFalse result, "文字列以外の値は無効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) > 0, "エラーメッセージが設定されているべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("StringLengthRule")
Private Sub TestValidate_WhenLengthIsWithinRange_ShouldReturnTrue()
    On Error GoTo TestFail
    
    'Arrange
    rule.MinLength = 3
    rule.MaxLength = 5
    Dim value As String
    value = "テスト"
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsTrue result, "範囲内の長さは有効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) = 0, "エラーメッセージは設定されていないべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("StringLengthRule")
Private Sub TestValidate_WhenLengthIsTooShort_ShouldReturnFalse()
    On Error GoTo TestFail
    
    'Arrange
    rule.MinLength = 3
    Dim value As String
    value = "ab"
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsFalse result, "最小長未満は無効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) > 0, "エラーメッセージが設定されているべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("StringLengthRule")
Private Sub TestValidate_WhenLengthIsTooLong_ShouldReturnFalse()
    On Error GoTo TestFail
    
    'Arrange
    rule.MaxLength = 5
    Dim value As String
    value = "123456"
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsFalse result, "最大長超過は無効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) > 0, "エラーメッセージが設定されているべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("StringLengthRule")
Private Sub TestValidate_WhenTrimValueIsTrue_ShouldTrimWhitespace()
    On Error GoTo TestFail
    
    'Arrange
    rule.MinLength = 3
    rule.MaxLength = 5
    rule.TrimValue = True
    Dim value As String
    value = "  abc  "
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsTrue result, "トリム後の長さが範囲内なら有効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) = 0, "エラーメッセージは設定されていないべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("StringLengthRule")
Private Sub TestValidate_WhenIgnoreWhitespaceIsTrue_ShouldIgnoreWhitespace()
    On Error GoTo TestFail
    
    'Arrange
    rule.MinLength = 3
    rule.MaxLength = 5
    rule.IgnoreWhitespace = True
    Dim value As String
    value = "a b c"
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsTrue result, "空白を無視した長さが範囲内なら有効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) = 0, "エラーメッセージは設定されていないべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("StringLengthRule")
Private Sub TestValidate_WhenCountByCharacterIsFalse_ShouldCountByBytes()
    On Error GoTo TestFail
    
    'Arrange
    rule.MinLength = 3
    rule.MaxLength = 5
    rule.CountByCharacter = False
    Dim value As String
    value = "あいう" ' 6バイト (Shift-JIS環境の場合)
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsFalse result, "バイト数で計算した場合は無効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) > 0, "エラーメッセージが設定されているべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("StringLengthRule")
Private Sub TestErrorMessage_WhenFieldNameIsSet_ShouldIncludeFieldName()
    On Error GoTo TestFail
    
    'Arrange
    rule.FieldName = "テストフィールド"
    rule.MinLength = 3
    Dim value As String
    value = "ab"
    
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

'@TestMethod("StringLengthRule")
Private Sub TestValidate_WhenRuleIsDisabled_ShouldReturnTrue()
    On Error GoTo TestFail
    
    'Arrange
    rule.IsEnabled = False
    rule.MinLength = 3
    Dim value As String
    value = "ab"
    
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

'@TestMethod("StringLengthRule")
Private Sub TestExecutionTime_ShouldBeGreaterThanZero()
    On Error GoTo TestFail
    
    'Arrange
    rule.MinLength = 3
    rule.MaxLength = 5
    Dim value As String
    value = "test"
    
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

'@TestMethod("StringLengthRule")
Private Sub TestValidate_WhenOnlyMinLengthIsSet_ShouldValidateCorrectly()
    On Error GoTo TestFail
    
    'Arrange
    rule.MinLength = 3
    Dim value As String
    value = "test123" ' 最小長以上なら何文字でもOK
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsTrue result, "最小長以上なら有効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) = 0, "エラーメッセージは設定されていないべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("StringLengthRule")
Private Sub TestValidate_WhenOnlyMaxLengthIsSet_ShouldValidateCorrectly()
    On Error GoTo TestFail
    
    'Arrange
    rule.MaxLength = 5
    Dim value As String
    value = "" ' 空文字列を含む最大長以下なら何文字でもOK
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsTrue result, "最大長以下なら有効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) = 0, "エラーメッセージは設定されていないべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub