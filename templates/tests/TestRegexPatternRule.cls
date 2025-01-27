VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestRegexPatternRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private rule As RegexPatternRule

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
    Set rule = New RegexPatternRule
    rule.FieldName = "TestField"
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'このメソッドは、各テストの後に実行されます
    Set rule = Nothing
End Sub

'@TestMethod("RegexPatternRule")
Private Sub TestValidate_WhenValueIsNotString_ShouldReturnFalse()
    On Error GoTo TestFail
    
    'Arrange
    rule.Pattern = "^\d+$"
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

'@TestMethod("RegexPatternRule")
Private Sub TestValidate_WhenPatternMatches_ShouldReturnTrue()
    On Error GoTo TestFail
    
    'Arrange
    rule.Pattern = "^\d{3}-\d{4}$"
    Dim value As String
    value = "123-4567"
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsTrue result, "パターンに一致する値は有効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) = 0, "エラーメッセージは設定されていないべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("RegexPatternRule")
Private Sub TestValidate_WhenPatternDoesNotMatch_ShouldReturnFalse()
    On Error GoTo TestFail
    
    'Arrange
    rule.Pattern = "^\d{3}-\d{4}$"
    Dim value As String
    value = "123-456"
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsFalse result, "パターンに一致しない値は無効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) > 0, "エラーメッセージが設定されているべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("RegexPatternRule")
Private Sub TestValidate_WhenPatternIsEmpty_ShouldReturnFalse()
    On Error GoTo TestFail
    
    'Arrange
    rule.Pattern = ""
    Dim value As String
    value = "test"
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsFalse result, "パターンが空の場合は無効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) > 0, "エラーメッセージが設定されているべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("RegexPatternRule")
Private Sub TestValidate_WhenIgnoreCaseIsTrue_ShouldMatchCaseInsensitive()
    On Error GoTo TestFail
    
    'Arrange
    rule.Pattern = "^[a-z]+$"
    rule.IgnoreCase = True
    Dim value As String
    value = "TEST"
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsTrue result, "大文字小文字を区別しない場合は有効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) = 0, "エラーメッセージは設定されていないべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("RegexPatternRule")
Private Sub TestValidate_WhenIgnoreCaseIsFalse_ShouldMatchCaseSensitive()
    On Error GoTo TestFail
    
    'Arrange
    rule.Pattern = "^[a-z]+$"
    rule.IgnoreCase = False
    Dim value As String
    value = "TEST"
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsFalse result, "大文字小文字を区別する場合は無効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) > 0, "エラーメッセージが設定されているべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("RegexPatternRule")
Private Sub TestValidate_WhenMultiLineIsTrue_ShouldMatchMultipleLines()
    On Error GoTo TestFail
    
    'Arrange
    rule.Pattern = "^\d+$"
    rule.MultiLine = True
    Dim value As String
    value = "123" & vbCrLf & "456"
    
    'Act
    Dim result As Boolean
    result = rule.Validate(value)
    
    'Assert
    Assert.IsTrue result, "マルチラインモードの場合は有効であるべき"
    Assert.IsTrue Len(rule.ErrorMessage) = 0, "エラーメッセージは設定されていないべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("RegexPatternRule")
Private Sub TestValidate_WhenValueIsNullAndAllowNullIsTrue_ShouldReturnTrue()
    On Error GoTo TestFail
    
    'Arrange
    rule.Pattern = "^\d+$"
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

'@TestMethod("RegexPatternRule")
Private Sub TestPredefinedPattern_Email_ShouldValidateCorrectly()
    On Error GoTo TestFail
    
    'Arrange
    rule.SetPredefinedPattern "EMAIL"
    
    'Valid email
    Dim validEmail As String
    validEmail = "test@example.com"
    
    'Invalid email
    Dim invalidEmail As String
    invalidEmail = "invalid.email"
    
    'Act & Assert
    Assert.IsTrue rule.Validate(validEmail), "有効なメールアドレスは正しく検証されるべき"
    Assert.IsFalse rule.Validate(invalidEmail), "無効なメールアドレスは正しく検証されるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("RegexPatternRule")
Private Sub TestPredefinedPattern_PhoneJP_ShouldValidateCorrectly()
    On Error GoTo TestFail
    
    'Arrange
    rule.SetPredefinedPattern "PHONE_JP"
    
    'Valid phone number
    Dim validPhone As String
    validPhone = "03-1234-5678"
    
    'Invalid phone number
    Dim invalidPhone As String
    invalidPhone = "03-1234-567"
    
    'Act & Assert
    Assert.IsTrue rule.Validate(validPhone), "有効な電話番号は正しく検証されるべき"
    Assert.IsFalse rule.Validate(invalidPhone), "無効な電話番号は正しく検証されるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("RegexPatternRule")
Private Sub TestPredefinedPattern_PostalCodeJP_ShouldValidateCorrectly()
    On Error GoTo TestFail
    
    'Arrange
    rule.SetPredefinedPattern "POSTAL_CODE_JP"
    
    'Valid postal code
    Dim validPostal As String
    validPostal = "123-4567"
    
    'Invalid postal code
    Dim invalidPostal As String
    invalidPostal = "123-456"
    
    'Act & Assert
    Assert.IsTrue rule.Validate(validPostal), "有効な郵便番号は正しく検証されるべき"
    Assert.IsFalse rule.Validate(invalidPostal), "無効な郵便番号は正しく検証されるべき"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("RegexPatternRule")
Private Sub TestErrorMessage_WhenFieldNameIsSet_ShouldIncludeFieldName()
    On Error GoTo TestFail
    
    'Arrange
    rule.FieldName = "テストフィールド"
    rule.Pattern = "^\d+$"
    Dim value As String
    value = "abc"
    
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

'@TestMethod("RegexPatternRule")
Private Sub TestValidate_WhenRuleIsDisabled_ShouldReturnTrue()
    On Error GoTo TestFail
    
    'Arrange
    rule.IsEnabled = False
    rule.Pattern = "^\d+$"
    Dim value As String
    value = "abc"
    
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

'@TestMethod("RegexPatternRule")
Private Sub TestExecutionTime_ShouldBeGreaterThanZero()
    On Error GoTo TestFail
    
    'Arrange
    rule.Pattern = "^\d+$"
    Dim value As String
    value = "123"
    
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