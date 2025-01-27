VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestDefaultMessageFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "TestDefaultMessageFormatter"

' テスト対象
Private mFormatter As DefaultMessageFormatter

' ======================
' セットアップ/クリーンアップ
' ======================
Private Sub Class_Initialize()
    Set mFormatter = New DefaultMessageFormatter
End Sub

Private Sub Class_Terminate()
    Set mFormatter = Nothing
End Sub

' ======================
' テストケース
' ======================

''' <summary>
''' 基本的なメッセージフォーマットのテスト
''' </summary>
Public Sub TestBasicFormatting()
    On Error GoTo TestFail
    
    ' テスト用エラー情報の作成
    Dim testError As ErrorInfo
    With testError
        .Code = ErrUnexpected
        .Description = "テストエラー"
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "TestBasicFormatting"
        .StackTrace = ""
    End With
    
    ' メッセージのフォーマット
    Dim result As String
    result = mFormatter.TestFormatMessage("テストメッセージ", testError)
    
    ' 結果の検証
    If InStr(result, "テストメッセージ") > 0 And _
       InStr(result, "テストエラー") > 0 And _
       InStr(result, MODULE_NAME) > 0 Then
        Debug.Print "TestBasicFormatting: OK"
    Else
        Debug.Print "TestBasicFormatting: Failed - Unexpected format"
        Debug.Print "Result: " & result
    End If
    Exit Sub
    
TestFail:
    Debug.Print "TestBasicFormatting: Failed - " & Err.Description
End Sub

''' <summary>
''' スタックトレース付きメッセージのテスト
''' </summary>
Public Sub TestWithStackTrace()
    On Error GoTo TestFail
    
    ' テスト用エラー情報の作成
    Dim testError As ErrorInfo
    With testError
        .Code = ErrUnexpected
        .Description = "テストエラー"
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "TestWithStackTrace"
        .StackTrace = "テストスタックトレース" & vbCrLf & "行1" & vbCrLf & "行2"
    End With
    
    ' メッセージのフォーマット
    Dim result As String
    result = mFormatter.TestFormatMessage("", testError)
    
    ' 結果の検証
    If InStr(result, "スタックトレース") > 0 And _
       InStr(result, "テストスタックトレース") > 0 Then
        Debug.Print "TestWithStackTrace: OK"
    Else
        Debug.Print "TestWithStackTrace: Failed - Stack trace not found"
        Debug.Print "Result: " & result
    End If
    Exit Sub
    
TestFail:
    Debug.Print "TestWithStackTrace: Failed - " & Err.Description
End Sub

''' <summary>
''' 空のメッセージ処理のテスト
''' </summary>
Public Sub TestEmptyMessage()
    On Error GoTo TestFail
    
    ' テスト用エラー情報の作成
    Dim testError As ErrorInfo
    With testError
        .Code = ErrUnexpected
        .Description = "テストエラー"
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "TestEmptyMessage"
    End With
    
    ' メッセージのフォーマット
    Dim result As String
    result = mFormatter.TestFormatMessage("", testError)
    
    ' 結果の検証
    If InStr(result, "エラーが発生しました") > 0 Then
        Debug.Print "TestEmptyMessage: OK"
    Else
        Debug.Print "TestEmptyMessage: Failed - Default message not found"
        Debug.Print "Result: " & result
    End If
    Exit Sub
    
TestFail:
    Debug.Print "TestEmptyMessage: Failed - " & Err.Description
End Sub

''' <summary>
''' エラー処理のテスト
''' </summary>
Public Sub TestErrorHandling()
    On Error Resume Next
    
    ' 不正なエラー情報でテスト
    Dim result As String
    result = mFormatter.TestFormatMessage("テスト", Nothing)
    
    ' エラーが適切に処理されることを確認
    If InStr(result, "エラーメッセージのフォーマットに失敗") > 0 Then
        Debug.Print "TestErrorHandling: OK"
    Else
        Debug.Print "TestErrorHandling: Failed - Error not handled properly"
        Debug.Print "Result: " & result
    End If
    
    On Error GoTo 0
End Sub

' ======================
' テスト実行
' ======================
Public Sub RunAllTests()
    TestBasicFormatting
    TestWithStackTrace
    TestEmptyMessage
    TestErrorHandling
End Sub