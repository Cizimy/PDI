VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBScriptRegexEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "VBScriptRegexEngine"

' ======================
' インターフェース実装
' ======================
Implements IRegexEngine

' ======================
' プライベート変数
' ======================
Private mRegex As Object ' VBScript.RegExp

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mRegex = CreateObject("VBScript.RegExp")
End Sub

Private Sub Class_Terminate()
    Set mRegex = Nothing
End Sub

' ======================
' メソッド
' ======================
Private Function IRegexEngine_Test(ByVal pattern As String, ByVal text As String, Optional ByVal ignoreCase As Boolean = True) As Boolean
    On Error GoTo ErrorHandler
    
    With mRegex
        .Pattern = pattern
        .Global = True
        .IgnoreCase = ignoreCase
        IRegexEngine_Test = .Test(text)
    End With
    Exit Function
    
ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "正規表現のテスト中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "IRegexEngine_Test"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    
    Err.Raise Err.Number, Err.Source, "正規表現のテスト中にエラーが発生しました: " & Err.Description
End Function

Private Function IRegexEngine_Replace(ByVal pattern As String, ByVal text As String, ByVal replacement As String, Optional ByVal ignoreCase As Boolean = True) As String
    On Error GoTo ErrorHandler
    
    With mRegex
        .Pattern = pattern
        .Global = True
        .IgnoreCase = ignoreCase
        IRegexEngine_Replace = .Replace(text, replacement)
    End With
    Exit Function
    
ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "正規表現による置換中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "IRegexEngine_Replace"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    
    Err.Raise Err.Number, Err.Source, "正規表現による置換中にエラーが発生しました: " & Err.Description
End Function

Private Function IRegexEngine_Match(ByVal pattern As String, ByVal text As String, Optional ByVal ignoreCase As Boolean = True) As Variant
    On Error GoTo ErrorHandler
    
    With mRegex
        .Pattern = pattern
        .Global = True
        .IgnoreCase = ignoreCase
        
        Dim matches As Object
        Set matches = .Execute(text)
        
        If matches.Count = 0 Then
            IRegexEngine_Match = Array()
            Exit Function
        End If
        
        Dim result() As String
        ReDim result(matches.Count - 1)
        
        Dim i As Long
        For i = 0 To matches.Count - 1
            result(i) = matches(i).Value
        Next i
        
        IRegexEngine_Match = result
    End With
    Exit Function
    
ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "正規表現によるマッチング中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "IRegexEngine_Match"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    
    Err.Raise Err.Number, Err.Source, "正規表現によるマッチング中にエラーが発生しました: " & Err.Description
End Function