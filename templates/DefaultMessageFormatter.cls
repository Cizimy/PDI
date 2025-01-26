VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DefaultMessageFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "DefaultMessageFormatter"

' ======================
' インターフェース実装
' ======================
Implements IMessageFormatter

' ======================
' IMessageFormatter実装
' ======================
Private Function IMessageFormatter_FormatMessage(ByVal message As String, ByRef errorInfo As ErrorInfo) As String
    On Error GoTo ErrorHandler
    
    Dim formattedMessage As String
    
    ' 基本メッセージの設定
    formattedMessage = IIf(Len(message) > 0, message & vbCrLf & vbCrLf, "エラーが発生しました。" & vbCrLf & vbCrLf)
    
    ' エラー詳細情報の追加
    formattedMessage = formattedMessage & _
                      "説明: " & errorInfo.Description & vbCrLf & _
                      "発生場所: " & errorInfo.Source & "." & errorInfo.ProcedureName
    
    ' スタックトレースが存在する場合は追加
    If Len(errorInfo.StackTrace) > 0 Then
        formattedMessage = formattedMessage & vbCrLf & vbCrLf & _
                          "スタックトレース:" & vbCrLf & _
                          errorInfo.StackTrace
    End If
    
    IMessageFormatter_FormatMessage = formattedMessage
    Exit Function
    
ErrorHandler:
    ' エラー情報の作成
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "メッセージのフォーマット中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "IMessageFormatter_FormatMessage"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    
    ' エラーハンドラに委譲
    modError.HandleError errDetail
    
    ' デフォルトメッセージを返す
    IMessageFormatter_FormatMessage = "エラーメッセージのフォーマットに失敗しました。"
End Function

#If DEBUG Then
    ' ======================
    ' テストサポート機能
    ' ======================
    Public Function TestFormatMessage(ByVal testMessage As String, ByRef testError As ErrorInfo) As String
        TestFormatMessage = IMessageFormatter_FormatMessage(testMessage, testError)
    End Function
#End If