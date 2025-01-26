VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "MockError"

' ======================
' インターフェース実装
' ======================
Implements IError

' ======================
' プライベート変数
' ======================
Private mErrors As Collection
Private mLastError As ErrorInfo

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mErrors = New Collection
End Sub

Private Sub Class_Terminate()
    Set mErrors = Nothing
    Set mLastError = Nothing
End Sub

' ======================
' プロパティ
' ======================
Public Property Get Errors() As Collection
    Set Errors = mErrors
End Property

Public Property Get LastError() As ErrorInfo
    Set LastError = mLastError
End Property

' ======================
' メソッド
' ======================
Private Sub IError_HandleError(ByVal errorInfo As ErrorInfo)
    Set mLastError = errorInfo
    mErrors.Add errorInfo
End Sub

Public Sub ClearErrors()
    Set mErrors = New Collection
    Set mLastError = Nothing
End Sub

Public Function GetErrorCount() As Long
    GetErrorCount = mErrors.Count
End Function

Public Function GetErrorAt(ByVal index As Long) As ErrorInfo
    If index <= 0 Or index > mErrors.Count Then
        Err.Raise 9, MODULE_NAME, "インデックスが範囲外です。"
    End If
    Set GetErrorAt = mErrors(index)
End Function

Public Function ContainsError(ByVal errorCode As Long) As Boolean
    Dim err As ErrorInfo
    For Each err In mErrors
        If err.Code = errorCode Then
            ContainsError = True
            Exit Function
        End If
    Next err
    ContainsError = False
End Function

Public Function ContainsErrorMessage(ByVal message As String) As Boolean
    Dim err As ErrorInfo
    For Each err In mErrors
        If InStr(1, err.Description, message, vbTextCompare) > 0 Then
            ContainsErrorMessage = True
            Exit Function
        End If
    Next err
    ContainsErrorMessage = False
End Function