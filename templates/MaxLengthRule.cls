VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MaxLengthRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "MaxLengthRule"

' ======================
' インターフェース実装
' ======================
Implements IValidationRule

' ======================
' プライベート変数
' ======================
Private mMaxLength As Long
Private mErrorMessage As String

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    mMaxLength = 2147483647 ' Long型の最大値
    mErrorMessage = ""
End Sub

' ======================
' プロパティ
' ======================
Public Property Let MaxLength(ByVal value As Long)
    If value < 0 Then
        Err.Raise 5, MODULE_NAME, "最大長は0以上である必要があります。"
    End If
    mMaxLength = value
End Property

Public Property Get MaxLength() As Long
    MaxLength = mMaxLength
End Property

Private Property Get IValidationRule_ErrorMessage() As String
    IValidationRule_ErrorMessage = mErrorMessage
End Property

' ======================
' メソッド
' ======================
Private Function IValidationRule_Validate(ByVal value As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' 文字列型チェック
    If VarType(value) <> vbString Then
        mErrorMessage = "値は文字列である必要があります。"
        IValidationRule_Validate = False
        Exit Function
    End If
    
    Dim strValue As String
    strValue = CStr(value)
    
    ' 長さチェック
    If Len(strValue) > mMaxLength Then
        mErrorMessage = "文字列の長さが長すぎます。最大" & mMaxLength & "文字までです。"
        IValidationRule_Validate = False
        Exit Function
    End If
    
    mErrorMessage = ""
    IValidationRule_Validate = True
    Exit Function

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "最大長の検証中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "IValidationRule_Validate"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    
    mErrorMessage = "検証中に予期せぬエラーが発生しました。"
    IValidationRule_Validate = False
End Function