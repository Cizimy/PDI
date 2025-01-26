VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AllowFutureDateRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "AllowFutureDateRule"

' ======================
' インターフェース実装
' ======================
Implements IValidationRule

' ======================
' プライベート変数
' ======================
Private mAllowFutureDate As Boolean
Private mErrorMessage As String

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    mAllowFutureDate = True
    mErrorMessage = ""
End Sub

' ======================
' プロパティ
' ======================
Public Property Let AllowFutureDate(ByVal value As Boolean)
    mAllowFutureDate = value
End Property

Public Property Get AllowFutureDate() As Boolean
    AllowFutureDate = mAllowFutureDate
End Property

Private Property Get IValidationRule_ErrorMessage() As String
    IValidationRule_ErrorMessage = mErrorMessage
End Property

' ======================
' メソッド
' ======================
Private Function IValidationRule_Validate(ByVal value As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' 日付型チェック
    If Not IsDate(value) Then
        mErrorMessage = "値は日付である必要があります。"
        IValidationRule_Validate = False
        Exit Function
    End If
    
    Dim dateValue As Date
    dateValue = CDate(value)
    
    ' 未来日付チェック
    If Not mAllowFutureDate And dateValue > Date Then
        mErrorMessage = "未来の日付は指定できません。"
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
        .Description = "未来日付の検証中にエラーが発生しました: " & Err.Description
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