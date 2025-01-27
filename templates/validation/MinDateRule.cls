VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MinDateRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "MinDateRule"

' ======================
' インターフェース実装
' ======================
Implements IValidationRule

' ======================
' プライベート変数
' ======================
Private mMinDate As Date
Private mErrorMessage As String

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    mMinDate = DateSerial(1900, 1, 1)
    mErrorMessage = ""
End Sub

' ======================
' プロパティ
' ======================
Public Property Let MinDate(ByVal value As Date)
    mMinDate = value
End Property

Public Property Get MinDate() As Date
    MinDate = mMinDate
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
    
    ' 最小日付チェック
    If dateValue < mMinDate Then
        mErrorMessage = "日付が古すぎます。" & Format$(mMinDate, "yyyy/mm/dd") & "以降の日付を指定してください。"
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
        .Description = "最小日付の検証中にエラーが発生しました: " & Err.Description
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