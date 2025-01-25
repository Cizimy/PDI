VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "DateValidator"

' ======================
' インターフェース実装
' ======================
Implements IValidator

' ======================
' プライベート変数
' ======================
Private mErrorMessage As String
Private mMinDate As Date
Private mMaxDate As Date
Private mAllowFutureDate As Boolean
Private mAllowPastDate As Boolean

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    mMinDate = DateSerial(1900, 1, 1)
    mMaxDate = DateSerial(9999, 12, 31)
    mAllowFutureDate = True
    mAllowPastDate = True
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

Public Property Let MaxDate(ByVal value As Date)
    mMaxDate = value
End Property

Public Property Get MaxDate() As Date
    MaxDate = mMaxDate
End Property

Public Property Let AllowFutureDate(ByVal value As Boolean)
    mAllowFutureDate = value
End Property

Public Property Get AllowFutureDate() As Boolean
    AllowFutureDate = mAllowFutureDate
End Property

Public Property Let AllowPastDate(ByVal value As Boolean)
    mAllowPastDate = value
End Property

Public Property Get AllowPastDate() As Boolean
    AllowPastDate = mAllowPastDate
End Property

Private Property Get IValidator_ErrorMessage() As String
    IValidator_ErrorMessage = mErrorMessage
End Property

' ======================
' メソッド
' ======================
Private Function IValidator_Validate(ByVal value As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' 日付型チェック
    If Not IsDate(value) Then
        mErrorMessage = "有効な日付形式ではありません。"
        IValidator_Validate = False
        Exit Function
    End If
    
    Dim dateValue As Date
    dateValue = CDate(value)
    
    ' 範囲チェック
    If dateValue < mMinDate Then
        mErrorMessage = "日付が古すぎます。" & Format$(mMinDate, "yyyy/mm/dd") & "以降の日付を指定してください。"
        IValidator_Validate = False
        Exit Function
    End If
    
    If dateValue > mMaxDate Then
        mErrorMessage = "日付が新しすぎます。" & Format$(mMaxDate, "yyyy/mm/dd") & "以前の日付を指定してください。"
        IValidator_Validate = False
        Exit Function
    End If
    
    ' 未来日付チェック
    If Not mAllowFutureDate And dateValue > Date Then
        mErrorMessage = "未来の日付は指定できません。"
        IValidator_Validate = False
        Exit Function
    End If
    
    ' 過去日付チェック
    If Not mAllowPastDate And dateValue < Date Then
        mErrorMessage = "過去の日付は指定できません。"
        IValidator_Validate = False
        Exit Function
    End If
    
    mErrorMessage = ""
    IValidator_Validate = True
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "日付の検証中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "IValidator_Validate"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    
    mErrorMessage = "検証中に予期せぬエラーが発生しました。"
    IValidator_Validate = False
End Function

' ======================
' ファクトリメソッド
' ======================
Public Function CreateBirthDateValidator() As DateValidator
    Set CreateBirthDateValidator = New DateValidator
    With CreateBirthDateValidator
        .MinDate = DateSerial(1900, 1, 1)
        .MaxDate = Date
        .AllowFutureDate = False
    End With
End Function

Public Function CreateExpiryDateValidator() As DateValidator
    Set CreateExpiryDateValidator = New DateValidator
    With CreateExpiryDateValidator
        .MinDate = Date
        .MaxDate = DateAdd("yyyy", 100, Date)
        .AllowPastDate = False
    End With
End Function