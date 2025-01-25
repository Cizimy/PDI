VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "StringValidator"

' ======================
' インターフェース実装
' ======================
Implements IValidator

' ======================
' プライベート変数
' ======================
Private mErrorMessage As String
Private mMinLength As Long
Private mMaxLength As Long
Private mPattern As String

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    mMinLength = 0
    mMaxLength = 2147483647 ' Long型の最大値
    mPattern = ""
    mErrorMessage = ""
End Sub

' ======================
' プロパティ
' ======================
Public Property Let MinLength(ByVal value As Long)
    mMinLength = value
End Property

Public Property Get MinLength() As Long
    MinLength = mMinLength
End Property

Public Property Let MaxLength(ByVal value As Long)
    mMaxLength = value
End Property

Public Property Get MaxLength() As Long
    MaxLength = mMaxLength
End Property

Public Property Let Pattern(ByVal value As String)
    mPattern = value
End Property

Public Property Get Pattern() As String
    Pattern = mPattern
End Property

Private Property Get IValidator_ErrorMessage() As String
    IValidator_ErrorMessage = mErrorMessage
End Property

' ======================
' メソッド
' ======================
Private Function IValidator_Validate(ByVal value As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' 文字列型チェック
    If VarType(value) <> vbString Then
        mErrorMessage = "値は文字列である必要があります。"
        IValidator_Validate = False
        Exit Function
    End If
    
    Dim strValue As String
    strValue = CStr(value)
    
    ' 長さチェック
    If Len(strValue) < mMinLength Then
        mErrorMessage = "文字列の長さが短すぎます。最小" & mMinLength & "文字必要です。"
        IValidator_Validate = False
        Exit Function
    End If
    
    If Len(strValue) > mMaxLength Then
        mErrorMessage = "文字列の長さが長すぎます。最大" & mMaxLength & "文字までです。"
        IValidator_Validate = False
        Exit Function
    End If
    
    ' パターンチェック
    If mPattern <> "" Then
        Dim regex As Object
        Set regex = CreateObject("VBScript.RegExp")
        With regex
            .Pattern = mPattern
            .Global = True
            .IgnoreCase = True
            
            If Not .Test(strValue) Then
                mErrorMessage = "文字列が指定されたパターンと一致しません。"
                IValidator_Validate = False
                Exit Function
            End If
        End With
    End If
    
    mErrorMessage = ""
    IValidator_Validate = True
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "文字列の検証中にエラーが発生しました: " & Err.Description
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
Public Function CreateEmailValidator() As StringValidator
    Set CreateEmailValidator = New StringValidator
    With CreateEmailValidator
        .Pattern = "^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$"
    End With
End Function

Public Function CreatePhoneValidator() As StringValidator
    Set CreatePhoneValidator = New StringValidator
    With CreatePhoneValidator
        .Pattern = "^[0-9\-\(\)]+$"
        .MinLength = 10
        .MaxLength = 15
    End With
End Function