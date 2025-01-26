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
Private mRules As Collection
Private mError As IError

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mRules = New Collection
End Sub

Private Sub Class_Terminate()
    Set mRules = Nothing
    Set mError = Nothing
End Sub

' ======================
' プロパティ
' ======================
Public Property Set Error(ByVal value As IError)
    Set mError = value
End Property

Public Property Get Error() As IError
    Set Error = mError
End Property

Private Property Get IValidator_ErrorMessage() As String
    ' エラーメッセージは個別のルールから取得するため、このプロパティは使用しない
    IValidator_ErrorMessage = ""
End Property

' ======================
' メソッド
' ======================
Public Sub AddRule(ByVal rule As IValidationRule)
    If rule Is Nothing Then
        Err.Raise 5, MODULE_NAME, "ルールがNothingです。"
    End If
    mRules.Add rule
End Sub

Public Sub ClearRules()
    Set mRules = New Collection
End Sub

Private Function IValidator_Validate(ByVal value As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' エラーハンドラーが設定されていない場合はエラー
    If mError Is Nothing Then
        Err.Raise 5, MODULE_NAME, "エラーハンドラーが設定されていません。"
    End If
    
    ' 日付型チェック
    If Not IsDate(value) Then
        Dim errInfo As ErrorInfo
        With errInfo
            .Code = ErrInvalidInput
            .Description = "値は日付である必要があります。"
            .Category = ECValidation
            .Source = MODULE_NAME
            .ProcedureName = "IValidator_Validate"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        mError.HandleError errInfo
        IValidator_Validate = False
        Exit Function
    End If
    
    ' 各ルールの検証
    Dim rule As IValidationRule
    For Each rule In mRules
        If Not rule.Validate(value) Then
            Dim errInfo2 As ErrorInfo
            With errInfo2
                .Code = ErrInvalidInput
                .Description = rule.ErrorMessage
                .Category = ECValidation
                .Source = MODULE_NAME
                .ProcedureName = "IValidator_Validate"
                .StackTrace = modStackTrace.GetStackTrace()
                .OccurredAt = Now
            End With
            mError.HandleError errInfo2
            IValidator_Validate = False
            Exit Function
        End If
    Next rule
    
    IValidator_Validate = True
    Exit Function

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "日付の検証中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "IValidator_Validate"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    
    IValidator_Validate = False
End Function

' ======================
' ファクトリメソッド
' ======================
Public Function CreateBirthDateValidator(ByVal errorHandler As IError) As DateValidator
    Set CreateBirthDateValidator = New DateValidator
    With CreateBirthDateValidator
        Set .Error = errorHandler
        
        Dim minDateRule As MinDateRule
        Set minDateRule = New MinDateRule
        minDateRule.MinDate = DateSerial(1900, 1, 1)
        .AddRule minDateRule
        
        Dim maxDateRule As MaxDateRule
        Set maxDateRule = New MaxDateRule
        maxDateRule.MaxDate = Date
        .AddRule maxDateRule
        
        Dim futureDateRule As AllowFutureDateRule
        Set futureDateRule = New AllowFutureDateRule
        futureDateRule.AllowFutureDate = False
        .AddRule futureDateRule
    End With
End Function

Public Function CreateExpiryDateValidator(ByVal errorHandler As IError) As DateValidator
    Set CreateExpiryDateValidator = New DateValidator
    With CreateExpiryDateValidator
        Set .Error = errorHandler
        
        Dim minDateRule As MinDateRule
        Set minDateRule = New MinDateRule
        minDateRule.MinDate = Date
        .AddRule minDateRule
        
        Dim maxDateRule As MaxDateRule
        Set maxDateRule = New MaxDateRule
        maxDateRule.MaxDate = DateAdd("yyyy", 100, Date)
        .AddRule maxDateRule
        
        Dim pastDateRule As AllowPastDateRule
        Set pastDateRule = New AllowPastDateRule
        pastDateRule.AllowPastDate = False
        .AddRule pastDateRule
    End With
End Function