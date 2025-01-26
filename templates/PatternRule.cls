VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PatternRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "PatternRule"

' ======================
' インターフェース実装
' ======================
Implements IValidationRule

' ======================
' プライベート変数
' ======================
Private mPattern As String
Private mErrorMessage As String
Private mRegexEngine As IRegexEngine

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    mPattern = ""
    mErrorMessage = ""
    Set mRegexEngine = New VBScriptRegexEngine ' デフォルトのエンジン
End Sub

Private Sub Class_Terminate()
    Set mRegexEngine = Nothing
End Sub

' ======================
' プロパティ
' ======================
Public Property Let Pattern(ByVal value As String)
    mPattern = value
End Property

Public Property Get Pattern() As String
    Pattern = mPattern
End Property

Public Property Set RegexEngine(ByVal value As IRegexEngine)
    Set mRegexEngine = value
End Property

Public Property Get RegexEngine() As IRegexEngine
    Set RegexEngine = mRegexEngine
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
    
    ' パターンが設定されていない場合は検証をスキップ
    If mPattern = "" Then
        mErrorMessage = ""
        IValidationRule_Validate = True
        Exit Function
    End If
    
    ' 正規表現エンジンが設定されていない場合はエラー
    If mRegexEngine Is Nothing Then
        Err.Raise 5, MODULE_NAME, "正規表現エンジンが設定されていません。"
    End If
    
    Dim strValue As String
    strValue = CStr(value)
    
    ' パターンチェック
    If Not mRegexEngine.Test(mPattern, strValue) Then
        mErrorMessage = "文字列が指定されたパターンと一致しません。"
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
        .Description = "パターンの検証中にエラーが発生しました: " & Err.Description
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

' ======================
' ファクトリメソッド
' ======================
Public Function CreateEmailRule() As PatternRule
    Set CreateEmailRule = New PatternRule
    With CreateEmailRule
        .Pattern = "^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$"
    End With
End Function

Public Function CreatePhoneRule() As PatternRule
    Set CreatePhoneRule = New PatternRule
    With CreatePhoneRule
        .Pattern = "^[0-9\-\(\)]+$"
    End With
End Function