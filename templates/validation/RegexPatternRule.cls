VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegexPatternRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("正規表現パターンのバリデーションルール")

Private WithEvents BaseRule As ValidationRuleBase
Private mFieldName As String           ' フィールド名
Private mPattern As String            ' 正規表現パターン
Private mErrorDescription As String    ' エラーの説明
Private mIgnoreCase As Boolean        ' 大文字小文字を区別しないかどうか
Private mGlobal As Boolean            ' グローバルマッチを使用するかどうか
Private mMultiLine As Boolean         ' マルチラインモードを使用するかどうか
Private mRegEx As Object              ' RegExpオブジェクト
Private mTrimValue As Boolean         ' 値をトリムするかどうか
Private mAllowNull As Boolean         ' Null値を許可するかどうか

' 定義済みパターン
Private Const PATTERN_EMAIL As String = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
Private Const PATTERN_PHONE_JP As String = "^0\d{1,4}-\d{1,4}-\d{4}$"
Private Const PATTERN_POSTAL_CODE_JP As String = "^\d{3}-\d{4}$"
Private Const PATTERN_URL As String = "^(http|https)://[a-zA-Z0-9\-\.]+\.[a-zA-Z]{2,}(:[0-9]+)?(/.*)?$"
Private Const PATTERN_DATE_ISO As String = "^\d{4}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$"
Private Const PATTERN_TIME_24H As String = "^([01]\d|2[0-3]):([0-5]\d):([0-5]\d)$"

Private Sub Class_Initialize()
    Set BaseRule = New ValidationRuleBase
    BaseRule.Name = "RegexPatternRule"
    BaseRule.Description = "正規表現パターンに基づいて文字列を検証します"
    mIgnoreCase = True
    mGlobal = False
    mMultiLine = False
    mTrimValue = True
    mAllowNull = False
    InitializeRegEx
End Sub

Private Sub Class_Terminate()
    Set mRegEx = Nothing
End Sub

'@Description("RegExpオブジェクトを初期化します")
Private Sub InitializeRegEx()
    Set mRegEx = CreateObject("VBScript.RegExp")
    mRegEx.IgnoreCase = mIgnoreCase
    mRegEx.Global = mGlobal
    mRegEx.MultiLine = mMultiLine
    If Len(mPattern) > 0 Then
        mRegEx.Pattern = mPattern
    End If
End Sub

'@Description("フィールド名を取得または設定します")
Public Property Get FieldName() As String
    FieldName = mFieldName
End Property

Public Property Let FieldName(ByVal value As String)
    mFieldName = value
    UpdateErrorMessage
End Property

'@Description("正規表現パターンを取得または設定します")
Public Property Get Pattern() As String
    Pattern = mPattern
End Property

Public Property Let Pattern(ByVal value As String)
    mPattern = value
    If Not mRegEx Is Nothing Then
        mRegEx.Pattern = value
    End If
    UpdateErrorMessage
End Property

'@Description("エラーの説明を取得または設定します")
Public Property Get ErrorDescription() As String
    ErrorDescription = mErrorDescription
End Property

Public Property Let ErrorDescription(ByVal value As String)
    mErrorDescription = value
    UpdateErrorMessage
End Property

'@Description("大文字小文字を区別しないかどうかを取得または設定します")
Public Property Get IgnoreCase() As Boolean
    IgnoreCase = mIgnoreCase
End Property

Public Property Let IgnoreCase(ByVal value As Boolean)
    mIgnoreCase = value
    If Not mRegEx Is Nothing Then
        mRegEx.IgnoreCase = value
    End If
End Property

'@Description("グローバルマッチを使用するかどうかを取得または設定します")
Public Property Get GlobalMatch() As Boolean
    GlobalMatch = mGlobal
End Property

Public Property Let GlobalMatch(ByVal value As Boolean)
    mGlobal = value
    If Not mRegEx Is Nothing Then
        mRegEx.Global = value
    End If
End Property

'@Description("マルチラインモードを使用するかどうかを取得または設定します")
Public Property Get MultiLine() As Boolean
    MultiLine = mMultiLine
End Property

Public Property Let MultiLine(ByVal value As Boolean)
    mMultiLine = value
    If Not mRegEx Is Nothing Then
        mRegEx.MultiLine = value
    End If
End Property

'@Description("値をトリムするかどうかを取得または設定します")
Public Property Get TrimValue() As Boolean
    TrimValue = mTrimValue
End Property

Public Property Let TrimValue(ByVal value As Boolean)
    mTrimValue = value
End Property

'@Description("Null値を許可するかどうかを取得または設定します")
Public Property Get AllowNull() As Boolean
    AllowNull = mAllowNull
End Property

Public Property Let AllowNull(ByVal value As Boolean)
    mAllowNull = value
End Property

'@Description("定義済みパターンを設定します")
Public Sub SetPredefinedPattern(ByVal patternType As String)
    Select Case UCase(patternType)
        Case "EMAIL"
            Pattern = PATTERN_EMAIL
            ErrorDescription = "有効なメールアドレス形式である必要があります"
            
        Case "PHONE_JP"
            Pattern = PATTERN_PHONE_JP
            ErrorDescription = "有効な電話番号形式である必要があります（例：03-1234-5678）"
            
        Case "POSTAL_CODE_JP"
            Pattern = PATTERN_POSTAL_CODE_JP
            ErrorDescription = "有効な郵便番号形式である必要があります（例：123-4567）"
            
        Case "URL"
            Pattern = PATTERN_URL
            ErrorDescription = "有効なURL形式である必要があります"
            
        Case "DATE_ISO"
            Pattern = PATTERN_DATE_ISO
            ErrorDescription = "有効な日付形式である必要があります（例：2025-01-27）"
            
        Case "TIME_24H"
            Pattern = PATTERN_TIME_24H
            ErrorDescription = "有効な時刻形式である必要があります（例：23:59:59）"
            
        Case Else
            Err.Raise 5, "RegexPatternRule", "未定義のパターンタイプです: " & patternType
    End Select
End Sub

'@Description("バリデーションを実行します")
Public Function Validate(ByVal value As Variant) As Boolean
    Validate = BaseRule.Validate(value)
End Function

'@Description("具体的なバリデーション処理を実行します")
Private Function ValidateCore(ByVal value As Variant) As Boolean
    ' Null値のチェック
    If IsNull(value) Then
        If mAllowNull Then
            ValidateCore = True
        Else
            BaseRule.ErrorMessage = GetErrorMessage("NULL値は許可されていません")
            ValidateCore = False
        End If
        Exit Function
    End If
    
    ' 文字列型のチェック
    If VarType(value) <> vbString Then
        BaseRule.ErrorMessage = GetErrorMessage("文字列型である必要があります")
        ValidateCore = False
        Exit Function
    End If
    
    Dim strValue As String
    strValue = CStr(value)
    
    ' トリムが有効な場合
    If mTrimValue Then
        strValue = Trim(strValue)
    End If
    
    ' パターンが設定されていない場合
    If Len(mPattern) = 0 Then
        BaseRule.ErrorMessage = GetErrorMessage("正規表現パターンが設定されていません")
        ValidateCore = False
        Exit Function
    End If
    
    ' パターンマッチング
    If mRegEx.Test(strValue) Then
        ValidateCore = True
    Else
        BaseRule.ErrorMessage = GetErrorMessage(mErrorDescription)
        ValidateCore = False
    End If
End Function

'@Description("エラーメッセージを更新します")
Private Sub UpdateErrorMessage()
    Dim msg As String
    
    If Len(mFieldName) > 0 Then
        msg = mFieldName & "は"
    End If
    
    If Len(mErrorDescription) > 0 Then
        msg = msg & mErrorDescription
    Else
        msg = msg & "指定されたパターンに一致する必要があります"
    End If
    
    BaseRule.ErrorMessage = msg
End Sub

'@Description("エラーメッセージを生成します")
Private Function GetErrorMessage(ByVal detail As String) As String
    If Len(mFieldName) > 0 Then
        GetErrorMessage = mFieldName & ": " & detail
    Else
        GetErrorMessage = detail
    End If
End Function

' BaseRuleのイベントハンドラ
Private Sub BaseRule_RuleExecuted(ByVal ruleName As String, ByVal success As Boolean, ByVal executionTime As Double)
    ' 必要に応じてイベントを処理
End Sub

' ValidationRuleBaseのプロパティとメソッドの委譲
Public Property Get Name() As String
    Name = BaseRule.Name
End Property

Public Property Let Name(ByVal value As String)
    BaseRule.Name = value
End Property

Public Property Get Description() As String
    Description = BaseRule.Description
End Property

Public Property Let Description(ByVal value As String)
    BaseRule.Description = value
End Property

Public Property Get ErrorMessage() As String
    ErrorMessage = BaseRule.ErrorMessage
End Property

Public Property Let ErrorMessage(ByVal value As String)
    BaseRule.ErrorMessage = value
End Property

Public Property Get IsEnabled() As Boolean
    IsEnabled = BaseRule.IsEnabled
End Property

Public Property Let IsEnabled(ByVal value As Boolean)
    BaseRule.IsEnabled = value
End Property

Public Property Get Priority() As ValidationPriority
    Priority = BaseRule.Priority
End Property

Public Property Let Priority(ByVal value As ValidationPriority)
    BaseRule.Priority = value
End Property

Public Property Get Dependencies() As Collection
    Set Dependencies = BaseRule.Dependencies
End Property

Public Sub AddDependency(ByVal rule As IValidationRule)
    BaseRule.AddDependency rule
End Sub

Public Property Get ValidationContext() As IValidationContext
    Set ValidationContext = BaseRule.ValidationContext
End Property

Public Property Set ValidationContext(ByVal value As IValidationContext)
    Set BaseRule.ValidationContext = value
End Property

Public Property Get ExecutionTime() As Double
    ExecutionTime = BaseRule.ExecutionTime
End Property

Public Property Get LastExecutionResult() As Boolean
    LastExecutionResult = BaseRule.LastExecutionResult
End Property

Public Property Get ErrorDetails() As Collection
    Set ErrorDetails = BaseRule.ErrorDetails
End Property

Public Function ToJSON() As String
    ToJSON = BaseRule.ToJSON
End Function