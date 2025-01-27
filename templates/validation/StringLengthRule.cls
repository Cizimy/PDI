VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringLengthRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("文字列の長さのバリデーションルール")

Private WithEvents BaseRule As ValidationRuleBase
Private mFieldName As String           ' フィールド名
Private mMinLength As Long            ' 最小長
Private mMaxLength As Long            ' 最大長
Private mTrimValue As Boolean         ' 値をトリムするかどうか
Private mIgnoreWhitespace As Boolean  ' 空白を無視するかどうか
Private mCountByCharacter As Boolean  ' 文字単位でカウントするかどうか

Private Sub Class_Initialize()
    Set BaseRule = New ValidationRuleBase
    BaseRule.Name = "StringLengthRule"
    BaseRule.Description = "文字列の長さを検証します"
    mMinLength = 0
    mMaxLength = 0
    mTrimValue = True
    mIgnoreWhitespace = False
    mCountByCharacter = True
End Sub

'@Description("フィールド名を取得または設定します")
Public Property Get FieldName() As String
    FieldName = mFieldName
End Property

Public Property Let FieldName(ByVal value As String)
    mFieldName = value
    UpdateErrorMessage
End Property

'@Description("最小長を取得または設定します")
Public Property Get MinLength() As Long
    MinLength = mMinLength
End Property

Public Property Let MinLength(ByVal value As Long)
    If value < 0 Then
        Err.Raise 5, "StringLengthRule", "最小長は0以上である必要があります"
    End If
    mMinLength = value
    UpdateErrorMessage
End Property

'@Description("最大長を取得または設定します")
Public Property Get MaxLength() As Long
    MaxLength = mMaxLength
End Property

Public Property Let MaxLength(ByVal value As Long)
    If value < 0 Then
        Err.Raise 5, "StringLengthRule", "最大長は0以上である必要があります"
    End If
    If value > 0 And value < mMinLength Then
        Err.Raise 5, "StringLengthRule", "最大長は最小長以上である必要があります"
    End If
    mMaxLength = value
    UpdateErrorMessage
End Property

'@Description("値をトリムするかどうかを取得または設定します")
Public Property Get TrimValue() As Boolean
    TrimValue = mTrimValue
End Property

Public Property Let TrimValue(ByVal value As Boolean)
    mTrimValue = value
End Property

'@Description("空白を無視するかどうかを取得または設定します")
Public Property Get IgnoreWhitespace() As Boolean
    IgnoreWhitespace = mIgnoreWhitespace
End Property

Public Property Let IgnoreWhitespace(ByVal value As Boolean)
    mIgnoreWhitespace = value
End Property

'@Description("文字単位でカウントするかどうかを取得または設定します")
Public Property Get CountByCharacter() As Boolean
    CountByCharacter = mCountByCharacter
End Property

Public Property Let CountByCharacter(ByVal value As Boolean)
    mCountByCharacter = value
End Property

'@Description("バリデーションを実行します")
Public Function Validate(ByVal value As Variant) As Boolean
    Validate = BaseRule.Validate(value)
End Function

'@Description("具体的なバリデーション処理を実行します")
Private Function ValidateCore(ByVal value As Variant) As Boolean
    ' 値がNothingまたはNullの場合
    If value Is Nothing Or IsNull(value) Then
        BaseRule.ErrorMessage = GetErrorMessage("値が設定されていません")
        ValidateCore = False
        Exit Function
    End If
    
    ' 文字列以外の型の場合
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
    
    ' 空白を無視する場合
    If mIgnoreWhitespace Then
        strValue = Replace(strValue, " ", "")
    End If
    
    ' 長さの計算
    Dim length As Long
    If mCountByCharacter Then
        length = Len(strValue)
    Else
        length = LenB(StrConv(strValue, vbFromUnicode))
    End If
    
    ' 最小長チェック
    If mMinLength > 0 And length < mMinLength Then
        BaseRule.ErrorMessage = GetErrorMessage("最小長(" & mMinLength & ")未満です")
        ValidateCore = False
        Exit Function
    End If
    
    ' 最大長チェック
    If mMaxLength > 0 And length > mMaxLength Then
        BaseRule.ErrorMessage = GetErrorMessage("最大長(" & mMaxLength & ")を超えています")
        ValidateCore = False
        Exit Function
    End If
    
    ValidateCore = True
End Function

'@Description("エラーメッセージを更新します")
Private Sub UpdateErrorMessage()
    Dim msg As String
    
    If Len(mFieldName) > 0 Then
        msg = mFieldName & "の長さは"
    Else
        msg = "長さは"
    End If
    
    If mMinLength > 0 And mMaxLength > 0 Then
        msg = msg & mMinLength & "文字以上" & mMaxLength & "文字以下である必要があります"
    ElseIf mMinLength > 0 Then
        msg = msg & mMinLength & "文字以上である必要があります"
    ElseIf mMaxLength > 0 Then
        msg = msg & mMaxLength & "文字以下である必要があります"
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
