VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RequiredFieldRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("必須フィールドのバリデーションルール")

Private WithEvents BaseRule As ValidationRuleBase
Private mFieldName As String           ' フィールド名
Private mAllowEmptyString As Boolean  ' 空文字を許可するかどうか
Private mTrimValue As Boolean         ' 値をトリムするかどうか

Private Sub Class_Initialize()
    Set BaseRule = New ValidationRuleBase
    BaseRule.Name = "RequiredFieldRule"
    BaseRule.Description = "指定されたフィールドが必須であることを検証します"
    mAllowEmptyString = False
    mTrimValue = True
End Sub

'@Description("フィールド名を取得または設定します")
Public Property Get FieldName() As String
    FieldName = mFieldName
End Property

Public Property Let FieldName(ByVal value As String)
    mFieldName = value
    UpdateErrorMessage
End Property

'@Description("空文字を許可するかどうかを取得または設定します")
Public Property Get AllowEmptyString() As Boolean
    AllowEmptyString = mAllowEmptyString
End Property

Public Property Let AllowEmptyString(ByVal value As Boolean)
    mAllowEmptyString = value
End Property

'@Description("値をトリムするかどうかを取得または設定します")
Public Property Get TrimValue() As Boolean
    TrimValue = mTrimValue
End Property

Public Property Let TrimValue(ByVal value As Boolean)
    mTrimValue = value
End Property

'@Description("バリデーションを実行します")
Public Function Validate(ByVal value As Variant) As Boolean
    ' BaseRuleのValidateメソッドを呼び出し
    Validate = BaseRule.Validate(value)
End Function

'@Description("具体的なバリデーション処理を実行します")
Private Function ValidateCore(ByVal value As Variant) As Boolean
    ' 値がNothingの場合
    If value Is Nothing Then
        BaseRule.ErrorMessage = GetErrorMessage("値が設定されていません")
        ValidateCore = False
        Exit Function
    End If
    
    ' 値の型チェック
    Select Case VarType(value)
        Case vbString
            Dim strValue As String
            strValue = CStr(value)
            
            ' トリムが有効な場合
            If mTrimValue Then
                strValue = Trim(strValue)
            End If
            
            ' 空文字チェック
            If Len(strValue) = 0 Then
                If Not mAllowEmptyString Then
                    BaseRule.ErrorMessage = GetErrorMessage("空の文字列は許可されていません")
                    ValidateCore = False
                    Exit Function
                End If
            End If
            
            ValidateCore = True
            
        Case vbNull
            BaseRule.ErrorMessage = GetErrorMessage("NULL値は許可されていません")
            ValidateCore = False
            
        Case vbEmpty
            BaseRule.ErrorMessage = GetErrorMessage("Empty値は許可されていません")
            ValidateCore = False
            
        Case Else
            ' その他の型は値が設定されているとみなす
            ValidateCore = True
    End Select
End Function

'@Description("エラーメッセージを更新します")
Private Sub UpdateErrorMessage()
    If Len(mFieldName) > 0 Then
        BaseRule.ErrorMessage = mFieldName & "は必須項目です"
    Else
        BaseRule.ErrorMessage = "この項目は必須です"
    End If
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