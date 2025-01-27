VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationRuleBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

Implements IValidationRule

'@Description("バリデーションルールの基本クラス")

' メンバー変数
Private mName As String                ' ルール名
Private mDescription As String         ' ルールの説明
Private mErrorMessage As String        ' エラーメッセージ
Private mIsEnabled As Boolean          ' 有効/無効フラグ
Private mPriority As ValidationPriority ' 優先度
Private mDependencies As Collection    ' 依存するルール
Private mValidationContext As IValidationContext ' バリデーションコンテキスト
Private mExecutionTime As Double       ' 実行時間
Private mLastExecutionResult As Boolean ' 最後の実行結果
Private mErrorDetails As Collection    ' エラー詳細

' イベント
Public Event RuleExecuted(ByVal ruleName As String, ByVal success As Boolean, ByVal executionTime As Double)
Public Event RuleDisabled(ByVal ruleName As String)
Public Event RuleEnabled(ByVal ruleName As String)
Public Event DependencyAdded(ByVal ruleName As String, ByVal dependencyName As String)

Private Sub Class_Initialize()
    mIsEnabled = True
    mPriority = vpNormal
    Set mDependencies = New Collection
    Set mErrorDetails = New Collection
    mExecutionTime = 0
    mLastExecutionResult = False
End Sub

'@Description("ルール名を取得または設定します")
Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)
    mName = value
End Property

'@Description("ルールの説明を取得または設定します")
Public Property Get Description() As String
    Description = mDescription
End Property

Public Property Let Description(ByVal value As String)
    mDescription = value
End Property

'@Description("エラーメッセージを取得または設定します")
Public Property Get ErrorMessage() As String
    ErrorMessage = mErrorMessage
End Property

Public Property Let ErrorMessage(ByVal value As String)
    mErrorMessage = value
End Property

'@Description("ルールが有効かどうかを取得または設定します")
Public Property Get IsEnabled() As Boolean
    IsEnabled = mIsEnabled
End Property

Public Property Let IsEnabled(ByVal value As Boolean)
    If mIsEnabled <> value Then
        mIsEnabled = value
        If value Then
            RaiseEvent RuleEnabled(mName)
        Else
            RaiseEvent RuleDisabled(mName)
        End If
    End If
End Property

'@Description("優先度を取得または設定します")
Public Property Get Priority() As ValidationPriority
    Priority = mPriority
End Property

Public Property Let Priority(ByVal value As ValidationPriority)
    mPriority = value
End Property

'@Description("依存するルールのコレクションを取得します")
Public Property Get Dependencies() As Collection
    Set Dependencies = mDependencies
End Property

'@Description("バリデーションコンテキストを取得または設定します")
Public Property Get ValidationContext() As IValidationContext
    Set ValidationContext = mValidationContext
End Property

Public Property Set ValidationContext(ByVal value As IValidationContext)
    Set mValidationContext = value
End Property

'@Description("実行時間を取得します")
Public Property Get ExecutionTime() As Double
    ExecutionTime = mExecutionTime
End Property

'@Description("最後の実行結果を取得します")
Public Property Get LastExecutionResult() As Boolean
    LastExecutionResult = mLastExecutionResult
End Property

'@Description("エラー詳細を取得します")
Public Property Get ErrorDetails() As Collection
    Set ErrorDetails = mErrorDetails
End Property

'@Description("依存するルールを追加します")
Public Sub AddDependency(ByVal rule As IValidationRule)
    mDependencies.Add rule
    RaiseEvent DependencyAdded(mName, rule.GetName)
End Sub

'@Description("エラー詳細を追加します")
Protected Sub AddErrorDetail(ByVal detail As String)
    mErrorDetails.Add detail
End Sub

'@Description("バリデーションを実行します")
Public Function Validate(ByVal value As Variant) As Boolean
    If Not mIsEnabled Then
        Validate = True
        Exit Function
    End If
    
    ' 依存するルールのバリデーション
    If Not ValidateDependencies(value) Then
        Validate = False
        Exit Function
    End If
    
    ' 実行時間の計測開始
    Dim startTime As Double
    startTime = Timer
    
    ' 具体的なバリデーション処理（派生クラスでオーバーライド）
    On Error GoTo ErrorHandler
    mLastExecutionResult = ValidateCore(value)
    
    ' 実行時間の計測終了
    mExecutionTime = (Timer - startTime) * 1000 ' ミリ秒に変換
    
    ' イベントの発行
    RaiseEvent RuleExecuted(mName, mLastExecutionResult, mExecutionTime)
    
    Validate = mLastExecutionResult
    Exit Function
    
ErrorHandler:
    mLastExecutionResult = False
    AddErrorDetail "Error in Validate: " & Err.Description
    Resume Next
End Function

'@Description("依存するルールのバリデーションを実行します")
Private Function ValidateDependencies(ByVal value As Variant) As Boolean
    Dim rule As IValidationRule
    
    ValidateDependencies = True
    For Each rule In mDependencies
        If Not rule.Validate(value) Then
            ValidateDependencies = False
            AddErrorDetail "Dependency rule '" & rule.GetName & "' failed"
            Exit Function
        End If
    Next rule
End Function

'@Description("具体的なバリデーション処理を実行します（派生クラスでオーバーライド）")
Protected Function ValidateCore(ByVal value As Variant) As Boolean
    ' 基本クラスでは常にTrueを返す
    ValidateCore = True
End Function

' IValidationRule インターフェースの実装
Private Function IValidationRule_Validate(ByVal value As Variant) As Boolean
    IValidationRule_Validate = Validate(value)
End Function

Private Property Get IValidationRule_Name() As String
    IValidationRule_Name = Name
End Property

Private Property Get IValidationRule_Description() As String
    IValidationRule_Description = Description
End Property

Private Property Get IValidationRule_ErrorMessage() As String
    IValidationRule_ErrorMessage = ErrorMessage
End Property

Private Property Get IValidationRule_IsEnabled() As Boolean
    IValidationRule_IsEnabled = IsEnabled
End Property

Private Property Get IValidationRule_Priority() As ValidationPriority
    IValidationRule_Priority = Priority
End Property

Private Property Get IValidationRule_Dependencies() As Collection
    Set IValidationRule_Dependencies = Dependencies
End Property

Private Sub IValidationRule_AddDependency(ByVal rule As IValidationRule)
    AddDependency rule
End Sub

Private Property Get IValidationRule_ValidationContext() As IValidationContext
    Set IValidationRule_ValidationContext = ValidationContext
End Property

Private Property Set IValidationRule_ValidationContext(ByVal value As IValidationContext)
    Set ValidationContext = value
End Property

'@Description("ルールの状態をJSONフォーマットで取得します")
Public Function ToJSON() As String
    Dim json As String
    json = "{"
    json = json & """Name"": """ & JsonEscape(mName) & ""","
    json = json & """Description"": """ & JsonEscape(mDescription) & ""","
    json = json & """IsEnabled"": " & LCase(CStr(mIsEnabled)) & ","
    json = json & """Priority"": " & mPriority & ","
    json = json & """ExecutionTime"": " & mExecutionTime & ","
    json = json & """LastExecutionResult"": " & LCase(CStr(mLastExecutionResult)) & ","
    json = json & """ErrorCount"": " & mErrorDetails.Count
    json = json & "}"
    ToJSON = json
End Function

'@Description("JSON文字列のエスケープ処理を行います")
Private Function JsonEscape(ByVal text As String) As String
    JsonEscape = Replace(Replace(Replace(text, "\", "\\"), """", "\"""), vbCrLf, "\n")
End Function