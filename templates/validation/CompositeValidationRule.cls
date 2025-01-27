VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CompositeValidationRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

Implements ICompositeValidationRule
Implements IValidationRule

'@Description("複合バリデーションルール")

Private WithEvents BaseRule As ValidationRuleBase
Private mRules As Collection           ' バリデーションルールのコレクション
Private mStopOnFirstFailure As Boolean ' 最初の失敗で停止するかどうか
Private mParallelExecution As Boolean  ' 並列実行を許可するかどうか
Private mValidationGroups As Collection ' バリデーショングループ
Private mExecutionPlan As ValidationExecutionPlan ' 実行計画

' イベント
Public Event RuleAdded(ByVal rule As IValidationRule)
Public Event RuleRemoved(ByVal rule As IValidationRule)
Public Event GroupCreated(ByVal groupName As String)
Public Event ValidationStarted(ByVal context As IValidationContext)
Public Event ValidationCompleted(ByVal success As Boolean, ByVal results As Collection)

Private Sub Class_Initialize()
    Set BaseRule = New ValidationRuleBase
    BaseRule.Name = "CompositeValidationRule"
    BaseRule.Description = "複数のバリデーションルールを組み合わせて実行します"
    Set mRules = New Collection
    Set mValidationGroups = New Collection
    Set mExecutionPlan = New ValidationExecutionPlan
    mStopOnFirstFailure = True
    mParallelExecution = False
End Sub

'@Description("最初の失敗で停止するかどうかを取得または設定します")
Public Property Get StopOnFirstFailure() As Boolean
    StopOnFirstFailure = mStopOnFirstFailure
End Property

Public Property Let StopOnFirstFailure(ByVal value As Boolean)
    mStopOnFirstFailure = value
End Property

'@Description("並列実行を許可するかどうかを取得または設定します")
Public Property Get ParallelExecution() As Boolean
    ParallelExecution = mParallelExecution
End Property

Public Property Let ParallelExecution(ByVal value As Boolean)
    mParallelExecution = value
End Property

'@Description("バリデーションルールを追加します")
Public Sub AddRule(ByVal rule As IValidationRule)
    mRules.Add rule
    RaiseEvent RuleAdded(rule)
    UpdateExecutionPlan
End Sub

'@Description("バリデーションルールを削除します")
Public Sub RemoveRule(ByVal rule As IValidationRule)
    Dim i As Long
    For i = 1 To mRules.Count
        If mRules(i) Is rule Then
            mRules.Remove i
            RaiseEvent RuleRemoved(rule)
            UpdateExecutionPlan
            Exit Sub
        End If
    Next i
End Sub

'@Description("バリデーショングループを作成します")
Public Sub CreateValidationGroup(ByVal groupName As String, ByVal rules As Collection)
    Dim group As New ValidationGroup
    group.GroupName = groupName
    Set group.Rules = rules
    mValidationGroups.Add group
    RaiseEvent GroupCreated(groupName)
    UpdateExecutionPlan
End Sub

'@Description("実行計画を更新します")
Private Sub UpdateExecutionPlan()
    Set mExecutionPlan = New ValidationExecutionPlan
    
    ' ルールの実行順序を設定
    Dim rule As IValidationRule
    For Each rule In mRules
        Dim ruleInfo As New RuleExecutionInfo
        Set ruleInfo.Rule = rule
        ruleInfo.Priority = rule.Priority
        mExecutionPlan.AddExecutionRule ruleInfo
    Next rule
    
    ' 依存関係を設定
    For Each rule In mRules
        Dim dependency As IValidationRule
        For Each dependency In rule.Dependencies
            mExecutionPlan.AddDependency rule, dependency, dtRequired
        Next dependency
    Next rule
    
    ' 並列実行可能なルールを設定
    If mParallelExecution Then
        For Each rule In mRules
            If CanExecuteInParallel(rule) Then
                mExecutionPlan.AddParallelizableRule rule
            End If
        Next rule
    End If
End Sub

'@Description("ルールが並列実行可能かどうかを判定します")
Private Function CanExecuteInParallel(ByVal rule As IValidationRule) As Boolean
    ' 依存関係がないルールは並列実行可能
    CanExecuteInParallel = (rule.Dependencies.Count = 0)
End Function

'@Description("バリデーションを実行します")
Public Function Validate(ByVal value As Variant) As Boolean
    ' バリデーション開始イベントを発行
    RaiseEvent ValidationStarted(BaseRule.ValidationContext)
    
    ' 実行計画に基づいてバリデーションを実行
    Dim results As New Collection
    Dim success As Boolean
    success = True
    
    Dim ruleInfo As RuleExecutionInfo
    For Each ruleInfo In mExecutionPlan.ExecutionOrder
        ' ルールが無効な場合はスキップ
        If Not ruleInfo.Rule.IsEnabled Then
            Continue For
        End If
        
        ' 依存するルールのチェック
        If Not ValidateDependencies(ruleInfo.Rule, value) Then
            success = False
            If mStopOnFirstFailure Then
                Exit For
            End If
            Continue For
        End If
        
        ' バリデーションの実行
        Dim result As ValidationResult
        Set result = ExecuteValidation(ruleInfo.Rule, value)
        results.Add result
        
        If Not result.Success Then
            success = False
            If mStopOnFirstFailure Then
                Exit For
            End If
        End If
    Next ruleInfo
    
    ' バリデーション完了イベントを発行
    RaiseEvent ValidationCompleted(success, results)
    
    Validate = success
End Function

'@Description("依存するルールのバリデーションを実行します")
Private Function ValidateDependencies(ByVal rule As IValidationRule, ByVal value As Variant) As Boolean
    ValidateDependencies = True
    
    Dim dependency As IValidationRule
    For Each dependency In rule.Dependencies
        If Not dependency.Validate(value) Then
            ValidateDependencies = False
            Exit Function
        End If
    Next dependency
End Function

'@Description("個別のバリデーションを実行します")
Private Function ExecuteValidation(ByVal rule As IValidationRule, ByVal value As Variant) As ValidationResult
    Dim result As New ValidationResult
    
    ' バリデーション実行前の設定
    result.ValidationId = GenerateValidationId
    result.StartTime = Now
    
    ' バリデーションの実行
    result.Success = rule.Validate(value)
    
    ' バリデーション実行後の設定
    result.EndTime = Now
    result.ExecutionTime = DateDiff("s", result.StartTime, result.EndTime) * 1000 ' ミリ秒に変換
    
    If Not result.Success Then
        result.ErrorMessage = rule.ErrorMessage
    End If
    
    Set ExecuteValidation = result
End Function

'@Description("一意のバリデーションIDを生成します")
Private Function GenerateValidationId() As String
    GenerateValidationId = "VAL_" & Format(Now, "yyyymmddhhnnss") & "_" & Format(Timer * 1000, "000")
End Function

' IValidationRule インターフェースの実装
Private Function IValidationRule_Validate(ByVal value As Variant) As Boolean
    IValidationRule_Validate = Validate(value)
End Function

Private Property Get IValidationRule_Name() As String
    IValidationRule_Name = BaseRule.Name
End Property

Private Property Get IValidationRule_Description() As String
    IValidationRule_Description = BaseRule.Description
End Property

Private Property Get IValidationRule_ErrorMessage() As String
    IValidationRule_ErrorMessage = BaseRule.ErrorMessage
End Property

Private Property Get IValidationRule_IsEnabled() As Boolean
    IValidationRule_IsEnabled = BaseRule.IsEnabled
End Property

Private Property Get IValidationRule_Priority() As ValidationPriority
    IValidationRule_Priority = BaseRule.Priority
End Property

Private Property Get IValidationRule_Dependencies() As Collection
    Set IValidationRule_Dependencies = BaseRule.Dependencies
End Property

Private Sub IValidationRule_AddDependency(ByVal rule As IValidationRule)
    BaseRule.AddDependency rule
End Sub

Private Property Get IValidationRule_ValidationContext() As IValidationContext
    Set IValidationRule_ValidationContext = BaseRule.ValidationContext
End Property

Private Property Set IValidationRule_ValidationContext(ByVal value As IValidationContext)
    Set BaseRule.ValidationContext = value
End Property

' ICompositeValidationRule インターフェースの実装
Private Sub ICompositeValidationRule_AddRule(ByVal rule As IValidationRule)
    AddRule rule
End Sub

Private Sub ICompositeValidationRule_RemoveRule(ByVal rule As IValidationRule)
    RemoveRule rule
End Sub

Private Property Get ICompositeValidationRule_Rules() As Collection
    Set ICompositeValidationRule_Rules = mRules
End Property

Private Sub ICompositeValidationRule_CreateValidationGroup(ByVal groupName As String, ByVal rules As Collection)
    CreateValidationGroup groupName, rules
End Sub

Private Property Get ICompositeValidationRule_ValidationGroups() As Collection
    Set ICompositeValidationRule_ValidationGroups = mValidationGroups
End Property

Private Property Get ICompositeValidationRule_ExecutionPlan() As ValidationExecutionPlan
    Set ICompositeValidationRule_ExecutionPlan = mExecutionPlan
End Property

Private Property Get ICompositeValidationRule_StopOnFirstFailure() As Boolean
    ICompositeValidationRule_StopOnFirstFailure = mStopOnFirstFailure
End Property

Private Property Let ICompositeValidationRule_StopOnFirstFailure(ByVal value As Boolean)
    StopOnFirstFailure = value
End Property

Private Property Get ICompositeValidationRule_ParallelExecution() As Boolean
    ICompositeValidationRule_ParallelExecution = mParallelExecution
End Property

Private Property Let ICompositeValidationRule_ParallelExecution(ByVal value As Boolean)
    ParallelExecution = value
End Property