VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationExecutionPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("バリデーションの実行計画を管理するクラス")

' メンバー変数
Private mExecutionOrder As Collection       ' ルールの実行順序
Private mDependencies As Collection        ' ルール間の依存関係
Private mParallelizableRules As Collection ' 並列実行可能なルール
Private mPlanId As String                  ' 実行計画ID
Private mCreatedAt As Date                 ' 作成日時
Private mEstimatedDuration As Long         ' 推定実行時間（ミリ秒）
Private mResourceRequirements As Collection ' リソース要件
Private mOptimizationHints As Collection   ' 最適化ヒント
Private mValidationGroups As Collection    ' バリデーショングループ

' イベント
Public Event PlanCreated(ByVal planId As String)
Public Event PlanOptimized(ByVal optimizationType As String, ByVal improvement As Double)
Public Event DependencyAdded(ByVal dependentRule As String, ByVal dependencyRule As String)
Public Event GroupCreated(ByVal groupName As String, ByVal ruleCount As Long)
Public Event ResourceRequirementUpdated(ByVal resourceType As String, ByVal requirement As Double)

Private Sub Class_Initialize()
    Set mExecutionOrder = New Collection
    Set mDependencies = New Collection
    Set mParallelizableRules = New Collection
    Set mResourceRequirements = New Collection
    Set mOptimizationHints = New Collection
    Set mValidationGroups = New Collection
    mPlanId = GenerateUniqueId
    mCreatedAt = Now
    mEstimatedDuration = 0
End Sub

'@Description("実行順序のコレクションを取得します")
Public Property Get ExecutionOrder() As Collection
    Set ExecutionOrder = mExecutionOrder
End Property

'@Description("依存関係のコレクションを取得します")
Public Property Get Dependencies() As Collection
    Set Dependencies = mDependencies
End Property

'@Description("並列実行可能なルールのコレクションを取得します")
Public Property Get ParallelizableRules() As Collection
    Set ParallelizableRules = mParallelizableRules
End Property

'@Description("実行計画IDを取得します")
Public Property Get PlanId() As String
    PlanId = mPlanId
End Property

'@Description("作成日時を取得します")
Public Property Get CreatedAt() As Date
    CreatedAt = mCreatedAt
End Property

'@Description("推定実行時間を取得または設定します")
Public Property Get EstimatedDuration() As Long
    EstimatedDuration = mEstimatedDuration
End Property

Public Property Let EstimatedDuration(ByVal value As Long)
    mEstimatedDuration = value
End Property

'@Description("実行順序にルールを追加します")
'@Param("rule: 追加するルール情報")
Public Sub AddExecutionRule(ByVal rule As RuleExecutionInfo)
    mExecutionOrder.Add rule
    OptimizeExecutionOrder
End Sub

'@Description("依存関係を追加します")
'@Param("dependentRule: 依存するルール")
'@Param("dependencyRule: 依存されるルール")
'@Param("dependencyType: 依存関係の種類")
Public Sub AddDependency(ByVal dependentRule As IValidationRule, ByVal dependencyRule As IValidationRule, ByVal dependencyType As DependencyType)
    Dim dependency As New DependencyInfo
    Set dependency.DependentRule = dependentRule
    Set dependency.DependencyRule = dependencyRule
    dependency.DependencyType = dependencyType
    mDependencies.Add dependency
    RaiseEvent DependencyAdded(dependentRule.GetDescription, dependencyRule.GetDescription)
    OptimizeExecutionOrder
End Sub

'@Description("並列実行可能なルールを追加します")
'@Param("rule: 追加するルール")
Public Sub AddParallelizableRule(ByVal rule As IValidationRule)
    mParallelizableRules.Add rule
    OptimizeExecutionOrder
End Sub

'@Description("バリデーショングループを作成します")
'@Param("groupName: グループ名")
'@Param("rules: グループに含めるルールのコレクション")
Public Sub CreateValidationGroup(ByVal groupName As String, ByVal rules As Collection)
    Dim group As New ValidationGroup
    group.GroupName = groupName
    Set group.Rules = rules
    mValidationGroups.Add group
    RaiseEvent GroupCreated(groupName, rules.Count)
End Sub

'@Description("リソース要件を追加します")
'@Param("resourceType: リソースの種類")
'@Param("requirement: 要件値")
Public Sub AddResourceRequirement(ByVal resourceType As String, ByVal requirement As Double)
    mResourceRequirements.Add requirement, resourceType
    RaiseEvent ResourceRequirementUpdated(resourceType, requirement)
End Sub

'@Description("最適化ヒントを追加します")
'@Param("hintType: ヒントの種類")
'@Param("hintValue: ヒントの値")
Public Sub AddOptimizationHint(ByVal hintType As String, ByVal hintValue As Variant)
    mOptimizationHints.Add hintValue, hintType
End Sub

'@Description("実行順序を最適化します")
Private Sub OptimizeExecutionOrder()
    ' 依存関係に基づいて実行順序を最適化
    Dim optimizedOrder As New Collection
    Dim visited As New Dictionary
    Dim improvement As Double
    
    ' トポロジカルソートを使用して最適な順序を決定
    Dim rule As RuleExecutionInfo
    For Each rule In mExecutionOrder
        If Not visited.Exists(rule.Rule.GetDescription) Then
            TopologicalSort rule, visited, optimizedOrder
        End If
    Next rule
    
    ' 並列実行可能なルールをグループ化
    GroupParallelizableRules optimizedOrder
    
    ' 最適化による改善度を計算
    improvement = CalculateOptimizationImprovement(mExecutionOrder, optimizedOrder)
    
    ' 最適化結果を適用
    Set mExecutionOrder = optimizedOrder
    RaiseEvent PlanOptimized("ExecutionOrder", improvement)
End Sub

'@Description("トポロジカルソートを実行します")
Private Sub TopologicalSort(ByVal rule As RuleExecutionInfo, ByRef visited As Dictionary, ByRef result As Collection)
    visited.Add rule.Rule.GetDescription, True
    
    ' 依存関係を持つルールを先に処理
    Dim dependency As DependencyInfo
    For Each dependency In mDependencies
        If dependency.DependentRule.GetDescription = rule.Rule.GetDescription Then
            If Not visited.Exists(dependency.DependencyRule.GetDescription) Then
                Dim dependencyRule As RuleExecutionInfo
                Set dependencyRule = FindRuleInfo(dependency.DependencyRule)
                If Not dependencyRule Is Nothing Then
                    TopologicalSort dependencyRule, visited, result
                End If
            End If
        End If
    Next dependency
    
    result.Add rule
End Sub

'@Description("並列実行可能なルールをグループ化します")
Private Sub GroupParallelizableRules(ByRef optimizedOrder As Collection)
    Dim parallelGroup As New Collection
    Dim rule As RuleExecutionInfo
    
    For Each rule In optimizedOrder
        If IsParallelizable(rule.Rule) Then
            parallelGroup.Add rule
        Else
            If parallelGroup.Count > 0 Then
                CreateParallelGroup parallelGroup
                Set parallelGroup = New Collection
            End If
        End If
    Next rule
    
    If parallelGroup.Count > 0 Then
        CreateParallelGroup parallelGroup
    End If
End Sub

'@Description("ルールが並列実行可能かどうかを判定します")
Private Function IsParallelizable(ByVal rule As IValidationRule) As Boolean
    Dim parallelRule As IValidationRule
    For Each parallelRule In mParallelizableRules
        If parallelRule.GetDescription = rule.GetDescription Then
            IsParallelizable = True
            Exit Function
        End If
    Next parallelRule
    IsParallelizable = False
End Function

'@Description("並列実行グループを作成します")
Private Sub CreateParallelGroup(ByVal rules As Collection)
    Dim groupName As String
    groupName = "ParallelGroup_" & mValidationGroups.Count + 1
    CreateValidationGroup groupName, rules
End Sub

'@Description("最適化による改善度を計算します")
Private Function CalculateOptimizationImprovement(ByVal originalOrder As Collection, ByVal optimizedOrder As Collection) As Double
    ' 単純な改善度計算の例：並列実行可能なグループ数の増加率
    Dim originalParallelGroups As Long
    Dim optimizedParallelGroups As Long
    
    originalParallelGroups = CountParallelGroups(originalOrder)
    optimizedParallelGroups = CountParallelGroups(optimizedOrder)
    
    If originalParallelGroups = 0 Then
        CalculateOptimizationImprovement = 0
    Else
        CalculateOptimizationImprovement = (optimizedParallelGroups - originalParallelGroups) / originalParallelGroups
    End If
End Function

'@Description("並列実行グループ数をカウントします")
Private Function CountParallelGroups(ByVal executionOrder As Collection) As Long
    Dim count As Long
    Dim rule As RuleExecutionInfo
    Dim inParallelGroup As Boolean
    
    count = 0
    inParallelGroup = False
    
    For Each rule In executionOrder
        If IsParallelizable(rule.Rule) Then
            If Not inParallelGroup Then
                count = count + 1
                inParallelGroup = True
            End If
        Else
            inParallelGroup = False
        End If
    Next rule
    
    CountParallelGroups = count
End Function

'@Description("ルール情報を検索します")
Private Function FindRuleInfo(ByVal rule As IValidationRule) As RuleExecutionInfo
    Dim ruleInfo As RuleExecutionInfo
    For Each ruleInfo In mExecutionOrder
        If ruleInfo.Rule.GetDescription = rule.GetDescription Then
            Set FindRuleInfo = ruleInfo
            Exit Function
        End If
    Next ruleInfo
    Set FindRuleInfo = Nothing
End Function

'@Description("一意のIDを生成します")
Private Function GenerateUniqueId() As String
    GenerateUniqueId = "PLAN_" & Format(Now, "yyyymmddhhnnss") & "_" & Format(Timer * 1000, "000")
End Function

'@Description("実行計画をJSONフォーマットで取得します")
Public Function ToJSON() As String
    Dim json As String
    json = "{"
    json = json & """PlanId"": """ & mPlanId & ""","
    json = json & """CreatedAt"": """ & Format(mCreatedAt, "yyyy-mm-dd hh:nn:ss") & ""","
    json = json & """EstimatedDuration"": " & mEstimatedDuration & ","
    json = json & """RuleCount"": " & mExecutionOrder.Count & ","
    json = json & """ParallelizableRuleCount"": " & mParallelizableRules.Count & ","
    json = json & """GroupCount"": " & mValidationGroups.Count
    json = json & "}"
    ToJSON = json
End Function