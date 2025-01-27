VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AsyncValidationImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

Implements IAsyncValidation

'@Description("非同期バリデーションの実装クラス")

' メンバー変数
Private mValidationQueue As Collection        ' バリデーションキュー
Private mActiveValidations As Collection     ' アクティブなバリデーション
Private mCompletedValidations As Collection  ' 完了したバリデーション
Private mQueueStatus As ValidationQueueStatus ' キューの状態
Private mCacheStrategy As IValidationCacheStrategy ' キャッシュ戦略
Private mSchedulingStrategy As IValidationSchedulingStrategy ' スケジューリング戦略
Private mResourceLimits As Dictionary        ' リソース制限
Private mMetrics As ValidationMetrics        ' メトリクス
Private mIsRunning As Boolean                ' 実行中フラグ
Private mMaxConcurrentValidations As Long    ' 最大同時実行数
Private mDefaultTimeout As Long              ' デフォルトのタイムアウト時間（ミリ秒）

' イベント
Public Event ValidationStarted(ByVal validationId As String, ByVal context As IValidationContext, ByVal executionPlan As ValidationExecutionPlan)
Public Event ValidationCompleted(ByVal validationId As String, ByVal success As Boolean, ByVal results As Collection, ByVal metrics As ValidationMetrics)
Public Event ValidationCancelled(ByVal validationId As String, ByVal reason As String)
Public Event ValidationQueued(ByVal validationId As String, ByVal queuePosition As Long, ByVal estimatedWaitTime As Long)
Public Event ValidationError(ByVal validationId As String, ByVal errorMessage As String, ByVal errorDetails As Collection)
Public Event QueueStatusChanged(ByVal status As ValidationQueueStatus)
Public Event ResourceLimitExceeded(ByVal resourceType As String, ByVal currentValue As Double, ByVal limit As Double)
Public Event PerformanceAlert(ByVal metricName As String, ByVal currentValue As Double, ByVal threshold As Double)

Private Sub Class_Initialize()
    Set mValidationQueue = New Collection
    Set mActiveValidations = New Collection
    Set mCompletedValidations = New Collection
    Set mQueueStatus = New ValidationQueueStatus
    Set mResourceLimits = New Dictionary
    Set mMetrics = New ValidationMetrics
    mIsRunning = False
    mMaxConcurrentValidations = 5 ' デフォルト値
    mDefaultTimeout = 30000 ' 30秒
End Sub

'@Description("非同期バリデーションを開始します")
Private Function IAsyncValidation_StartAsyncValidation(ByVal value As Variant, ByVal context As IValidationContext, Optional ByVal priority As ValidationPriority = vpNormal, Optional ByVal options As AsyncValidationOptions) As String
    ' バリデーションIDの生成
    Dim validationId As String
    validationId = GenerateValidationId
    
    ' バリデーション情報の作成
    Dim validationInfo As New Dictionary
    validationInfo.Add "ID", validationId
    validationInfo.Add "Value", value
    Set validationInfo.Add("Context", context)
    validationInfo.Add "Priority", priority
    Set validationInfo.Add("Options", options)
    validationInfo.Add "Status", "Queued"
    validationInfo.Add "QueuedTime", Now
    
    ' キャッシュのチェック
    If Not mCacheStrategy Is Nothing Then
        If mCacheStrategy.IsCached(GetCacheKey(value, context)) Then
            Dim cachedResult As ValidationResult
            Set cachedResult = mCacheStrategy.GetCachedResult(GetCacheKey(value, context))
            mCompletedValidations.Add cachedResult, validationId
            RaiseEvent ValidationCompleted(validationId, cachedResult.Success, cachedResult.Results, mMetrics)
            IAsyncValidation_StartAsyncValidation = validationId
            Exit Function
        End If
    End If
    
    ' キューへの追加
    mValidationQueue.Add validationInfo, validationId
    
    ' キュー状態の更新
    UpdateQueueStatus
    
    ' イベントの発行
    RaiseEvent ValidationQueued(validationId, mValidationQueue.Count, EstimateWaitTime(priority))
    
    ' バリデーション処理の開始（非同期）
    If Not mIsRunning Then
        mIsRunning = True
        ProcessValidationQueue
    End If
    
    IAsyncValidation_StartAsyncValidation = validationId
End Function

'@Description("バッチバリデーションを開始します")
Private Function IAsyncValidation_StartBatchValidation(ByVal values As Collection, ByVal context As IValidationContext, Optional ByVal options As BatchValidationOptions) As String
    ' バッチIDの生成
    Dim batchId As String
    batchId = "BATCH_" & GenerateValidationId
    
    ' 各値に対してバリデーションを開始
    Dim value As Variant
    Dim validationIds As New Collection
    
    For Each value In values
        Dim validationId As String
        validationId = IAsyncValidation_StartAsyncValidation(value, context, vpNormal, Nothing)
        validationIds.Add validationId
    Next value
    
    ' バッチ情報の保存
    Dim batchInfo As New Dictionary
    batchInfo.Add "ID", batchId
    Set batchInfo.Add("ValidationIds", validationIds)
    Set batchInfo.Add("Options", options)
    batchInfo.Add "Status", "InProgress"
    
    IAsyncValidation_StartBatchValidation = batchId
End Function

'@Description("バリデーションをキャンセルします")
Private Sub IAsyncValidation_CancelValidation(ByVal validationId As String, Optional ByVal options As CancellationOptions)
    ' キューからの削除
    RemoveFromQueue validationId
    
    ' アクティブなバリデーションの停止
    StopActiveValidation validationId
    
    ' イベントの発行
    RaiseEvent ValidationCancelled(validationId, "User cancelled")
    
    ' キュー状態の更新
    UpdateQueueStatus
End Sub

'@Description("バリデーション完了を待機します")
Private Function IAsyncValidation_WaitForValidation(ByVal validationId As String, Optional ByVal timeout As Long = 0, Optional ByVal options As WaitOptions) As Boolean
    Dim startTime As Date
    startTime = Now
    
    ' タイムアウト値の設定
    If timeout <= 0 Then
        timeout = mDefaultTimeout
    End If
    
    ' 完了を待機
    Do While Not IsValidationCompleted(validationId)
        DoEvents
        
        ' タイムアウトチェック
        If DateDiff("s", startTime, Now) * 1000 > timeout Then
            IAsyncValidation_WaitForValidation = False
            Exit Function
        End If
    Loop
    
    IAsyncValidation_WaitForValidation = True
End Function

'@Description("すべてのバリデーション完了を待機します")
Private Function IAsyncValidation_WaitForAll(Optional ByVal timeout As Long = 0, Optional ByVal options As WaitOptions) As Boolean
    Dim startTime As Date
    startTime = Now
    
    ' タイムアウト値の設定
    If timeout <= 0 Then
        timeout = mDefaultTimeout
    End If
    
    ' すべての完了を待機
    Do While mValidationQueue.Count > 0 Or mActiveValidations.Count > 0
        DoEvents
        
        ' タイムアウトチェック
        If DateDiff("s", startTime, Now) * 1000 > timeout Then
            IAsyncValidation_WaitForAll = False
            Exit Function
        End If
    Loop
    
    IAsyncValidation_WaitForAll = True
End Function

'@Description("バリデーション結果を取得します")
Private Function IAsyncValidation_TryGetResults(ByVal validationId As String, Optional ByVal options As ResultOptions) As ValidationResult
    ' 完了したバリデーションから結果を取得
    Dim result As ValidationResult
    On Error Resume Next
    Set result = mCompletedValidations(validationId)
    On Error GoTo 0
    
    If Not result Is Nothing Then
        Set IAsyncValidation_TryGetResults = result
    End If
End Function

'@Description("キューの状態を取得します")
Private Function IAsyncValidation_GetQueueStatus() As ValidationQueueStatus
    Set IAsyncValidation_GetQueueStatus = mQueueStatus
End Function

'@Description("バリデーションの優先度を変更します")
Private Sub IAsyncValidation_ChangePriority(ByVal validationId As String, ByVal newPriority As ValidationPriority, Optional ByVal options As PriorityOptions)
    ' キュー内のバリデーションの優先度を変更
    Dim validationInfo As Dictionary
    On Error Resume Next
    For Each validationInfo In mValidationQueue
        If validationInfo("ID") = validationId Then
            validationInfo("Priority") = newPriority
            Exit For
        End If
    Next validationInfo
    On Error GoTo 0
    
    ' キューの再構築（必要に応じて）
    If Not options Is Nothing Then
        If options.ReorderQueue Then
            ReorderQueue
        End If
    End If
End Sub

'@Description("リソース制限を設定します")
Private Sub IAsyncValidation_SetResourceLimit(ByVal resourceType As String, ByVal limit As Double, Optional ByVal options As ResourceLimitOptions)
    mResourceLimits(resourceType) = limit
End Sub

'@Description("キャッシュ戦略を設定します")
Private Sub IAsyncValidation_SetCacheStrategy(ByVal strategy As IValidationCacheStrategy)
    Set mCacheStrategy = strategy
End Sub

'@Description("スケジューリング戦略を設定します")
Private Sub IAsyncValidation_SetSchedulingStrategy(ByVal strategy As IValidationSchedulingStrategy)
    Set mSchedulingStrategy = strategy
End Sub

'@Description("パフォーマンスメトリクスを取得します")
Private Function IAsyncValidation_GetPerformanceMetrics() As ValidationMetrics
    Set IAsyncValidation_GetPerformanceMetrics = mMetrics
End Function

'@Description("バリデーションキューを処理します")
Private Sub ProcessValidationQueue()
    Do While mValidationQueue.Count > 0 And mActiveValidations.Count < mMaxConcurrentValidations
        ' 次のバリデーションを取得
        Dim nextValidation As Dictionary
        Set nextValidation = GetNextValidation
        
        If Not nextValidation Is Nothing Then
            ' アクティブなバリデーションに移動
            mActiveValidations.Add nextValidation
            RemoveFromQueue nextValidation("ID")
            
            ' バリデーションの実行
            ExecuteValidation nextValidation
        End If
        
        DoEvents
    Loop
    
    mIsRunning = (mValidationQueue.Count > 0)
End Sub

'@Description("次のバリデーションを取得します")
Private Function GetNextValidation() As Dictionary
    If mSchedulingStrategy Is Nothing Then
        ' デフォルトのスケジューリング（優先度ベース）
        Dim highestPriority As ValidationPriority
        highestPriority = vpLowest
        
        Dim validationInfo As Dictionary
        Dim selectedValidation As Dictionary
        
        For Each validationInfo In mValidationQueue
            If validationInfo("Priority") > highestPriority Then
                highestPriority = validationInfo("Priority")
                Set selectedValidation = validationInfo
            End If
        Next validationInfo
        
        Set GetNextValidation = selectedValidation
    Else
        ' カスタムスケジューリング戦略を使用
        Dim nextId As String
        nextId = mSchedulingStrategy.GetNextValidation
        
        For Each validationInfo In mValidationQueue
            If validationInfo("ID") = nextId Then
                Set GetNextValidation = validationInfo
                Exit Function
            End If
        Next validationInfo
    End If
End Function

'@Description("バリデーションを実行します")
Private Sub ExecuteValidation(ByVal validationInfo As Dictionary)
    ' 実行計画の作成
    Dim executionPlan As New ValidationExecutionPlan
    
    ' コンテキストからルールを取得
    Dim context As IValidationContext
    Set context = validationInfo("Context")
    
    ' バリデーションの実行
    Dim result As New ValidationResult
    result.ValidationId = validationInfo("ID")
    result.StartTime = Now
    
    ' イベントの発行
    RaiseEvent ValidationStarted(validationInfo("ID"), context, executionPlan)
    
    On Error GoTo ErrorHandler
    
    ' ルールの実行
    Dim rule As IValidationRule
    For Each rule In context.GetValidationRules
        Dim ruleResult As ValidationResult
        Set ruleResult = ExecuteRule(rule, validationInfo("Value"))
        result.AddResult ruleResult
        
        If Not ruleResult.Success And validationInfo("Options").StopOnFirstFailure Then
            Exit For
        End If
    Next rule
    
    ' 結果の設定
    result.EndTime = Now
    result.ExecutionTime = DateDiff("s", result.StartTime, result.EndTime) * 1000
    result.Success = (result.ErrorDetails.Count = 0)
    
    ' キャッシュの更新
    If Not mCacheStrategy Is Nothing Then
        mCacheStrategy.StoreResult GetCacheKey(validationInfo("Value"), context), result
    End If
    
    ' メトリクスの更新
    UpdateMetrics result
    
    ' 完了処理
    CompleteValidation validationInfo("ID"), result
    
    Exit Sub

ErrorHandler:
    ' エラー処理
    result.Success = False
    result.ErrorMessage = "Validation error: " & Err.Description
    RaiseEvent ValidationError(validationInfo("ID"), Err.Description, result.ErrorDetails)
    CompleteValidation validationInfo("ID"), result
End Sub

'@Description("個別のルールを実行します")
Private Function ExecuteRule(ByVal rule As IValidationRule, ByVal value As Variant) As ValidationResult
    Dim result As New ValidationResult
    
    result.StartTime = Now
    result.Success = rule.Validate(value)
    result.EndTime = Now
    result.ExecutionTime = DateDiff("s", result.StartTime, result.EndTime) * 1000
    
    If Not result.Success Then
        result.ErrorMessage = rule.ErrorMessage
    End If
    
    Set ExecuteRule = result
End Function

'@Description("バリデーションを完了します")
Private Sub CompleteValidation(ByVal validationId As String, ByVal result As ValidationResult)
    ' アクティブなバリデーションから削除
    RemoveFromActive validationId
    
    ' 完了したバリデーションに追加
    mCompletedValidations.Add result, validationId
    
    ' イベントの発行
    RaiseEvent ValidationCompleted(validationId, result.Success, result.Results, mMetrics)
    
    ' キュー状態の更新
    UpdateQueueStatus
    
    ' キューの処理を継続
    ProcessValidationQueue
End Sub

'@Description("キューからバリデーションを削除します")
Private Sub RemoveFromQueue(ByVal validationId As String)
    Dim i As Long
    For i = mValidationQueue.Count To 1 Step -1
        If mValidationQueue(i)("ID") = validationId Then
            mValidationQueue.Remove i
            Exit Sub
        End If
    Next i
End Sub

'@Description("アクティブなバリデーションから削除します")
Private Sub RemoveFromActive(ByVal validationId As String)
    Dim i As Long
    For i = mActiveValidations.Count To 1 Step -1
        If mActiveValidations(i)("ID") = validationId Then
            mActiveValidations.Remove i
            Exit Sub
        End If
    Next i
End Sub

'@Description("アクティブなバリデーションを停止します")
Private Sub StopActiveValidation(ByVal validationId As String)
    Dim i As Long
    For i = mActiveValidations.Count To 1 Step -1
        If mActiveValidations(i)("ID") = validationId Then
            ' 必要に応じてクリーンアップ処理を実行
            RemoveFromActive validationId
            Exit Sub
        End If
    Next i
End Sub

'@Description("キューの状態を更新します")
Private Sub UpdateQueueStatus()
    With mQueueStatus
        .QueueLength = mValidationQueue.Count
        .ActiveJobs = mActiveValidations.Count
        .PendingJobs = mValidationQueue.Count
        .LastUpdateTime = Now
        
        ' 平均待機時間の計算
        If mValidationQueue.Count > 0 Then
            Dim totalWaitTime As Double
            Dim validationInfo As Dictionary
            For Each validationInfo In mValidationQueue
                totalWaitTime = totalWaitTime + DateDiff("s", validationInfo("QueuedTime"), Now)
            Next validationInfo
            .AverageWaitingTime = totalWaitTime / mValidationQueue.Count * 1000 ' ミリ秒に変換
        End If
    End With
    
    RaiseEvent QueueStatusChanged(mQueueStatus)
End Sub

'@Description("メトリクスを更新します")
Private Sub UpdateMetrics(ByVal result As ValidationResult)
    With mMetrics
        .TotalExecutionTime = .TotalExecutionTime + result.ExecutionTime
        .RuleCount = .RuleCount + result.Results.Count
        
        If result.Success Then
            .SuccessCount = .SuccessCount + 1
        Else
            .FailureCount = .FailureCount + 1
        End If
        
        If result.ExecutionTime > .MaxRuleExecutionTime Then
            .MaxRuleExecutionTime = result.ExecutionTime
        End If
        If result.ExecutionTime < .MinRuleExecutionTime Then
            .MinRuleExecutionTime = result.ExecutionTime
        End If
    End With
End Sub

'@Description("キャッシュキーを生成します")
Private Function GetCacheKey(ByVal value As Variant, ByVal context As IValidationContext) As String
    ' 単純な実装 - 実際の要件に応じて改善が必要
    GetCacheKey = TypeName(value) & "_" & context.GetContextId
End Function

'@Description("待機時間を推定します")
Private Function EstimateWaitTime(ByVal priority As ValidationPriority) As Long
    ' 単純な推定 - 実際の要件に応じて改善が必要
    EstimateWaitTime = mValidationQueue.Count * 1000 ' 1検証あたり1秒と仮定
End Function

'@Description("キューを再構築します")
Private Sub ReorderQueue()
    ' 優先度に基づいてキューを並べ替え
    Dim orderedQueue As New Collection
    Dim priorities(vpHighest To vpLowest) As Collection
    Dim i As ValidationPriority
    
    ' 優先度ごとのコレクションを初期化
    For i = vpHighest To vpLowest
        Set priorities(i) = New Collection
    Next i
    
    ' バリデーションを優先度ごとに分類
    Dim validationInfo As Dictionary
    For Each validationInfo In mValidationQueue
        priorities(validationInfo("Priority")).Add validationInfo
    Next validationInfo
    
    ' 優先度順に新しいキューを構築
    Set mValidationQueue = New Collection
    For i = vpHighest To vpLowest
        For Each validationInfo In priorities(i)
            mValidationQueue.Add validationInfo
        Next validationInfo
    Next i
End Sub

'@Description("バリデーションが完了しているかどうかを確認します")
Private Function IsValidationCompleted(ByVal validationId As String) As Boolean
    On Error Resume Next
    Dim result As ValidationResult
    Set result = mCompletedValidations(validationId)
    IsValidationCompleted = Not result Is Nothing
    On Error GoTo 0
End Function

'@Description("一意のバリデーションIDを生成します")
Private Function GenerateValidationId() As String
    GenerateValidationId = "VAL_" & Format(Now, "yyyymmddhhnnss") & "_" & Format(Timer * 1000, "000")
End Function