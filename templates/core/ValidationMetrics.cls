VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationMetrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("バリデーション実行のパフォーマンスメトリクスを管理するクラス")

' メンバー変数
Private mTotalExecutionTime As Double    ' 総実行時間（ミリ秒）
Private mRuleCount As Long              ' 実行されたルールの総数
Private mSuccessCount As Long           ' 成功したルールの数
Private mFailureCount As Long           ' 失敗したルールの数
Private mErrorRate As Double            ' エラー発生率
Private mAverageRuleExecutionTime As Double ' 平均ルール実行時間
Private mMaxRuleExecutionTime As Double  ' 最大ルール実行時間
Private mMinRuleExecutionTime As Double  ' 最小ルール実行時間
Private mCacheHitRate As Double         ' キャッシュヒット率
Private mMemoryUsage As Double          ' メモリ使用量（バイト）
Private mRuleMetrics As Collection      ' 個々のルールのメトリクス
Private mTimeSeriesData As Collection   ' 時系列データ
Private mResourceUtilization As Collection ' リソース使用率

' イベント
Public Event MetricsUpdated(ByVal metricName As String, ByVal oldValue As Double, ByVal newValue As Double)
Public Event ThresholdExceeded(ByVal metricName As String, ByVal currentValue As Double, ByVal threshold As Double)
Public Event PerformanceAlert(ByVal alertType As String, ByVal details As String)

Private Sub Class_Initialize()
    mTotalExecutionTime = 0
    mRuleCount = 0
    mSuccessCount = 0
    mFailureCount = 0
    mErrorRate = 0
    mAverageRuleExecutionTime = 0
    mMaxRuleExecutionTime = 0
    mMinRuleExecutionTime = CDbl(2147483647) ' 最大値で初期化
    mCacheHitRate = 0
    mMemoryUsage = 0
    Set mRuleMetrics = New Collection
    Set mTimeSeriesData = New Collection
    Set mResourceUtilization = New Collection
End Sub

'@Description("総実行時間を取得または設定します")
Public Property Get TotalExecutionTime() As Double
    TotalExecutionTime = mTotalExecutionTime
End Property

Public Property Let TotalExecutionTime(ByVal value As Double)
    Dim oldValue As Double
    oldValue = mTotalExecutionTime
    mTotalExecutionTime = value
    RaiseEvent MetricsUpdated("TotalExecutionTime", oldValue, value)
    CheckThresholds "TotalExecutionTime", value
End Property

'@Description("実行されたルールの総数を取得または設定します")
Public Property Get RuleCount() As Long
    RuleCount = mRuleCount
End Property

Public Property Let RuleCount(ByVal value As Long)
    Dim oldValue As Long
    oldValue = mRuleCount
    mRuleCount = value
    RaiseEvent MetricsUpdated("RuleCount", oldValue, value)
End Property

'@Description("成功したルールの数を取得または設定します")
Public Property Get SuccessCount() As Long
    SuccessCount = mSuccessCount
End Property

Public Property Let SuccessCount(ByVal value As Long)
    Dim oldValue As Long
    oldValue = mSuccessCount
    mSuccessCount = value
    UpdateErrorRate
    RaiseEvent MetricsUpdated("SuccessCount", oldValue, value)
End Property

'@Description("失敗したルールの数を取得または設定します")
Public Property Get FailureCount() As Long
    FailureCount = mFailureCount
End Property

Public Property Let FailureCount(ByVal value As Long)
    Dim oldValue As Long
    oldValue = mFailureCount
    mFailureCount = value
    UpdateErrorRate
    RaiseEvent MetricsUpdated("FailureCount", oldValue, value)
End Property

'@Description("エラー発生率を取得します")
Public Property Get ErrorRate() As Double
    ErrorRate = mErrorRate
End Property

'@Description("平均ルール実行時間を取得します")
Public Property Get AverageRuleExecutionTime() As Double
    AverageRuleExecutionTime = mAverageRuleExecutionTime
End Property

'@Description("最大ルール実行時間を取得または設定します")
Public Property Get MaxRuleExecutionTime() As Double
    MaxRuleExecutionTime = mMaxRuleExecutionTime
End Property

Public Property Let MaxRuleExecutionTime(ByVal value As Double)
    Dim oldValue As Double
    oldValue = mMaxRuleExecutionTime
    mMaxRuleExecutionTime = value
    RaiseEvent MetricsUpdated("MaxRuleExecutionTime", oldValue, value)
    CheckThresholds "MaxRuleExecutionTime", value
End Property

'@Description("最小ルール実行時間を取得または設定します")
Public Property Get MinRuleExecutionTime() As Double
    MinRuleExecutionTime = mMinRuleExecutionTime
End Property

Public Property Let MinRuleExecutionTime(ByVal value As Double)
    Dim oldValue As Double
    oldValue = mMinRuleExecutionTime
    mMinRuleExecutionTime = value
    RaiseEvent MetricsUpdated("MinRuleExecutionTime", oldValue, value)
End Property

'@Description("キャッシュヒット率を取得または設定します")
Public Property Get CacheHitRate() As Double
    CacheHitRate = mCacheHitRate
End Property

Public Property Let CacheHitRate(ByVal value As Double)
    Dim oldValue As Double
    oldValue = mCacheHitRate
    mCacheHitRate = value
    RaiseEvent MetricsUpdated("CacheHitRate", oldValue, value)
    CheckThresholds "CacheHitRate", value
End Property

'@Description("メモリ使用量を取得または設定します")
Public Property Get MemoryUsage() As Double
    MemoryUsage = mMemoryUsage
End Property

Public Property Let MemoryUsage(ByVal value As Double)
    Dim oldValue As Double
    oldValue = mMemoryUsage
    mMemoryUsage = value
    RaiseEvent MetricsUpdated("MemoryUsage", oldValue, value)
    CheckThresholds "MemoryUsage", value
End Property

'@Description("ルールのメトリクスを追加します")
'@Param("ruleName: ルール名")
'@Param("executionTime: 実行時間")
'@Param("success: 成功フラグ")
Public Sub AddRuleMetric(ByVal ruleName As String, ByVal executionTime As Double, ByVal success As Boolean)
    Dim metric As New Dictionary
    metric.Add "RuleName", ruleName
    metric.Add "ExecutionTime", executionTime
    metric.Add "Success", success
    metric.Add "Timestamp", Now
    
    mRuleMetrics.Add metric
    
    ' 統計の更新
    UpdateExecutionTimeStats executionTime
    If success Then
        SuccessCount = SuccessCount + 1
    Else
        FailureCount = FailureCount + 1
    End If
    
    ' 時系列データの追加
    AddTimeSeriesData ruleName, executionTime, success
End Sub

'@Description("時系列データを追加します")
Private Sub AddTimeSeriesData(ByVal ruleName As String, ByVal executionTime As Double, ByVal success As Boolean)
    Dim timeSeriesPoint As New Dictionary
    timeSeriesPoint.Add "RuleName", ruleName
    timeSeriesPoint.Add "ExecutionTime", executionTime
    timeSeriesPoint.Add "Success", success
    timeSeriesPoint.Add "Timestamp", Now
    
    mTimeSeriesData.Add timeSeriesPoint
End Sub

'@Description("実行時間の統計を更新します")
Private Sub UpdateExecutionTimeStats(ByVal executionTime As Double)
    ' 最大・最小値の更新
    If executionTime > mMaxRuleExecutionTime Then
        MaxRuleExecutionTime = executionTime
    End If
    If executionTime < mMinRuleExecutionTime Then
        MinRuleExecutionTime = executionTime
    End If
    
    ' 平均値の更新
    mAverageRuleExecutionTime = (mAverageRuleExecutionTime * (mRuleCount) + executionTime) / (mRuleCount + 1)
    RuleCount = RuleCount + 1
End Sub

'@Description("エラー率を更新します")
Private Sub UpdateErrorRate()
    If mRuleCount > 0 Then
        mErrorRate = mFailureCount / mRuleCount
        RaiseEvent MetricsUpdated("ErrorRate", 0, mErrorRate)
        CheckThresholds "ErrorRate", mErrorRate
    End If
End Sub

'@Description("リソース使用率を追加します")
'@Param("resourceType: リソースの種類")
'@Param("utilization: 使用率")
Public Sub AddResourceUtilization(ByVal resourceType As String, ByVal utilization As Double)
    Dim resource As New Dictionary
    resource.Add "ResourceType", resourceType
    resource.Add "Utilization", utilization
    resource.Add "Timestamp", Now
    
    mResourceUtilization.Add resource
    CheckThresholds "ResourceUtilization_" & resourceType, utilization
End Sub

'@Description("閾値をチェックします")
Private Sub CheckThresholds(ByVal metricName As String, ByVal value As Double)
    Select Case metricName
        Case "TotalExecutionTime"
            If value > 30000 Then ' 30秒を閾値とする
                RaiseEvent ThresholdExceeded(metricName, value, 30000)
            End If
        
        Case "MaxRuleExecutionTime"
            If value > 5000 Then ' 5秒を閾値とする
                RaiseEvent ThresholdExceeded(metricName, value, 5000)
            End If
            
        Case "ErrorRate"
            If value > 0.1 Then ' 10%を閾値とする
                RaiseEvent ThresholdExceeded(metricName, value, 0.1)
            End If
            
        Case "CacheHitRate"
            If value < 0.5 Then ' 50%を閾値とする
                RaiseEvent ThresholdExceeded(metricName, value, 0.5)
            End If
            
        Case "MemoryUsage"
            If value > 1073741824 Then ' 1GBを閾値とする
                RaiseEvent ThresholdExceeded(metricName, value, 1073741824)
            End If
    End Select
    
    ' リソース使用率の閾値チェック
    If Left(metricName, 19) = "ResourceUtilization_" Then
        If value > 0.8 Then ' 80%を閾値とする
            RaiseEvent ThresholdExceeded(metricName, value, 0.8)
        End If
    End If
End Sub

'@Description("メトリクスをJSONフォーマットで取得します")
Public Function ToJSON() As String
    Dim json As String
    json = "{"
    json = json & """TotalExecutionTime"": " & mTotalExecutionTime & ","
    json = json & """RuleCount"": " & mRuleCount & ","
    json = json & """SuccessCount"": " & mSuccessCount & ","
    json = json & """FailureCount"": " & mFailureCount & ","
    json = json & """ErrorRate"": " & mErrorRate & ","
    json = json & """AverageRuleExecutionTime"": " & mAverageRuleExecutionTime & ","
    json = json & """MaxRuleExecutionTime"": " & mMaxRuleExecutionTime & ","
    json = json & """MinRuleExecutionTime"": " & mMinRuleExecutionTime & ","
    json = json & """CacheHitRate"": " & mCacheHitRate & ","
    json = json & """MemoryUsage"": " & mMemoryUsage
    json = json & "}"
    ToJSON = json
End Function

'@Description("メトリクスをリセットします")
Public Sub Reset()
    mTotalExecutionTime = 0
    mRuleCount = 0
    mSuccessCount = 0
    mFailureCount = 0
    mErrorRate = 0
    mAverageRuleExecutionTime = 0
    mMaxRuleExecutionTime = 0
    mMinRuleExecutionTime = CDbl(2147483647)
    mCacheHitRate = 0
    mMemoryUsage = 0
    Set mRuleMetrics = New Collection
    Set mTimeSeriesData = New Collection
    Set mResourceUtilization = New Collection
End Sub