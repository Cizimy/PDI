VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorPerformanceMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("エラー処理のパフォーマンスを監視するクラス")

' メトリクス情報を格納する型
Private Type MetricInfo
    Value As Double           ' メトリック値
    Timestamp As Date        ' 記録時刻
    Category As String       ' メトリックカテゴリ
    Threshold As Double      ' 警告しきい値
    CriticalThreshold As Double ' 重大しきい値
End Type

' プライベート変数
Private mMetrics As Dictionary           ' メトリクスの保存
Private mHistory As Collection           ' 履歴データ
Private mThresholds As Dictionary        ' しきい値設定
Private mAlerts As Collection            ' アラート履歴
Private mIsEnabled As Boolean            ' 監視有効フラグ
Private mSamplingInterval As Long        ' サンプリング間隔（ミリ秒）
Private mRetentionPeriod As Long         ' データ保持期間（日）

' イベント
Public Event ThresholdExceeded(ByVal metricName As String, ByVal currentValue As Double, ByVal threshold As Double)
Public Event CriticalThresholdExceeded(ByVal metricName As String, ByVal currentValue As Double, ByVal threshold As Double)
Public Event MetricUpdated(ByVal metricName As String, ByVal newValue As Double)
Public Event PerformanceDataCollected(ByVal timestamp As Date, ByVal metrics As Dictionary)

Private Sub Class_Initialize()
    Set mMetrics = New Dictionary
    Set mHistory = New Collection
    Set mThresholds = New Dictionary
    Set mAlerts = New Collection
    
    ' デフォルト設定
    mIsEnabled = True
    mSamplingInterval = 1000 ' 1秒
    mRetentionPeriod = 30 ' 30日
    
    InitializeMetrics
End Sub

Private Sub InitializeMetrics()
    ' 基本メトリクス
    AddMetric "AverageHandlingTime", "Time", 0, 1000, 5000 ' 警告: 1秒, 重大: 5秒
    AddMetric "ErrorRate", "Rate", 0, 10, 50 ' 警告: 10%, 重大: 50%
    AddMetric "SuccessRate", "Rate", 100, 90, 80 ' 警告: 90%未満, 重大: 80%未満
    AddMetric "RetryRate", "Rate", 0, 20, 40 ' 警告: 20%, 重大: 40%
    AddMetric "ResourceUsage", "Usage", 0, 80, 95 ' 警告: 80%, 重大: 95%
    
    ' カテゴリ別メトリクス
    Dim category As ErrorCategory
    For category = ecInputError To ecUnknownCategory
        AddMetric "ErrorCount_" & category, "Count", 0, 100, 500
    Next category
End Sub

'@Description("新しいメトリクスを追加します")
Private Sub AddMetric(ByVal name As String, ByVal category As String, _
                     ByVal initialValue As Double, ByVal warningThreshold As Double, _
                     ByVal criticalThreshold As Double)
    Dim info As MetricInfo
    
    With info
        .Value = initialValue
        .Timestamp = Now
        .Category = category
        .Threshold = warningThreshold
        .CriticalThreshold = criticalThreshold
    End With
    
    mMetrics.Add name, info
End Sub

'@Description("メトリクス値を更新します")
Public Sub UpdateMetric(ByVal name As String, ByVal value As Double)
    If Not mIsEnabled Then Exit Sub
    
    If mMetrics.Exists(name) Then
        Dim info As MetricInfo
        info = mMetrics(name)
        
        ' 値の更新
        info.Value = value
        info.Timestamp = Now
        mMetrics(name) = info
        
        ' 履歴の記録
        RecordHistory name, value
        
        ' しきい値チェック
        CheckThresholds name, value
        
        ' イベント通知
        RaiseEvent MetricUpdated(name, value)
    End If
End Sub

'@Description("履歴データを記録します")
Private Sub RecordHistory(ByVal name As String, ByVal value As Double)
    Dim record As Dictionary
    Set record = New Dictionary
    
    With record
        .Add "Name", name
        .Add "Value", value
        .Add "Timestamp", Now
        .Add "Category", mMetrics(name).Category
    End With
    
    mHistory.Add record
    
    ' 古いデータの削除
    CleanupHistory
End Sub

'@Description("古い履歴データを削除します")
Private Sub CleanupHistory()
    Dim cutoffDate As Date
    cutoffDate = DateAdd("d", -mRetentionPeriod, Now)
    
    Dim i As Long
    For i = mHistory.Count To 1 Step -1
        If mHistory(i)("Timestamp") < cutoffDate Then
            mHistory.Remove i
        End If
    Next i
End Sub

'@Description("しきい値をチェックします")
Private Sub CheckThresholds(ByVal name As String, ByVal value As Double)
    Dim info As MetricInfo
    info = mMetrics(name)
    
    ' 重大しきい値のチェック
    If value >= info.CriticalThreshold Then
        RecordAlert name, value, info.CriticalThreshold, True
        RaiseEvent CriticalThresholdExceeded(name, value, info.CriticalThreshold)
    ' 警告しきい値のチェック
    ElseIf value >= info.Threshold Then
        RecordAlert name, value, info.Threshold, False
        RaiseEvent ThresholdExceeded(name, value, info.Threshold)
    End If
End Sub

'@Description("アラートを記録します")
Private Sub RecordAlert(ByVal name As String, ByVal value As Double, _
                       ByVal threshold As Double, ByVal isCritical As Boolean)
    Dim alert As Dictionary
    Set alert = New Dictionary
    
    With alert
        .Add "Name", name
        .Add "Value", value
        .Add "Threshold", threshold
        .Add "IsCritical", isCritical
        .Add "Timestamp", Now
    End With
    
    mAlerts.Add alert
End Sub

'@Description("エラー処理時間を記録します")
Public Sub TrackHandlingTime(ByVal errorCode As ErrorCode, ByVal duration As Long)
    UpdateMetric "AverageHandlingTime", CalculateAverageHandlingTime(duration)
    UpdateMetric "ErrorCount_" & GetErrorCategory(errorCode), GetErrorCount(errorCode) + 1
End Sub

'@Description("平均処理時間を計算します")
Private Function CalculateAverageHandlingTime(ByVal newDuration As Long) As Double
    Dim totalTime As Double
    Dim count As Long
    
    ' 最新の100件の履歴から計算
    Dim i As Long
    Dim startIndex As Long
    startIndex = WorksheetFunction.Max(1, mHistory.Count - 99)
    
    For i = startIndex To mHistory.Count
        If mHistory(i)("Name") = "AverageHandlingTime" Then
            totalTime = totalTime + CDbl(mHistory(i)("Value"))
            count = count + 1
        End If
    Next i
    
    If count = 0 Then
        CalculateAverageHandlingTime = newDuration
    Else
        CalculateAverageHandlingTime = (totalTime + newDuration) / (count + 1)
    End If
End Function

'@Description("エラーカテゴリを取得します")
Private Function GetErrorCategory(ByVal errorCode As ErrorCode) As ErrorCategory
    Select Case errorCode
        Case ecInvalidInput: GetErrorCategory = ecInputError
        Case ecDatabaseError: GetErrorCategory = ecDatabaseError
        Case ecNetworkError: GetErrorCategory = ecNetworkError
        Case ecFileNotFound, ecFileAccessDenied: GetErrorCategory = ecFileIOError
        Case Else: GetErrorCategory = ecUnknownCategory
    End Select
End Function

'@Description("エラー数を取得します")
Private Function GetErrorCount(ByVal errorCode As ErrorCode) As Long
    Dim categoryMetric As String
    categoryMetric = "ErrorCount_" & GetErrorCategory(errorCode)
    
    If mMetrics.Exists(categoryMetric) Then
        GetErrorCount = mMetrics(categoryMetric).Value
    Else
        GetErrorCount = 0
    End If
End Function

'@Description("パーセンタイル値を計算します")
Public Function GetHandlingTimePercentiles() As Dictionary
    Dim result As New Dictionary
    Dim times As Collection
    Set times = New Collection
    
    ' 処理時間データの収集
    Dim record As Variant
    For Each record In mHistory
        If record("Name") = "AverageHandlingTime" Then
            times.Add record("Value")
        End If
    Next record
    
    ' データが十分にある場合のみパーセンタイルを計算
    If times.Count >= 10 Then
        result.Add "p50", CalculatePercentile(times, 0.5)
        result.Add "p90", CalculatePercentile(times, 0.9)
        result.Add "p95", CalculatePercentile(times, 0.95)
        result.Add "p99", CalculatePercentile(times, 0.99)
    End If
    
    Set GetHandlingTimePercentiles = result
End Function

'@Description("パーセンタイル値を計算します")
Private Function CalculatePercentile(ByVal values As Collection, ByVal percentile As Double) As Double
    ' 値を配列にコピー
    Dim arr() As Double
    ReDim arr(1 To values.Count)
    
    Dim i As Long
    For i = 1 To values.Count
        arr(i) = values(i)
    Next i
    
    ' 配列をソート
    QuickSort arr, 1, values.Count
    
    ' パーセンタイルの位置を計算
    Dim position As Double
    position = 1 + (values.Count - 1) * percentile
    
    ' 線形補間
    Dim lower As Long, upper As Long
    lower = Int(position)
    upper = lower + 1
    
    If upper > values.Count Then
        CalculatePercentile = arr(lower)
    Else
        Dim fraction As Double
        fraction = position - lower
        CalculatePercentile = arr(lower) + fraction * (arr(upper) - arr(lower))
    End If
End Function

'@Description("クイックソートを実行します")
Private Sub QuickSort(ByRef arr() As Double, ByVal left As Long, ByVal right As Long)
    Dim i As Long, j As Long
    Dim pivot As Double
    Dim temp As Double
    
    i = left
    j = right
    pivot = arr((left + right) \ 2)
    
    Do While i <= j
        Do While arr(i) < pivot And i < right
            i = i + 1
        Loop
        
        Do While arr(j) > pivot And j > left
            j = j - 1
        Loop
        
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    
    If left < j Then QuickSort arr, left, j
    If i < right Then QuickSort arr, i, right
End Sub

'@Description("監視を有効化します")
Public Sub Enable()
    mIsEnabled = True
End Sub

'@Description("監視を無効化します")
Public Sub Disable()
    mIsEnabled = False
End Sub

'@Description("サンプリング間隔を設定します")
Public Property Let SamplingInterval(ByVal value As Long)
    If value >= 100 Then ' 最小100ミリ秒
        mSamplingInterval = value
    End If
End Property

'@Description("データ保持期間を設定します")
Public Property Let RetentionPeriod(ByVal value As Long)
    If value > 0 Then
        mRetentionPeriod = value
        CleanupHistory ' 即座にクリーンアップを実行
    End If
End Property

'@Description("メトリクスのサマリーを取得します")
Public Function GetMetricsSummary() As Dictionary
    Dim summary As New Dictionary
    Dim metric As Variant
    
    For Each metric In mMetrics.Keys
        Dim info As MetricInfo
        info = mMetrics(metric)
        
        Dim metricSummary As New Dictionary
        With metricSummary
            .Add "Value", info.Value
            .Add "Category", info.Category
            .Add "LastUpdated", info.Timestamp
            .Add "WarningThreshold", info.Threshold
            .Add "CriticalThreshold", info.CriticalThreshold
        End With
        
        summary.Add CStr(metric), metricSummary
    Next metric
    
    Set GetMetricsSummary = summary
End Function

'@Description("アラート履歴を取得します")
Public Function GetAlertHistory() As Collection
    Set GetAlertHistory = mAlerts
End Function

'@Description("メトリクス履歴を取得します")
Public Function GetMetricHistory(ByVal metricName As String, _
                               Optional ByVal startDate As Date, _
                               Optional ByVal endDate As Date) As Collection
    Dim result As New Collection
    
    If startDate = #12:00:00 AM# Then startDate = DateAdd("d", -mRetentionPeriod, Now)
    If endDate = #12:00:00 AM# Then endDate = Now
    
    Dim record As Variant
    For Each record In mHistory
        If record("Name") = metricName And _
           record("Timestamp") >= startDate And _
           record("Timestamp") <= endDate Then
            result.Add record
        End If
    Next record
    
    Set GetMetricHistory = result
End Function