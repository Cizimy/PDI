VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationQueueStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("バリデーションキューの状態情報を管理するクラス")

' メンバー変数
Private mQueueLength As Long            ' キューの長さ
Private mPendingJobs As Long           ' 待機中のジョブ数
Private mAverageWaitingTime As Double   ' 平均待機時間（ミリ秒）
Private mMaxQueueLength As Long         ' 最大キュー長
Private mCurrentThroughput As Double    ' 現在のスループット（処理/秒）
Private mActiveJobs As Long            ' アクティブなジョブ数
Private mCompletedJobs As Long         ' 完了したジョブ数
Private mFailedJobs As Long            ' 失敗したジョブ数
Private mLastUpdateTime As Date        ' 最終更新時刻
Private mMetrics As Collection         ' パフォーマンスメトリクス

' イベント
Public Event StatusUpdated(ByVal propertyName As String, ByVal oldValue As Variant, ByVal newValue As Variant)
Public Event ThresholdExceeded(ByVal metricName As String, ByVal currentValue As Double, ByVal threshold As Double)

Private Sub Class_Initialize()
    mQueueLength = 0
    mPendingJobs = 0
    mAverageWaitingTime = 0
    mMaxQueueLength = 0
    mCurrentThroughput = 0
    mActiveJobs = 0
    mCompletedJobs = 0
    mFailedJobs = 0
    mLastUpdateTime = Now
    Set mMetrics = New Collection
End Sub

'@Description("キューの長さを取得または設定します")
Public Property Get QueueLength() As Long
    QueueLength = mQueueLength
End Property

Public Property Let QueueLength(ByVal value As Long)
    Dim oldValue As Long
    oldValue = mQueueLength
    mQueueLength = value
    RaiseEvent StatusUpdated("QueueLength", oldValue, value)
    CheckThresholds
End Property

'@Description("待機中のジョブ数を取得または設定します")
Public Property Get PendingJobs() As Long
    PendingJobs = mPendingJobs
End Property

Public Property Let PendingJobs(ByVal value As Long)
    Dim oldValue As Long
    oldValue = mPendingJobs
    mPendingJobs = value
    RaiseEvent StatusUpdated("PendingJobs", oldValue, value)
    CheckThresholds
End Property

'@Description("平均待機時間を取得または設定します")
Public Property Get AverageWaitingTime() As Double
    AverageWaitingTime = mAverageWaitingTime
End Property

Public Property Let AverageWaitingTime(ByVal value As Double)
    Dim oldValue As Double
    oldValue = mAverageWaitingTime
    mAverageWaitingTime = value
    RaiseEvent StatusUpdated("AverageWaitingTime", oldValue, value)
    CheckThresholds
End Property

'@Description("最大キュー長を取得または設定します")
Public Property Get MaxQueueLength() As Long
    MaxQueueLength = mMaxQueueLength
End Property

Public Property Let MaxQueueLength(ByVal value As Long)
    Dim oldValue As Long
    oldValue = mMaxQueueLength
    mMaxQueueLength = value
    RaiseEvent StatusUpdated("MaxQueueLength", oldValue, value)
End Property

'@Description("現在のスループットを取得または設定します")
Public Property Get CurrentThroughput() As Double
    CurrentThroughput = mCurrentThroughput
End Property

Public Property Let CurrentThroughput(ByVal value As Double)
    Dim oldValue As Double
    oldValue = mCurrentThroughput
    mCurrentThroughput = value
    RaiseEvent StatusUpdated("CurrentThroughput", oldValue, value)
    CheckThresholds
End Property

'@Description("アクティブなジョブ数を取得または設定します")
Public Property Get ActiveJobs() As Long
    ActiveJobs = mActiveJobs
End Property

Public Property Let ActiveJobs(ByVal value As Long)
    Dim oldValue As Long
    oldValue = mActiveJobs
    mActiveJobs = value
    RaiseEvent StatusUpdated("ActiveJobs", oldValue, value)
    CheckThresholds
End Property

'@Description("完了したジョブ数を取得または設定します")
Public Property Get CompletedJobs() As Long
    CompletedJobs = mCompletedJobs
End Property

Public Property Let CompletedJobs(ByVal value As Long)
    Dim oldValue As Long
    oldValue = mCompletedJobs
    mCompletedJobs = value
    RaiseEvent StatusUpdated("CompletedJobs", oldValue, value)
End Property

'@Description("失敗したジョブ数を取得または設定します")
Public Property Get FailedJobs() As Long
    FailedJobs = mFailedJobs
End Property

Public Property Let FailedJobs(ByVal value As Long)
    Dim oldValue As Long
    oldValue = mFailedJobs
    mFailedJobs = value
    RaiseEvent StatusUpdated("FailedJobs", oldValue, value)
    CheckThresholds
End Property

'@Description("最終更新時刻を取得または設定します")
Public Property Get LastUpdateTime() As Date
    LastUpdateTime = mLastUpdateTime
End Property

Public Property Let LastUpdateTime(ByVal value As Date)
    Dim oldValue As Date
    oldValue = mLastUpdateTime
    mLastUpdateTime = value
    RaiseEvent StatusUpdated("LastUpdateTime", oldValue, value)
End Property

'@Description("パフォーマンスメトリクスを取得します")
Public Property Get Metrics() As Collection
    Set Metrics = mMetrics
End Property

'@Description("メトリクスを追加します")
'@Param("metricName: メトリクス名")
'@Param("value: メトリクス値")
Public Sub AddMetric(ByVal metricName As String, ByVal value As Double)
    Dim metric As New Dictionary
    metric.Add "Name", metricName
    metric.Add "Value", value
    metric.Add "Timestamp", Now
    mMetrics.Add metric
    CheckThresholds
End Sub

'@Description("閾値をチェックします")
Private Sub CheckThresholds()
    ' キュー長の閾値チェック
    If mQueueLength > mMaxQueueLength * 0.8 Then
        RaiseEvent ThresholdExceeded("QueueLength", mQueueLength, mMaxQueueLength * 0.8)
    End If
    
    ' 待機時間の閾値チェック
    If mAverageWaitingTime > 5000 Then ' 5秒を閾値とする
        RaiseEvent ThresholdExceeded("AverageWaitingTime", mAverageWaitingTime, 5000)
    End If
    
    ' スループットの閾値チェック
    If mCurrentThroughput < 1 Then ' 1処理/秒未満を警告とする
        RaiseEvent ThresholdExceeded("CurrentThroughput", mCurrentThroughput, 1)
    End If
End Sub

'@Description("状態をJSONフォーマットで取得します")
Public Function ToJSON() As String
    Dim json As String
    json = "{"
    json = json & """QueueLength"": " & mQueueLength & ","
    json = json & """PendingJobs"": " & mPendingJobs & ","
    json = json & """AverageWaitingTime"": " & mAverageWaitingTime & ","
    json = json & """MaxQueueLength"": " & mMaxQueueLength & ","
    json = json & """CurrentThroughput"": " & mCurrentThroughput & ","
    json = json & """ActiveJobs"": " & mActiveJobs & ","
    json = json & """CompletedJobs"": " & mCompletedJobs & ","
    json = json & """FailedJobs"": " & mFailedJobs & ","
    json = json & """LastUpdateTime"": """ & Format(mLastUpdateTime, "yyyy-mm-dd hh:nn:ss") & """"
    json = json & "}"
    ToJSON = json
End Function

'@Description("状態をリセットします")
Public Sub Reset()
    mQueueLength = 0
    mPendingJobs = 0
    mAverageWaitingTime = 0
    mCurrentThroughput = 0
    mActiveJobs = 0
    mCompletedJobs = 0
    mFailedJobs = 0
    mLastUpdateTime = Now
    Set mMetrics = New Collection
End Sub