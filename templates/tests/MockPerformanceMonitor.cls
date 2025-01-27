VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockPerformanceMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements IPerformanceMonitor

Private mEnabled As Boolean
Private mStartCallCount As Long
Private mEndCallCount As Long
Private mMetrics As Collection
Private mCurrentOperation As String

Private Sub Class_Initialize()
    mEnabled = True
    Set mMetrics = New Collection
    ResetCounters
End Sub

Private Sub Class_Terminate()
    Set mMetrics = Nothing
End Sub

'@Description("パフォーマンス監視の有効/無効を設定")
Public Property Let Enabled(ByVal value As Boolean)
    mEnabled = value
End Property

'@Description("パフォーマンス監視の有効/無効を取得")
Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

'@Description("StartMeasurementの呼び出し回数を取得")
Public Property Get StartCallCount() As Long
    StartCallCount = mStartCallCount
End Property

'@Description("EndMeasurementの呼び出し回数を取得")
Public Property Get EndCallCount() As Long
    EndCallCount = mEndCallCount
End Property

'@Description("記録されたメトリクスを取得")
Public Property Get Metrics() As Collection
    Set Metrics = mMetrics
End Property

'@Description("カウンターをリセット")
Public Sub ResetCounters()
    mStartCallCount = 0
    mEndCallCount = 0
    Set mMetrics = New Collection
    mCurrentOperation = ""
End Sub

'@Description("メトリクスを追加")
Public Sub AddMetric(ByVal operationName As String, ByVal executionTime As Double)
    Dim metric As New Dictionary
    metric.Add "Operation", operationName
    metric.Add "ExecutionTime", executionTime
    metric.Add "Timestamp", Now
    mMetrics.Add metric
End Sub

' IPerformanceMonitor実装

Private Sub IPerformanceMonitor_StartMeasurement(ByVal operationName As String)
    If Not mEnabled Then Exit Sub
    
    mStartCallCount = mStartCallCount + 1
    mCurrentOperation = operationName
End Sub

Private Sub IPerformanceMonitor_EndMeasurement()
    If Not mEnabled Then Exit Sub
    
    mEndCallCount = mEndCallCount + 1
    
    If mCurrentOperation <> "" Then
        ' 実行時間をシミュレート（0.1秒）
        AddMetric mCurrentOperation, 0.1
        mCurrentOperation = ""
    End If
End Sub

Private Function IPerformanceMonitor_GetMetrics() As Collection
    Set IPerformanceMonitor_GetMetrics = mMetrics
End Function

Private Sub IPerformanceMonitor_ClearMetrics()
    Set mMetrics = New Collection
End Sub

Private Property Get IPerformanceMonitor_Enabled() As Boolean
    IPerformanceMonitor_Enabled = mEnabled
End Property

Private Property Let IPerformanceMonitor_Enabled(ByVal value As Boolean)
    mEnabled = value
End Property