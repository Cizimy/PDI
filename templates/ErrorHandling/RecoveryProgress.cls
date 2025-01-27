VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecoveryProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("回復処理の進捗状況を管理するクラス")

' プライベート変数
Private mTotalSteps As Long
Private mCompletedSteps As Long
Private mCurrentStep As Long
Private mStartTime As Date
Private mEstimatedEndTime As Date
Private mCurrentPhase As String
Private mPhaseProgress As Dictionary
Private mMetrics As Dictionary
Private mStatus As RecoveryStatus

' 進捗状況を定義する列挙型
Public Enum RecoveryStatus
    rsNotStarted = 0    ' 未開始
    rsInProgress = 1    ' 実行中
    rsPaused = 2        ' 一時停止
    rsCompleted = 3     ' 完了
    rsFailed = 4        ' 失敗
    rsCancelled = 5     ' キャンセル
End Enum

Private Sub Class_Initialize()
    mStartTime = Now
    Set mPhaseProgress = New Dictionary
    Set mMetrics = New Dictionary
    mStatus = rsNotStarted
End Sub

'@Description("総ステップ数を設定")
Public Property Let TotalSteps(ByVal value As Long)
    mTotalSteps = value
End Property

'@Description("総ステップ数を取得")
Public Property Get TotalSteps() As Long
    TotalSteps = mTotalSteps
End Property

'@Description("完了ステップ数を設定")
Public Property Let CompletedSteps(ByVal value As Long)
    mCompletedSteps = value
End Property

'@Description("完了ステップ数を取得")
Public Property Get CompletedSteps() As Long
    CompletedSteps = mCompletedSteps
End Property

'@Description("現在のステップを設定")
Public Property Let CurrentStep(ByVal value As Long)
    mCurrentStep = value
End Property

'@Description("現在のステップを取得")
Public Property Get CurrentStep() As Long
    CurrentStep = mCurrentStep
End Property

'@Description("開始時刻を取得")
Public Property Get StartTime() As Date
    StartTime = mStartTime
End Property

'@Description("推定終了時刻を設定")
Public Property Let EstimatedEndTime(ByVal value As Date)
    mEstimatedEndTime = value
End Property

'@Description("推定終了時刻を取得")
Public Property Get EstimatedEndTime() As Date
    EstimatedEndTime = mEstimatedEndTime
End Property

'@Description("現在のフェーズを設定")
Public Property Let CurrentPhase(ByVal value As String)
    mCurrentPhase = value
End Property

'@Description("現在のフェーズを取得")
Public Property Get CurrentPhase() As String
    CurrentPhase = mCurrentPhase
End Property

'@Description("進捗状況を設定")
Public Property Let Status(ByVal value As RecoveryStatus)
    mStatus = value
End Property

'@Description("進捗状況を取得")
Public Property Get Status() As RecoveryStatus
    Status = mStatus
End Property

'@Description("進捗率を計算（パーセント）")
Public Function GetProgressPercentage() As Double
    If mTotalSteps = 0 Then
        GetProgressPercentage = 0
        Exit Function
    End If
    
    GetProgressPercentage = (mCompletedSteps / mTotalSteps) * 100
End Function

'@Description("経過時間を取得（秒）")
Public Function GetElapsedTime() As Long
    GetElapsedTime = DateDiff("s", mStartTime, Now)
End Function

'@Description("推定残り時間を取得（秒）")
Public Function GetEstimatedRemainingTime() As Long
    If mCompletedSteps = 0 Then
        GetEstimatedRemainingTime = 0
        Exit Function
    End If
    
    Dim elapsedTime As Long
    elapsedTime = GetElapsedTime()
    
    Dim timePerStep As Double
    timePerStep = elapsedTime / mCompletedSteps
    
    GetEstimatedRemainingTime = CLng(timePerStep * (mTotalSteps - mCompletedSteps))
End Function

'@Description("フェーズごとの進捗を記録")
Public Sub UpdatePhaseProgress(ByVal phase As String, ByVal progress As Double)
    If mPhaseProgress.Exists(phase) Then
        mPhaseProgress(phase) = progress
    Else
        mPhaseProgress.Add phase, progress
    End If
End Sub

'@Description("フェーズごとの進捗を取得")
Public Function GetPhaseProgress(ByVal phase As String) As Double
    If mPhaseProgress.Exists(phase) Then
        GetPhaseProgress = mPhaseProgress(phase)
    Else
        GetPhaseProgress = 0
    End If
End Function

'@Description("メトリクスを追加")
Public Sub AddMetric(ByVal name As String, ByVal value As Variant)
    If mMetrics.Exists(name) Then
        mMetrics(name) = value
    Else
        mMetrics.Add name, value
    End If
End Sub

'@Description("メトリクスを取得")
Public Function GetMetric(ByVal name As String) As Variant
    If mMetrics.Exists(name) Then
        GetMetric = mMetrics(name)
    Else
        GetMetric = Empty
    End If
End Function

'@Description("進捗情報を文字列形式で取得")
Public Function ToString() As String
    Dim result As String
    
    result = "Recovery Progress:" & vbCrLf & _
             "----------------" & vbCrLf & _
             "Status: " & GetStatusDescription(mStatus) & vbCrLf & _
             "Progress: " & Format(GetProgressPercentage(), "0.00") & "%" & vbCrLf & _
             "Steps: " & mCompletedSteps & "/" & mTotalSteps & vbCrLf & _
             "Current Phase: " & mCurrentPhase & vbCrLf & _
             "Elapsed Time: " & GetElapsedTime() & " seconds" & vbCrLf & _
             "Estimated Remaining: " & GetEstimatedRemainingTime() & " seconds"
    
    ToString = result
End Function

'@Description("進捗状況の説明を取得")
Private Function GetStatusDescription(ByVal status As RecoveryStatus) As String
    Select Case status
        Case rsNotStarted: GetStatusDescription = "未開始"
        Case rsInProgress: GetStatusDescription = "実行中"
        Case rsPaused: GetStatusDescription = "一時停止"
        Case rsCompleted: GetStatusDescription = "完了"
        Case rsFailed: GetStatusDescription = "失敗"
        Case rsCancelled: GetStatusDescription = "キャンセル"
        Case Else: GetStatusDescription = "不明"
    End Select
End Function