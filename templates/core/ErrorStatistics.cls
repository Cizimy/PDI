VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("エラー統計情報を格納するクラス")

'@Description("エラーコード別の発生回数")
Public ErrorCountByCode As Object ' Dictionary

'@Description("重大度別のエラー発生回数")
Public ErrorCountBySeverity As Object ' Dictionary

'@Description("カテゴリ別のエラー発生回数")
Public ErrorCountByCategory As Object ' Dictionary

'@Description("時間帯別のエラー発生回数")
Public ErrorCountByTimeSlot As Object ' Dictionary

'@Description("エラーの総数")
Public TotalErrorCount As Long

'@Description("最初のエラー発生時刻")
Public FirstErrorTimestamp As Date

'@Description("最後のエラー発生時刻")
Public LastErrorTimestamp As Date

'@Description("平均エラー処理時間")
Public AverageHandlingTime As Double ' in seconds

'@Description("エラー発生率")
Public ErrorRate As Double ' errors per hour

'@Description("エラー解決率")
Public ResolutionRate As Double ' resolved errors / total errors

'@Description("平均解決時間")
Public AverageResolutionTime As Double ' in seconds

'@Description("リトライ成功率")
Public RetrySuccessRate As Double ' successful retries / total retries

'@Description("エラー傾向情報")
Public TrendData As Collection ' Collection of TrendInfo

Private Sub Class_Initialize()
    Set ErrorCountByCode = CreateObject("Scripting.Dictionary")
    Set ErrorCountBySeverity = CreateObject("Scripting.Dictionary")
    Set ErrorCountByCategory = CreateObject("Scripting.Dictionary")
    Set ErrorCountByTimeSlot = CreateObject("Scripting.Dictionary")
    Set TrendData = New Collection
    
    TotalErrorCount = 0
    FirstErrorTimestamp = Now
    LastErrorTimestamp = Now
    AverageHandlingTime = 0
    ErrorRate = 0
    ResolutionRate = 0
    AverageResolutionTime = 0
    RetrySuccessRate = 0
End Sub

Private Sub Class_Terminate()
    Set ErrorCountByCode = Nothing
    Set ErrorCountBySeverity = Nothing
    Set ErrorCountByCategory = Nothing
    Set ErrorCountByTimeSlot = Nothing
    Set TrendData = Nothing
End Sub

'@Description("エラーコード別のエラー発生回数を更新します")
Public Sub IncrementErrorCount(ByVal errorCode As ErrorCode)
    If ErrorCountByCode.Exists(errorCode) Then
        ErrorCountByCode(errorCode) = ErrorCountByCode(errorCode) + 1
    Else
        ErrorCountByCode.Add errorCode, 1
    End If
    TotalErrorCount = TotalErrorCount + 1
    LastErrorTimestamp = Now
End Sub

'@Description("重大度別のエラー発生回数を更新します")
Public Sub IncrementSeverityCount(ByVal severity As ErrorSeverity)
    If ErrorCountBySeverity.Exists(severity) Then
        ErrorCountBySeverity(severity) = ErrorCountBySeverity(severity) + 1
    Else
        ErrorCountBySeverity.Add severity, 1
    End If
End Sub

'@Description("カテゴリ別のエラー発生回数を更新します")
Public Sub IncrementCategoryCount(ByVal category As ErrorCategory)
    If ErrorCountByCategory.Exists(category) Then
        ErrorCountByCategory(category) = ErrorCountByCategory(category) + 1
    Else
        ErrorCountByCategory.Add category, 1
    End If
End Sub

'@Description("時間帯別のエラー発生回数を更新します")
Public Sub IncrementTimeSlotCount(ByVal timeSlot As String)
    If ErrorCountByTimeSlot.Exists(timeSlot) Then
        ErrorCountByTimeSlot(timeSlot) = ErrorCountByTimeSlot(timeSlot) + 1
    Else
        ErrorCountByTimeSlot.Add timeSlot, 1
    End If
End Sub

'@Description("エラー統計情報を文字列形式で取得します")
Public Function ToString() As String
    Dim result As String
    
    result = "Error Statistics:" & vbCrLf & _
             "----------------" & vbCrLf & _
             "Total Errors: " & TotalErrorCount & vbCrLf & _
             "First Error: " & Format(FirstErrorTimestamp, "yyyy/mm/dd hh:nn:ss") & vbCrLf & _
             "Last Error: " & Format(LastErrorTimestamp, "yyyy/mm/dd hh:nn:ss") & vbCrLf & _
             "Average Handling Time: " & Format(AverageHandlingTime, "0.000") & " seconds" & vbCrLf & _
             "Error Rate: " & Format(ErrorRate, "0.00") & " errors/hour" & vbCrLf & _
             "Resolution Rate: " & Format(ResolutionRate * 100, "0.00") & "%" & vbCrLf & _
             "Average Resolution Time: " & Format(AverageResolutionTime, "0.000") & " seconds" & vbCrLf & _
             "Retry Success Rate: " & Format(RetrySuccessRate * 100, "0.00") & "%"
    
    ToString = result
End Function

'@Description("エラー統計情報を更新します")
Public Sub UpdateStatistics(ByVal errorInfo As ErrorInfo)
    If errorInfo Is Nothing Then Exit Sub
    
    ' カウンターの更新
    IncrementErrorCount errorInfo.ErrorCode
    IncrementSeverityCount errorInfo.Severity
    IncrementCategoryCount errorInfo.Category
    IncrementTimeSlotCount Format(errorInfo.Timestamp, "hh:00")
    
    ' 時刻の更新
    If errorInfo.Timestamp < FirstErrorTimestamp Then
        FirstErrorTimestamp = errorInfo.Timestamp
    End If
    LastErrorTimestamp = errorInfo.Timestamp
    
    ' レート計算の更新
    UpdateRates
End Sub

'@Description("レート計算を更新します")
Private Sub UpdateRates()
    ' エラー発生率の計算
    Dim timeSpan As Double
    timeSpan = DateDiff("h", FirstErrorTimestamp, LastErrorTimestamp)
    If timeSpan > 0 Then
        ErrorRate = TotalErrorCount / timeSpan
    End If
End Sub

'@Description("統計情報をリセットします")
Public Sub Reset()
    Set ErrorCountByCode = CreateObject("Scripting.Dictionary")
    Set ErrorCountBySeverity = CreateObject("Scripting.Dictionary")
    Set ErrorCountByCategory = CreateObject("Scripting.Dictionary")
    Set ErrorCountByTimeSlot = CreateObject("Scripting.Dictionary")
    Set TrendData = New Collection
    
    TotalErrorCount = 0
    FirstErrorTimestamp = Now
    LastErrorTimestamp = Now
    AverageHandlingTime = 0
    ErrorRate = 0
    ResolutionRate = 0
    AverageResolutionTime = 0
    RetrySuccessRate = 0
End Sub

'@Description("統計情報をマージします")
Public Sub MergeWith(ByVal other As ErrorStatistics)
    If other Is Nothing Then Exit Sub
    
    ' ディクショナリのマージ
    MergeDictionary ErrorCountByCode, other.ErrorCountByCode
    MergeDictionary ErrorCountBySeverity, other.ErrorCountBySeverity
    MergeDictionary ErrorCountByCategory, other.ErrorCountByCategory
    MergeDictionary ErrorCountByTimeSlot, other.ErrorCountByTimeSlot
    
    ' 数値の更新
    TotalErrorCount = TotalErrorCount + other.TotalErrorCount
    
    ' 時刻の更新
    If other.FirstErrorTimestamp < FirstErrorTimestamp Then
        FirstErrorTimestamp = other.FirstErrorTimestamp
    End If
    If other.LastErrorTimestamp > LastErrorTimestamp Then
        LastErrorTimestamp = other.LastErrorTimestamp
    End If
    
    ' 平均値の更新
    AverageHandlingTime = (AverageHandlingTime + other.AverageHandlingTime) / 2
    AverageResolutionTime = (AverageResolutionTime + other.AverageResolutionTime) / 2
    
    ' レートの更新
    UpdateRates
End Sub

'@Description("ディクショナリをマージします")
Private Sub MergeDictionary(ByVal target As Object, ByVal source As Object)
    Dim key As Variant
    For Each key In source.Keys
        If target.Exists(key) Then
            target(key) = target(key) + source(key)
        Else
            target.Add key, source(key)
        End If
    Next
End Sub