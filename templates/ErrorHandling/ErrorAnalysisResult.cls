VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorAnalysisResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("エラー分析結果を格納するクラス")

'@Description("分析されたエラーの根本原因")
Public RootCause As String

'@Description("エラーの影響範囲")
Public ImpactArea As String

'@Description("エラーの推奨される対策")
Public RemediationSteps As String

'@Description("関連するエラー情報")
Public RelatedErrors As Collection

'@Description("分析の信頼度")
Public ConfidenceLevel As Double ' 0.0 - 1.0

'@Description("分析の実行時間")
Public AnalysisDuration As Double ' in seconds

'@Description("分析に使用されたデータソース")
Public DataSources As Collection

'@Description("エラーパターン")
Public ErrorPatterns As Collection

'@Description("重大度の分布")
Public SeverityDistribution As Object ' Dictionary

'@Description("カテゴリの分布")
Public CategoryDistribution As Object ' Dictionary

'@Description("時間帯別の発生頻度")
Public TimeDistribution As Object ' Dictionary

'@Description("分析の開始時刻")
Public AnalysisStartTime As Date

'@Description("分析の終了時刻")
Public AnalysisEndTime As Date

'@Description("分析に使用したルール")
Public AppliedRules As Collection

'@Description("検出された異常")
Public DetectedAnomalies As Collection

Private Sub Class_Initialize()
    Set RelatedErrors = New Collection
    Set DataSources = New Collection
    Set ErrorPatterns = New Collection
    Set SeverityDistribution = CreateObject("Scripting.Dictionary")
    Set CategoryDistribution = CreateObject("Scripting.Dictionary")
    Set TimeDistribution = CreateObject("Scripting.Dictionary")
    Set AppliedRules = New Collection
    Set DetectedAnomalies = New Collection
    ConfidenceLevel = 0
    AnalysisDuration = 0
    AnalysisStartTime = Now
End Sub

Private Sub Class_Terminate()
    Set RelatedErrors = Nothing
    Set DataSources = Nothing
    Set ErrorPatterns = Nothing
    Set SeverityDistribution = Nothing
    Set CategoryDistribution = Nothing
    Set TimeDistribution = Nothing
    Set AppliedRules = Nothing
    Set DetectedAnomalies = Nothing
End Sub

'@Description("分析結果を文字列形式で取得します")
Public Function ToString() As String
    Dim result As String
    
    result = "Error Analysis Result:" & vbCrLf & _
             "--------------------" & vbCrLf & _
             "Root Cause: " & RootCause & vbCrLf & _
             "Impact Area: " & ImpactArea & vbCrLf & _
             "Remediation Steps: " & RemediationSteps & vbCrLf & _
             "Confidence Level: " & Format(ConfidenceLevel, "0.00") & vbCrLf & _
             "Analysis Duration: " & Format(AnalysisDuration, "0.000") & " seconds" & vbCrLf & _
             "Related Errors: " & RelatedErrors.Count & vbCrLf & _
             "Data Sources: " & DataSources.Count & vbCrLf & _
             "Error Patterns: " & ErrorPatterns.Count & vbCrLf & _
             "Detected Anomalies: " & DetectedAnomalies.Count & vbCrLf & _
             "Analysis Period: " & Format(AnalysisStartTime, "yyyy/mm/dd hh:nn:ss") & " - " & _
                                Format(AnalysisEndTime, "yyyy/mm/dd hh:nn:ss")
    
    ToString = result
End Function

'@Description("分析結果を検証します")
Public Function Validate() As Boolean
    ' 必須フィールドの検証
    If Len(RootCause) = 0 Then Exit Function
    If Len(ImpactArea) = 0 Then Exit Function
    If Len(RemediationSteps) = 0 Then Exit Function
    If RelatedErrors.Count = 0 Then Exit Function
    If DataSources.Count = 0 Then Exit Function
    
    ' 値の範囲チェック
    If ConfidenceLevel < 0 Or ConfidenceLevel > 1 Then Exit Function
    If AnalysisDuration < 0 Then Exit Function
    
    ' 時刻の整合性チェック
    If AnalysisEndTime < AnalysisStartTime Then Exit Function
    
    Validate = True
End Function

'@Description("分析結果をマージします")
Public Sub MergeWith(ByVal other As ErrorAnalysisResult)
    If other Is Nothing Then Exit Sub
    
    ' 空の場合のみ更新
    If Len(Me.RootCause) = 0 Then Me.RootCause = other.RootCause
    If Len(Me.ImpactArea) = 0 Then Me.ImpactArea = other.ImpactArea
    If Len(Me.RemediationSteps) = 0 Then Me.RemediationSteps = other.RemediationSteps
    
    ' コレクションのマージ
    MergeCollection Me.RelatedErrors, other.RelatedErrors
    MergeCollection Me.DataSources, other.DataSources
    MergeCollection Me.ErrorPatterns, other.ErrorPatterns
    MergeCollection Me.AppliedRules, other.AppliedRules
    MergeCollection Me.DetectedAnomalies, other.DetectedAnomalies
    
    ' ディクショナリのマージ
    MergeDictionary Me.SeverityDistribution, other.SeverityDistribution
    MergeDictionary Me.CategoryDistribution, other.CategoryDistribution
    MergeDictionary Me.TimeDistribution, other.TimeDistribution
    
    ' 数値の更新
    Me.ConfidenceLevel = (Me.ConfidenceLevel + other.ConfidenceLevel) / 2
    Me.AnalysisDuration = Me.AnalysisDuration + other.AnalysisDuration
    
    ' 時刻の更新
    If other.AnalysisStartTime < Me.AnalysisStartTime Then
        Me.AnalysisStartTime = other.AnalysisStartTime
    End If
    If other.AnalysisEndTime > Me.AnalysisEndTime Then
        Me.AnalysisEndTime = other.AnalysisEndTime
    End If
End Sub

'@Description("コレクションをマージします")
Private Sub MergeCollection(ByVal target As Collection, ByVal source As Collection)
    Dim item As Variant
    For Each item In source
        If Not Contains(target, item) Then
            target.Add item
        End If
    Next
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

'@Description("コレクションに要素が含まれているかチェックします")
Private Function Contains(ByVal col As Collection, ByVal item As Variant) As Boolean
    Dim existing As Variant
    For Each existing In col
        If existing = item Then
            Contains = True
            Exit Function
        End If
    Next
    Contains = False
End Function