VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("エラー情報を格納するクラス")

'@Description("エラーコード")
Public ErrorCode As ErrorCode

'@Description("エラーメッセージ")
Public Message As String

'@Description("エラー発生時刻")
Public Timestamp As Date

'@Description("エラー発生モジュール")
Public SourceModule As String

'@Description("エラー発生プロシージャ")
Public SourceProcedure As String

'@Description("エラー発生時のライン番号")
Public SourceLine As Long

'@Description("追加のエラー情報")
Public AdditionalInfo As Variant

'@Description("エラーの根本原因")
Public RootCause As String

'@Description("エラーの影響範囲")
Public ImpactArea As String

'@Description("エラーの推奨される対策")
Public RemediationSteps As String

'@Description("エラー発生時のスタックトレース")
Public StackTrace As String

'@Description("エラーの重大度")
Public Severity As ErrorSeverity

'@Description("エラーのカテゴリ")
Public Category As ErrorCategory

'@Description("エラーの状態")
Public Status As ErrorStatus

'@Description("エラーの処理試行回数")
Public RetryCount As Long

'@Description("最後の処理試行時刻")
Public LastRetryTimestamp As Date

'@Description("エラーに関連するリソース")
Public RelatedResources As Collection

'@Description("エラーのタグ")
Public Tags As Collection

Private Sub Class_Initialize()
    Set RelatedResources = New Collection
    Set Tags = New Collection
    Timestamp = Now
    Status = esNew
    RetryCount = 0
End Sub

Private Sub Class_Terminate()
    Set RelatedResources = Nothing
    Set Tags = Nothing
End Sub

'@Description("エラー情報を文字列形式で取得します")
Public Function ToString() As String
    Dim result As String
    
    result = "Error Information:" & vbCrLf & _
             "----------------" & vbCrLf & _
             "Code: " & ErrorCode & vbCrLf & _
             "Message: " & Message & vbCrLf & _
             "Time: " & Format(Timestamp, "yyyy/mm/dd hh:nn:ss") & vbCrLf & _
             "Module: " & SourceModule & vbCrLf & _
             "Procedure: " & SourceProcedure & vbCrLf & _
             "Line: " & SourceLine & vbCrLf & _
             "Severity: " & Severity & vbCrLf & _
             "Category: " & Category & vbCrLf & _
             "Status: " & Status & vbCrLf & _
             "Root Cause: " & RootCause & vbCrLf & _
             "Impact Area: " & ImpactArea & vbCrLf & _
             "Retry Count: " & RetryCount
    
    ToString = result
End Function

'@Description("エラー情報をクローンします")
Public Function Clone() As ErrorInfo
    Dim newInfo As ErrorInfo
    Set newInfo = New ErrorInfo
    
    With newInfo
        .ErrorCode = Me.ErrorCode
        .Message = Me.Message
        .Timestamp = Me.Timestamp
        .SourceModule = Me.SourceModule
        .SourceProcedure = Me.SourceProcedure
        .SourceLine = Me.SourceLine
        .AdditionalInfo = Me.AdditionalInfo
        .RootCause = Me.RootCause
        .ImpactArea = Me.ImpactArea
        .RemediationSteps = Me.RemediationSteps
        .StackTrace = Me.StackTrace
        .Severity = Me.Severity
        .Category = Me.Category
        .Status = Me.Status
        .RetryCount = Me.RetryCount
        .LastRetryTimestamp = Me.LastRetryTimestamp
        
        ' コレクションのコピー
        Dim item As Variant
        For Each item In Me.RelatedResources
            .RelatedResources.Add item
        Next
        
        For Each item In Me.Tags
            .Tags.Add item
        Next
    End With
    
    Set Clone = newInfo
End Function

'@Description("エラー情報を検証します")
Public Function Validate() As Boolean
    ' 必須フィールドの検証
    If ErrorCode = 0 Then Exit Function
    If Len(Message) = 0 Then Exit Function
    If Len(SourceModule) = 0 Then Exit Function
    If Len(SourceProcedure) = 0 Then Exit Function
    
    Validate = True
End Function

'@Description("エラー情報をマージします")
Public Sub MergeWith(ByVal other As ErrorInfo)
    If other Is Nothing Then Exit Sub
    
    ' 空の場合のみ更新
    If Len(Me.RootCause) = 0 Then Me.RootCause = other.RootCause
    If Len(Me.ImpactArea) = 0 Then Me.ImpactArea = other.ImpactArea
    If Len(Me.RemediationSteps) = 0 Then Me.RemediationSteps = other.RemediationSteps
    
    ' コレクションのマージ
    Dim item As Variant
    For Each item In other.RelatedResources
        If Not Contains(Me.RelatedResources, item) Then
            Me.RelatedResources.Add item
        End If
    Next
    
    For Each item In other.Tags
        If Not Contains(Me.Tags, item) Then
            Me.Tags.Add item
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