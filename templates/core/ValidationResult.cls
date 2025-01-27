VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("バリデーション結果の詳細情報を保持するクラス")

' メンバー変数
Private mSuccess As Boolean              ' 検証成功フラグ
Private mErrorMessage As String          ' エラーメッセージ
Private mExecutionTime As Double         ' 実行時間（ミリ秒）
Private mResults As Collection           ' 個々のルール結果を格納
Private mValidationId As String          ' バリデーションID
Private mStartTime As Date               ' 開始時刻
Private mEndTime As Date                 ' 終了時刻
Private mRuleCounts As Collection        ' ルール別実行回数
Private mErrorDetails As Collection      ' エラー詳細情報
Private mPerformanceMetrics As Collection ' パフォーマンスメトリクス

' イベント
Public Event ResultUpdated(ByVal propertyName As String, ByVal oldValue As Variant, ByVal newValue As Variant)

Private Sub Class_Initialize()
    Set mResults = New Collection
    Set mRuleCounts = New Collection
    Set mErrorDetails = New Collection
    Set mPerformanceMetrics = New Collection
    mSuccess = False
    mErrorMessage = ""
    mExecutionTime = 0
    mValidationId = ""
    mStartTime = Now
    mEndTime = Now
End Sub

'@Description("検証が成功したかどうかを取得または設定します")
Public Property Get Success() As Boolean
    Success = mSuccess
End Property

Public Property Let Success(ByVal value As Boolean)
    Dim oldValue As Boolean
    oldValue = mSuccess
    mSuccess = value
    RaiseEvent ResultUpdated("Success", oldValue, value)
End Property

'@Description("エラーメッセージを取得または設定します")
Public Property Get ErrorMessage() As String
    ErrorMessage = mErrorMessage
End Property

Public Property Let ErrorMessage(ByVal value As String)
    Dim oldValue As String
    oldValue = mErrorMessage
    mErrorMessage = value
    RaiseEvent ResultUpdated("ErrorMessage", oldValue, value)
End Property

'@Description("実行時間を取得または設定します")
Public Property Get ExecutionTime() As Double
    ExecutionTime = mExecutionTime
End Property

Public Property Let ExecutionTime(ByVal value As Double)
    Dim oldValue As Double
    oldValue = mExecutionTime
    mExecutionTime = value
    RaiseEvent ResultUpdated("ExecutionTime", oldValue, value)
End Property

'@Description("個々のルール結果のコレクションを取得します")
Public Property Get Results() As Collection
    Set Results = mResults
End Property

'@Description("バリデーションIDを取得または設定します")
Public Property Get ValidationId() As String
    ValidationId = mValidationId
End Property

Public Property Let ValidationId(ByVal value As String)
    Dim oldValue As String
    oldValue = mValidationId
    mValidationId = value
    RaiseEvent ResultUpdated("ValidationId", oldValue, value)
End Property

'@Description("開始時刻を取得または設定します")
Public Property Get StartTime() As Date
    StartTime = mStartTime
End Property

Public Property Let StartTime(ByVal value As Date)
    Dim oldValue As Date
    oldValue = mStartTime
    mStartTime = value
    RaiseEvent ResultUpdated("StartTime", oldValue, value)
End Property

'@Description("終了時刻を取得または設定します")
Public Property Get EndTime() As Date
    EndTime = mEndTime
End Property

Public Property Let EndTime(ByVal value As Date)
    Dim oldValue As Date
    oldValue = mEndTime
    mEndTime = value
    RaiseEvent ResultUpdated("EndTime", oldValue, value)
End Property

'@Description("ルール別実行回数を取得します")
Public Property Get RuleCounts() As Collection
    Set RuleCounts = mRuleCounts
End Property

'@Description("エラー詳細情報を取得します")
Public Property Get ErrorDetails() As Collection
    Set ErrorDetails = mErrorDetails
End Property

'@Description("パフォーマンスメトリクスを取得します")
Public Property Get PerformanceMetrics() As Collection
    Set PerformanceMetrics = mPerformanceMetrics
End Property

'@Description("結果を追加します")
'@Param("result: 追加する結果")
Public Sub AddResult(ByVal result As ResultItem)
    mResults.Add result
    UpdateRuleCount result.RuleName
    If Not result.Success Then
        mErrorDetails.Add result
    End If
End Sub

'@Description("ルールの実行回数を更新します")
'@Param("ruleName: ルール名")
Private Sub UpdateRuleCount(ByVal ruleName As String)
    On Error Resume Next
    mRuleCounts(ruleName) = mRuleCounts(ruleName) + 1
    If Err.Number <> 0 Then
        mRuleCounts.Add 1, ruleName
    End If
    On Error GoTo 0
End Sub

'@Description("パフォーマンスメトリクスを追加します")
'@Param("metricName: メトリクス名")
'@Param("value: メトリクス値")
Public Sub AddPerformanceMetric(ByVal metricName As String, ByVal value As Double)
    Dim metric As New Dictionary
    metric.Add "Name", metricName
    metric.Add "Value", value
    metric.Add "Timestamp", Now
    mPerformanceMetrics.Add metric
End Sub

'@Description("結果をJSONフォーマットで取得します")
Public Function ToJSON() As String
    Dim json As String
    json = "{"
    json = json & """Success"": " & LCase(CStr(mSuccess)) & ","
    json = json & """ErrorMessage"": """ & JsonEscape(mErrorMessage) & ""","
    json = json & """ExecutionTime"": " & mExecutionTime & ","
    json = json & """ValidationId"": """ & mValidationId & ""","
    json = json & """StartTime"": """ & Format(mStartTime, "yyyy-mm-dd hh:nn:ss") & ""","
    json = json & """EndTime"": """ & Format(mEndTime, "yyyy-mm-dd hh:nn:ss") & """"
    json = json & "}"
    ToJSON = json
End Function

'@Description("JSON文字列のエスケープ処理を行います")
'@Param("text: エスケープする文字列")
Private Function JsonEscape(ByVal text As String) As String
    JsonEscape = Replace(Replace(Replace(text, "\", "\\"), """", "\"""), vbCrLf, "\n")
End Function

'@Description("結果をクリアします")
Public Sub Clear()
    Set mResults = New Collection
    Set mRuleCounts = New Collection
    Set mErrorDetails = New Collection
    Set mPerformanceMetrics = New Collection
    mSuccess = False
    mErrorMessage = ""
    mExecutionTime = 0
    mValidationId = ""
    mStartTime = Now
    mEndTime = Now
End Sub