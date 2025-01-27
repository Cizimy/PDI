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

'@Description("バリデーション結果を格納するクラス")
'@Remarks("設定値の検証結果と詳細情報を保持")

' メンバ変数
Private mIsValid As Boolean
Private mErrors As Collection
Private mWarnings As Collection
Private mValidatedRules As Collection
Private mState As ValidationState
Private mStartTime As Date
Private mEndTime As Date
Private mValidatedSettingCount As Long

'@Description("クラスの初期化")
Private Sub Class_Initialize()
    Set mErrors = New Collection
    Set mWarnings = New Collection
    Set mValidatedRules = New Collection
    mIsValid = True
    mState = vsNotStarted
    mValidatedSettingCount = 0
End Sub

'@Description("クラスの終了処理")
Private Sub Class_Terminate()
    Set mErrors = Nothing
    Set mWarnings = Nothing
    Set mValidatedRules = Nothing
End Sub

'@Description("バリデーション結果が有効かどうか")
Public Property Get IsValid() As Boolean
    IsValid = mIsValid
End Property

Public Property Let IsValid(ByVal value As Boolean)
    mIsValid = value
End Property

'@Description("エラーメッセージのコレクション")
Public Property Get Errors() As Collection
    Set Errors = mErrors
End Property

Public Property Set Errors(ByVal value As Collection)
    Set mErrors = value
End Property

'@Description("警告メッセージのコレクション")
Public Property Get Warnings() As Collection
    Set Warnings = mWarnings
End Property

Public Property Set Warnings(ByVal value As Collection)
    Set mWarnings = value
End Property

'@Description("検証されたルールのコレクション")
Public Property Get ValidatedRules() As Collection
    Set ValidatedRules = mValidatedRules
End Property

Public Property Set ValidatedRules(ByVal value As Collection)
    Set mValidatedRules = value
End Property

'@Description("バリデーションの状態")
Public Property Get State() As ValidationState
    State = mState
End Property

Public Property Let State(ByVal value As ValidationState)
    mState = value
End Property

'@Description("バリデーション開始時刻")
Public Property Get StartTime() As Date
    StartTime = mStartTime
End Property

Public Property Let StartTime(ByVal value As Date)
    mStartTime = value
End Property

'@Description("バリデーション終了時刻")
Public Property Get EndTime() As Date
    EndTime = mEndTime
End Property

Public Property Let EndTime(ByVal value As Date)
    mEndTime = value
End Property

'@Description("検証された設定の数")
Public Property Get ValidatedSettingCount() As Long
    ValidatedSettingCount = mValidatedSettingCount
End Property

Public Property Let ValidatedSettingCount(ByVal value As Long)
    mValidatedSettingCount = value
End Property

'@Description("エラーを追加する")
'@Param("errorMessage: エラーメッセージ")
Public Sub AddError(ByVal errorMessage As String)
    mErrors.Add errorMessage
    mIsValid = False
End Sub

'@Description("警告を追加する")
'@Param("warningMessage: 警告メッセージ")
Public Sub AddWarning(ByVal warningMessage As String)
    mWarnings.Add warningMessage
End Sub

'@Description("検証ルールを追加する")
'@Param("ruleName: ルール名")
'@Param("ruleResult: ルールの検証結果")
Public Sub AddValidatedRule(ByVal ruleName As String, ByVal ruleResult As Boolean)
    Dim ruleInfo As New Dictionary
    ruleInfo.Add "Name", ruleName
    ruleInfo.Add "Result", ruleResult
    ruleInfo.Add "Timestamp", Now
    mValidatedRules.Add ruleInfo
End Sub

'@Description("バリデーション結果の要約を取得する")
'@Return("要約文字列")
Public Function GetSummary() As String
    Dim summary As String
    summary = "バリデーション結果:" & vbCrLf
    summary = summary & "- 有効性: " & IIf(mIsValid, "有効", "無効") & vbCrLf
    summary = summary & "- 状態: " & GetStateName(mState) & vbCrLf
    summary = summary & "- エラー数: " & mErrors.Count & vbCrLf
    summary = summary & "- 警告数: " & mWarnings.Count & vbCrLf
    summary = summary & "- 検証ルール数: " & mValidatedRules.Count & vbCrLf
    summary = summary & "- 検証設定数: " & mValidatedSettingCount & vbCrLf
    
    If mStartTime <> 0 And mEndTime <> 0 Then
        summary = summary & "- 実行時間: " & Format$(mEndTime - mStartTime, "hh:nn:ss") & vbCrLf
    End If
    
    GetSummary = summary
End Function

'@Description("バリデーション状態名を取得する")
'@Param("state: ValidationState列挙型の値")
'@Return("状態名")
Private Function GetStateName(ByVal state As ValidationState) As String
    Select Case state
        Case vsNotStarted: GetStateName = "未開始"
        Case vsPending: GetStateName = "保留中"
        Case vsRunning: GetStateName = "実行中"
        Case vsPaused: GetStateName = "一時停止"
        Case vsCompleted: GetStateName = "完了"
        Case vsFailed: GetStateName = "失敗"
        Case vsCancelled: GetStateName = "キャンセル"
        Case Else: GetStateName = "不明"
    End Select
End Function