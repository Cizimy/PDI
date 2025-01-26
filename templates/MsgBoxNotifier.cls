VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MsgBoxNotifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "MsgBoxNotifier"

' ======================
' インターフェース実装
' ======================
Implements IUserNotifier

' ======================
' プライベート変数
' ======================
Private mNotificationStyle As VbMsgBoxStyle
Private mDefaultTitle As String
Private mLock As clsLock

' ======================
' 初期化
' ======================
Private Sub Class_Initialize()
    mNotificationStyle = vbOKOnly Or vbInformation
    Set mLock = New clsLock
    mDefaultTitle = "エラー"
End Sub

Private Sub Class_Terminate()
    Set mLock = Nothing
End Sub

' ======================
' IUserNotifier実装
' ======================
Private Function IUserNotifier_Notify(ByRef errorDetail As typErrorDetail, _
                                    Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
                                    Optional ByVal title As String = "エラー") As VbMsgBoxResult
    On Error GoTo ErrorHandler
    
    mLock.AcquireLock
    
    ' エラーの重大度に応じてアイコンを設定
    Dim icon As VbMsgBoxStyle
    icon = GetIconForErrorCategory(errorDetail.Category)
    
    ' メッセージの組み立て
    Dim msg As String
    msg = BuildErrorMessage(errorDetail)
    
    ' メッセージボックスの表示（NotificationStyleを適用）
    IUserNotifier_Notify = MsgBox(msg, mNotificationStyle Or buttons Or icon, title)
    
Cleanup:
    mLock.ReleaseLock
    Exit Function

    
ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "エラー通知の表示中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "IUserNotifier_Notify"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    
        ' エラーの種類に応じて適切な戻り値を設定
        Select Case errDetail.Category
            Case ECSystem, ECSecurity
                ' システムエラーやセキュリティエラーの場合は中断を示すvbCancelを返す
                IUserNotifier_Notify = vbCancel
            Case ECDatabase, ECNetwork
                ' データベースやネットワークエラーの場合はリトライを促すvbRetryを返す
                IUserNotifier_Notify = vbRetry
            Case Else
                ' その他のエラーの場合は一般的なvbAbortを返す
                IUserNotifier_Notify = vbAbort
        End Select
        
    Resume Cleanup
End Function

Private Property Let IUserNotifier_NotificationStyle(ByVal style As VbMsgBoxStyle)
    mNotificationStyle = style
End Property

Private Property Get IUserNotifier_NotificationStyle() As VbMsgBoxStyle
    IUserNotifier_NotificationStyle = mNotificationStyle
End Property

Private Property Let IUserNotifier_DefaultTitle(ByVal title As String)
    mDefaultTitle = title
End Property

Private Property Get IUserNotifier_DefaultTitle() As String
    IUserNotifier_DefaultTitle = mDefaultTitle
End Property

' ======================
' プライベートメソッド
' ======================
Private Function GetIconForErrorCategory(ByVal category As ErrorCodeCategory) As VbMsgBoxStyle
    Select Case category
        Case ECSystem, ECSecurity
            GetIconForErrorCategory = vbCritical
        Case ECDatabase, ECNetwork
            GetIconForErrorCategory = vbExclamation
        Case Else
            GetIconForErrorCategory = vbInformation
    End Select
End Function

Private Function BuildErrorMessage(ByRef errorDetail As typErrorDetail) As String
    Dim msg As String
    
    msg = "エラーが発生しました。" & vbCrLf & vbCrLf & _
          "説明: " & errorDetail.Description & vbCrLf & _
          "発生場所: " & errorDetail.Source & "." & errorDetail.ProcedureName
    
    ' スタックトレースが存在する場合は追加
    If Len(errorDetail.StackTrace) > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "スタックトレース:" & vbCrLf & _
              errorDetail.StackTrace
    End If
    
    BuildErrorMessage = msg
End Function