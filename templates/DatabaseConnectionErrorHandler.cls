VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseConnectionErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "DatabaseConnectionErrorHandler"

' ======================
' 定数定義
' ======================
Private Const DEFAULT_RETRY_INTERVAL As Long = 1000 ' ミリ秒
Private Const MAX_RETRY_COUNT As Long = 3

Implements IErrorHandler

' ======================
' シングルトンインスタンス
' ======================
Private Shared instance As DatabaseConnectionErrorHandler

' ======================
' ファクトリメソッド
' ======================
Public Function Create() As DatabaseConnectionErrorHandler
    If instance Is Nothing Then
        Set instance = New DatabaseConnectionErrorHandler
    End If
    Set Create = instance
End Function

' ======================
' IErrorHandler実装
' ======================
Private Function IErrorHandler_HandleError(ByRef errorDetail As typErrorDetail) As Boolean
    Dim proceed As Boolean
    proceed = False  ' デフォルトでは処理を中断
    
    ' エラー情報をログに記録
    LogError errorDetail
    
    ' リトライ処理の実行
    If AttemptRetry(errorDetail) Then
        proceed = True  ' リトライ成功
    Else
        ' リトライ失敗時はユーザーに通知
        NotifyUser errorDetail, vbCritical
    End If
    
    IErrorHandler_HandleError = proceed
End Function

' ======================
' プライベートメソッド
' ======================
Private Function AttemptRetry(ByRef errorDetail As typErrorDetail) As Boolean
    Dim retryCount As Long
    Dim success As Boolean
    
    For retryCount = 1 To MAX_RETRY_COUNT
        ' リトライ間隔を設ける
        Wait DEFAULT_RETRY_INTERVAL
        
        If TryDatabaseConnection Then
            success = True
            Exit For
        End If
        
        ' リトライ試行をログに記録
        Dim logger As ILogger
        Set logger = New FileLogger
        logger.Log MODULE_NAME, _
                  "データベース接続リトライ " & retryCount & "/" & MAX_RETRY_COUNT, _
                  errorDetail.ErrorCode
    Next retryCount
    
    AttemptRetry = success
End Function

Private Function TryDatabaseConnection() As Boolean
    On Error Resume Next
    
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    ' 接続文字列は設定から取得
    conn.Open modConfig.Settings.DatabaseConnectionString
    
    If Err.Number = 0 Then
        conn.Close
        TryDatabaseConnection = True
    End If
    
    Set conn = Nothing
    On Error GoTo 0
End Function

Private Sub LogError(ByRef errorDetail As typErrorDetail)
    ' エラーログの出力
    Dim logger As ILogger
    Set logger = New FileLogger
    
    logger.Log MODULE_NAME, _
              "データベース接続エラー: " & errorDetail.Description, _
              errorDetail.ErrorCode
End Sub

Private Sub NotifyUser(ByRef errorDetail As typErrorDetail, ByVal style As VbMsgBoxStyle)
    ' ユーザーへの通知
    Dim notifier As IUserNotifier
    Set notifier = New MsgBoxNotifier
    
    notifier.Notify errorDetail, style, "データベースエラー"
End Sub

Private Sub Wait(ByVal milliseconds As Long)
    ' Sleep APIを使用して待機
    modWindowsAPI.Sleep milliseconds
End Sub