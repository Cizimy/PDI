VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileNotFoundErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "FileNotFoundErrorHandler"

Implements IErrorHandler

' ======================
' シングルトンインスタンス
' ======================
Private Shared instance As FileNotFoundErrorHandler

' ======================
' ファクトリメソッド
' ======================
Public Function Create() As FileNotFoundErrorHandler
    If instance Is Nothing Then
        Set instance = New FileNotFoundErrorHandler
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
    
    ' ユーザーに通知
    NotifyUser errorDetail, vbExclamation
    
    ' 代替ファイルパスの確認や、ファイル作成の試行など、
    ' ファイル不在時の回復処理をここに実装可能
    
    IErrorHandler_HandleError = proceed
End Function

' ======================
' プライベートメソッド
' ======================
Private Sub LogError(ByRef errorDetail As typErrorDetail)
    ' エラーログの出力
    Dim logger As ILogger
    Set logger = New FileLogger  ' または適切なログ実装を使用
    
    logger.Log MODULE_NAME, _
              "ファイルが見つかりません: " & errorDetail.Description, _
              errorDetail.ErrorCode
End Sub

Private Sub NotifyUser(ByRef errorDetail As typErrorDetail, ByVal style As VbMsgBoxStyle)
    ' ユーザーへの通知
    Dim notifier As IUserNotifier
    Set notifier = New MsgBoxNotifier  ' または適切な通知実装を使用
    
    notifier.Notify errorDetail, style, "ファイルエラー"
End Sub