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

' ======================
' プライベート変数
' ======================
Private mLock As clsLock

Implements IErrorHandler

' ======================
' シングルトンインスタンス
' ======================
Private Shared instance As FileNotFoundErrorHandler

' ======================
' 初期化処理
' ======================
Private Sub Class_Initialize()
    Set mLock = New clsLock
End Sub

' ======================
' ファクトリメソッド
' ======================
Public Function Create() As FileNotFoundErrorHandler
    On Error GoTo ErrorHandler
    
    If instance Is Nothing Then
        Set instance = New FileNotFoundErrorHandler
    End If
    Set Create = instance
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "ファイルエラーハンドラの作成中にエラーが発生しました: " & Err.Description
        .Category = ECFileIO
        .Source = MODULE_NAME
        .ProcedureName = "Create"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    Set Create = Nothing
End Function

' ======================
' IErrorHandler実装
' ======================
Private Function IErrorHandler_HandleError(ByRef errorDetail As typErrorDetail) As Boolean
    Dim proceed As Boolean
    proceed = False  ' デフォルトでは処理を中断
    
    mLock.AcquireLock
    
    ' エラー情報をログに記録
    LogError errorDetail
    mLock.ReleaseLock
    
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
    On Error GoTo ErrorHandler
    
    mLock.AcquireLock
    
    ' エラーログの出力
    With New clsLogger
        .Configure New DefaultLoggerSettings
    
        .Log MODULE_NAME, _
             "ファイルが見つかりません: " & errorDetail.Description, _
             errorDetail.ErrorCode
    End With
    
    mLock.ReleaseLock
    Exit Sub

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "ファイルエラーのログ出力中にエラーが発生しました: " & Err.Description
        .Category = ECFileIO
        .Source = MODULE_NAME
        .ProcedureName = "LogError"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
End Sub

Private Sub NotifyUser(ByRef errorDetail As typErrorDetail, ByVal style As VbMsgBoxStyle)
    On Error GoTo ErrorHandler
    
    ' ユーザーへの通知
    Dim notifier As IUserNotifier
    Set notifier = New MsgBoxNotifier  ' または適切な通知実装を使用
    
    notifier.Notify errorDetail, style, "ファイルエラー"
    Exit Sub

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "ファイルエラーの通知中にエラーが発生しました: " & Err.Description
        .Category = ECFileIO
        .Source = MODULE_NAME
        .ProcedureName = "NotifyUser"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
End Sub