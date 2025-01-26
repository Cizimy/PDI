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

' ======================
' プライベート変数
' ======================
Private mConnectionString As String
Private mConfig As IDatabaseConfig
Private mLock As clsLock

Implements IErrorHandler

' ======================
' シングルトンインスタンス
' ======================
Private Shared instance As DatabaseConnectionErrorHandler

' ======================
' 初期化処理
' ======================
Public Sub Initialize(ByVal config As IDatabaseConfig)
    If config Is Nothing Then Exit Sub
    Set mConfig = config
    Set mLock = New clsLock
    mConnectionString = mConfig.GetConnectionString
End Sub


' ======================
' ファクトリメソッド
' ======================
Public Function Create() As DatabaseConnectionErrorHandler
    On Error GoTo ErrorHandler
    
    If instance Is Nothing Then
        Set instance = New DatabaseConnectionErrorHandler
        instance.Initialize config
    End If
    Set Create = instance
    Exit Function

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "データベースエラーハンドラの作成中にエラーが発生しました: " & Err.Description
        .Category = ECDatabase
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
Private Function IErrorHandler_HandleError(ByRef errorDetail As ErrorInfo) As Boolean
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
Private Function AttemptRetry(ByRef errorDetail As ErrorInfo) As Boolean
    On Error GoTo ErrorHandler
    
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
        Dim loggerSettings As DefaultLoggerSettings
        Set loggerSettings = New DefaultLoggerSettings
        loggerSettings.Initialize mConnectionString
        With New clsLogger
            .Configure loggerSettings
            .Log MODULE_NAME, _
                 "データベース接続リトライ " & retryCount & "/" & MAX_RETRY_COUNT, _
                 errorDetail.Code
        End With
    Next retryCount
    
    AttemptRetry = success
    Exit Function

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "データベース接続のリトライ中にエラーが発生しました: " & Err.Description
        .Category = ECDatabase
        .Source = MODULE_NAME
        .ProcedureName = "AttemptRetry"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    AttemptRetry = False
 ' エラー発生時は失敗として扱う
    Exit Function
End Function

Private Function TryDatabaseConnection() As Boolean
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
       conn.Open mConnectionString
    
    ' 接続テストのためのシンプルなクエリを実行
        conn.Execute "SELECT 1"
        
        conn.Close
        TryDatabaseConnection = True
        Exit Function
        
    ErrorHandler:
        Dim errNum As Long
        errNum = Err.Number
        
        ' エラー種別に応じた処理
        Select Case errNum
            Case -2147467259  ' ネットワークエラー
                LogSpecificError "ネットワーク接続エラー: サーバーに接続できません。", ECNetwork
            Case -2147217843  ' タイムアウト
                LogSpecificError "データベース接続がタイムアウトしました。", ECNetwork
            Case -2147217865  ' 認証エラー
                LogSpecificError "データベース認証エラー: ユーザー名またはパスワードが無効です。", ECSecurity
            Case Else
                LogSpecificError "データベース接続エラー: " & Err.Description, ECDatabase
        End Select
        
        If Not conn Is Nothing Then
            If conn.State = 1 Then  ' adStateOpen
                conn.Close
            End If
    End If
    Set conn = Nothing
    TryDatabaseConnection = False
        Exit Function
    End Function
    
    Private Sub LogSpecificError(ByVal message As String, ByVal category As ErrorCodeCategory)
        Dim errDetail As ErrorInfo
        With errDetail
            .Code = ErrDbConnectionFailed
            .Description = message
            .Category = category
            .Source = MODULE_NAME
            .ProcedureName = "TryDatabaseConnection"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        modError.HandleError errDetail
    End Sub

Private Sub LogError(ByRef errorDetail As ErrorInfo)
    On Error GoTo ErrorHandler
    
    mLock.AcquireLock
    
    ' エラーログの出力
    Dim loggerSettings As DefaultLoggerSettings
    Set loggerSettings = New DefaultLoggerSettings
    loggerSettings.Initialize mConnectionString
    With New clsLogger
        .Configure loggerSettings
    
    .Log MODULE_NAME, _
             "データベース接続エラー: " & errorDetail.Description, _
             errorDetail.Code
    End With
    
    mLock.ReleaseLock
    Exit Sub

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "データベースエラーのログ出力中にエラーが発生しました: " & Err.Description
        .Category = ECDatabase
        .Source = MODULE_NAME
        .ProcedureName = "LogError"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
End Sub

Private Sub NotifyUser(ByRef errorDetail As ErrorInfo, ByVal style As VbMsgBoxStyle)
    On Error GoTo ErrorHandler
    
    ' ユーザーへの通知
    Dim notifier As IUserNotifier
    Set notifier = New MsgBoxNotifier
    
    notifier.Notify errorDetail, style, "データベースエラー"
    Exit Sub

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "データベースエラーの通知中にエラーが発生しました: " & Err.Description
        .Category = ECDatabase
        .Source = MODULE_NAME
        .ProcedureName = "NotifyUser"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
End Sub

Private Sub Wait(ByVal milliseconds As Long)
    ' Sleep APIを使用して待機
    modWindowsAPI.Sleep milliseconds
End Sub