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
Private mLock As ILock
Private mLogger As ILogger
Private mEmergencyLogger As IEmergencyLogger
Private mConfig As IDatabaseConfig
Private mUserNotifier As IUserNotifier

Implements IErrorHandler

' =====================
' コンストラクタ
' =====================
Public Function Create( _
    ByVal lock As ILock, _
    ByVal logger As ILogger, _
    ByVal emergencyLogger As IEmergencyLogger, _
    ByVal config As IDatabaseConfig, _
    ByVal userNotifier As IUserNotifier _
) As DatabaseConnectionErrorHandler
    
    Dim instance As New DatabaseConnectionErrorHandler
    
    ' 依存性の注入
    With instance
        If lock Is Nothing Then
            Err.Raise 5, "DatabaseConnectionErrorHandler.Create", "ILock is required"
        End If
        Set .Lock = lock
        
        If logger Is Nothing Then
            Err.Raise 5, "DatabaseConnectionErrorHandler.Create", "ILogger is required"
        End If
        Set .Logger = logger
        
        If emergencyLogger Is Nothing Then
            Err.Raise 5, "DatabaseConnectionErrorHandler.Create", "IEmergencyLogger is required"
        End If
        Set .EmergencyLogger = emergencyLogger
        
        If config Is Nothing Then
            Err.Raise 5, "DatabaseConnectionErrorHandler.Create", "IDatabaseConfig is required"
        End If
        Set .Config = config
        
        If userNotifier Is Nothing Then
            Err.Raise 5, "DatabaseConnectionErrorHandler.Create", "IUserNotifier is required"
        End If
        Set .UserNotifier = userNotifier
    End With
    
    Set Create = instance
End Function

' プロパティ
Public Property Set Lock(ByVal value As ILock)
    Set mLock = value
End Property

Public Property Set Logger(ByVal value As ILogger)
    Set mLogger = value
End Property

Public Property Set EmergencyLogger(ByVal value As IEmergencyLogger)
    Set mEmergencyLogger = value
End Property

Public Property Set Config(ByVal value As IDatabaseConfig)
    Set mConfig = value
    mConnectionString = mConfig.GetConnectionString
End Property

Public Property Set UserNotifier(ByVal value As IUserNotifier)
    Set mUserNotifier = value
End Property

' ======================
' IErrorHandler実装
' ======================
Private Function IErrorHandler_HandleError(ByRef errorDetail As ErrorInfo) As Boolean
    Dim proceed As Boolean
    proceed = False  ' デフォルトでは処理を中断
    
    ' エラー情報をログに記録
    LogError errorDetail
    
    ' エラー情報の補完
    errorDetail.AddAdditionalInfo "ConnectionString", mConnectionString
    errorDetail.AddAdditionalInfo "MaxRetries", CStr(MAX_RETRY_COUNT)
    errorDetail.RecoveryAttempted = True
    
    ' リトライ処理の実行
    If AttemptRetry(errorDetail) Then
        proceed = True  ' リトライ成功
        errorDetail.RecoverySuccessful = True
    Else
        ' リトライ失敗時はユーザーに通知
        errorDetail.RecoverySuccessful = False
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
        
        errorDetail.AddAdditionalInfo "CurrentRetry", CStr(retryCount)
        
        If TryDatabaseConnection Then
            success = True
            Exit For
        End If
        
        ' リトライ試行をログに記録
        With mLogger
            .Log MODULE_NAME, "データベース接続リトライ " & retryCount & "/" & MAX_RETRY_COUNT, errorDetail.Code
        End With
    Next retryCount
    
    AttemptRetry = success
    Exit Function

ErrorHandler:
    Dim errDetail As ErrorInfo
    Set errDetail = New ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "データベース接続のリトライ中にエラーが発生しました: " & Err.Description
        .Category = ECDatabase
        .Source = MODULE_NAME
        .ProcedureName = "AttemptRetry"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
        .Severity = ESError
        Set .InnerError = errorDetail.Clone
    End With
    modError.HandleError errDetail
    AttemptRetry = False
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
    Set conn = Nothing
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
        Set conn = Nothing
    End If
    
    TryDatabaseConnection = False
    Exit Function
End Function

Private Sub LogSpecificError(ByVal message As String, ByVal category As ErrorCodeCategory)
    Dim errDetail As ErrorInfo
    Set errDetail = New ErrorInfo
    With errDetail
        .Code = ErrDbConnectionFailed
        .Description = message
        .Category = category
        .Source = MODULE_NAME
        .ProcedureName = "TryDatabaseConnection"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
        .Severity = ESError
    End With
    modError.HandleError errDetail
End Sub

Private Sub LogError(ByRef errorDetail As ErrorInfo)
    On Error GoTo ErrorHandler
    
    mLock.AcquireLock
    
    ' 通常のログ出力
    mLogger.Log MODULE_NAME, "データベース接続エラー: " & errorDetail.Description, errorDetail.Code
    ' 重要なエラーは緊急ログにも記録
    mEmergencyLogger.LogEmergency MODULE_NAME, errorDetail.ToString
    
    mLock.ReleaseLock
    Exit Sub

ErrorHandler:
    Dim errDetail As ErrorInfo
    Set errDetail = New ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "データベースエラーのログ出力中にエラーが発生しました: " & Err.Description
        .Category = ECDatabase
        .Source = MODULE_NAME
        .ProcedureName = "LogError"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
        .Severity = ESError
        Set .InnerError = errorDetail.Clone
    End With
    modError.HandleError errDetail
End Sub

Private Sub NotifyUser(ByRef errorDetail As ErrorInfo, ByVal style As VbMsgBoxStyle)
    On Error GoTo ErrorHandler
    
    mUserNotifier.Notify errorDetail, style, "データベースエラー"
    Exit Sub

ErrorHandler:
    Dim errDetail As ErrorInfo
    Set errDetail = New ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "データベースエラーの通知中にエラーが発生しました: " & Err.Description
        .Category = ECDatabase
        .Source = MODULE_NAME
        .ProcedureName = "NotifyUser"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
        .Severity = ESError
        Set .InnerError = errorDetail.Clone
    End With
    modError.HandleError errDetail
End Sub

Private Sub Wait(ByVal milliseconds As Long)
    ' Sleep APIを使用して待機
    modWindowsAPI.Sleep milliseconds
End Sub