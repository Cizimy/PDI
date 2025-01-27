VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILogger

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "DatabaseLogger"

' ======================
' メンバ変数
' ======================
Private mConnectionPool As IConnectionPool
Private mLock As ILock
Private mPerformanceMonitor As IPerformanceMonitor
Private mDatabaseConfig As IDatabaseConfig
Private mTableName As String
Private Const MAX_RETRY_COUNT As Long = 3
Private Const RETRY_DELAY_MS As Long = 1000

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    ' 何もしない（依存性は外部から注入）
End Sub

Private Sub Class_Terminate()
    Set mConnectionPool = Nothing
    Set mLock = Nothing
    Set mPerformanceMonitor = Nothing
    Set mDatabaseConfig = Nothing
End Sub

' ======================
' 公開メソッド
' ======================
Public Sub Initialize(ByVal connectionPool As IConnectionPool, _
                     ByVal lock As ILock, _
                     ByVal performanceMonitor As IPerformanceMonitor, _
                     ByVal databaseConfig As IDatabaseConfig, _
                     ByVal tableName As String)
    Set mConnectionPool = connectionPool
    Set mLock = lock
    Set mPerformanceMonitor = performanceMonitor
    Set mDatabaseConfig = databaseConfig
    mTableName = tableName
End Sub

' ======================
' ILogger実装
' ======================
Private Sub ILogger_Log(ByVal logLevel As String, _
                       ByVal message As String, _
                       Optional ByVal stackTrace As String, _
                       Optional ByVal errorCode As Long)
    If mConnectionPool Is Nothing Then Exit Sub
    
    mPerformanceMonitor.StartMeasurement "DatabaseLogger_Log"
    mLock.AcquireLock
    
    Dim conn As Object
    Set conn = Nothing
    
    On Error GoTo ErrorHandler
    
    Dim retryCount As Long
    Do While retryCount < MAX_RETRY_COUNT
        If WriteLogToDatabase(logLevel, message, stackTrace, errorCode) Then Exit Do
        retryCount = retryCount + 1
        If retryCount < MAX_RETRY_COUNT Then Sleep RETRY_DELAY_MS
    Loop
    
Cleanup:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "DatabaseLogger_Log"
    Exit Sub
    
ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "データベースログ出力中にエラーが発生しました: " & Err.Description
        .Category = ECDatabase
        .Source = MODULE_NAME
        .ProcedureName = "ILogger_Log"
        .StackTrace = stackTrace
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    Resume Cleanup
End Sub

' ======================
' プライベートメソッド
' ======================
Private Function WriteLogToDatabase(ByVal logLevel As String, _
                                  ByVal message As String, _
                                  ByVal stackTrace As String, _
                                  ByVal errorCode As Long) As Boolean
    On Error Resume Next
    
    Dim conn As Object
    Set conn = mConnectionPool.AcquireConnection()
    
    If conn Is Nothing Then
        WriteLogToDatabase = False
        Exit Function
    End If
    
    ' コマンドオブジェクトの作成
    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        Set .ActiveConnection = conn
        .CommandText = "INSERT INTO " & mTableName & " " & _
                      "(LogTime, LogLevel, Message, StackTrace, ErrorCode) " & _
                      "VALUES (?, ?, ?, ?, ?)"
        .Parameters.Append .CreateParameter("LogTime", adDate, adParamInput, , Now)
        .Parameters.Append .CreateParameter("LogLevel", adVarChar, adParamInput, 50, logLevel)
        .Parameters.Append .CreateParameter("Message", adVarChar, adParamInput, 4000, message)
        .Parameters.Append .CreateParameter("StackTrace", adVarChar, adParamInput, 4000, stackTrace)
        .Parameters.Append .CreateParameter("ErrorCode", adInteger, adParamInput, , errorCode)
        .Execute
    End With
    
    WriteLogToDatabase = (Err.Number = 0)
    
    ' 接続の解放
    If Not conn Is Nothing Then
        mConnectionPool.ReleaseConnection conn
        Set conn = Nothing
    End If
    
    On Error GoTo 0
End Function