VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILogger

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "clsLogger"

' ======================
' メンバ変数
' ======================
Private mSettings As ILoggerSettings
Private mLogQueue As Queue
Private mConnectionPool As ConnectionPool
Private mLock As clsLock
Private mTimerID As LongPtr
Private mPerformanceMonitor As clsPerformanceMonitor
Private mIsInitialized As Boolean

' ======================
' イベント
' ======================
Public Event Logged(ByVal logMessage As String, ByVal logLevel As LogLevel)

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mLogQueue = New Queue
    Set mConnectionPool = New ConnectionPool
    Set mLock = New clsLock
    Set mPerformanceMonitor = New clsPerformanceMonitor
End Sub

Private Sub Class_Terminate()
    If mTimerID <> 0 Then
        KillTimer 0, mTimerID
        mTimerID = 0
    End If
    
    Set mLogQueue = Nothing
    Set mConnectionPool = Nothing
    Set mLock = Nothing
    Set mPerformanceMonitor = Nothing
    Set mSettings = Nothing
    
    mIsInitialized = False
End Sub

' ======================
' 公開メソッド
' ======================
Public Sub Configure(ByVal settings As ILoggerSettings)
    Set mSettings = settings
    
    ' コネクションプールの初期化
    If Not mSettings Is Nothing Then
        mConnectionPool.Initialize mSettings.ConnectionString
    End If
    
    ' タイマーの設定
    If mTimerID <> 0 Then
        KillTimer 0, mTimerID
    End If
    mTimerID = SetTimer(0, 0, mSettings.TimerInterval, AddressOf TimerProc)
    
    mIsInitialized = True
End Sub

' ======================
' ILogger実装
' ======================
Private Sub ILogger_Log(ByVal moduleName As String, ByVal message As String, Optional ByVal errorCode As ErrorCode)
    If Not mIsInitialized Then Exit Sub
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.StartMeasurement "Logger_Log"
    End If
    
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    ' ログレベルチェック
    If errorCode >= mSettings.LogLevel Then
        ' ログメッセージをキューに追加
        Dim logEntry As Variant
        logEntry = Array(moduleName, message, errorCode, Now)
        mLogQueue.Enqueue logEntry
    End If
    
Cleanup:
    mLock.ReleaseLock
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "Logger_Log"
    End If
    Exit Sub
    
ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "ログ出力中にエラーが発生しました: " & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "ILogger_Log"
    End With
    modError.HandleError errDetail
    Resume Cleanup
End Sub

Private Property Let ILogger_LogLevel(ByVal level As LogLevel)
    If Not mSettings Is Nothing Then
        mSettings.LogLevel = level
    End If
End Property

Private Property Get ILogger_LogLevel() As LogLevel
    If Not mSettings Is Nothing Then
        ILogger_LogLevel = mSettings.LogLevel
    End If
End Property

Private Property Let ILogger_LogDestination(ByVal destination As LogDestinationEnum)
    If Not mSettings Is Nothing Then
        mSettings.LogDestination = destination
    End If
End Property

Private Property Get ILogger_LogDestination() As LogDestinationEnum
    If Not mSettings Is Nothing Then
        ILogger_LogDestination = mSettings.LogDestination
    End If
End Property

' ======================
' プライベートメソッド
' ======================
Private Sub ProcessLogQueue()
    If Not mIsInitialized Then Exit Sub
    
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    While mLogQueue.Count > 0
        Dim logEntry As Variant
        logEntry = mLogQueue.Dequeue
        
        ' ログ出力先に応じた処理
        Select Case mSettings.LogDestination
            Case LOG_DESTINATION_FILE
                WriteToFile logEntry
            Case LOG_DESTINATION_DATABASE
                WriteToDatabase logEntry
            Case LOG_DESTINATION_EVENTLOG
                WriteToEventLog logEntry
            Case LOG_DESTINATION_CONSOLE
                WriteToConsole logEntry
        End Select
        
        ' イベント発行
        RaiseEvent Logged(logEntry(1), logEntry(2))
    Wend
    
Cleanup:
    mLock.ReleaseLock
    Exit Sub
    
ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "ログキュー処理中にエラーが発生しました: " & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "ProcessLogQueue"
    End With
    modError.HandleError errDetail
    Resume Cleanup
End Sub

Private Sub WriteToFile(ByVal logEntry As Variant)
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open mSettings.LogFilePath For Append As #fileNum
    Print #fileNum, Format$(logEntry(3), "yyyy/mm/dd hh:nn:ss") & " " & _
                   "[" & logEntry(0) & "] " & _
                   IIf(logEntry(2) <> 0, "(Code: " & logEntry(2) & ") ", "") & _
                   logEntry(1)
    Close #fileNum
    Exit Sub
    
ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_FILEIO_ACCESS_DENIED
        .Description = "ファイルへのログ出力に失敗しました: " & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "WriteToFile"
    End With
    modError.HandleError errDetail
End Sub

Private Sub WriteToDatabase(ByVal logEntry As Variant)
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Set conn = mConnectionPool.GetConnection()
    
    If Not conn Is Nothing Then
        Dim cmd As Object
        Set cmd = CreateObject("ADODB.Command")
        With cmd
            .ActiveConnection = conn
            .CommandText = "INSERT INTO " & mSettings.LogTableName & " (LogTime, ModuleName, LogLevel, Message) VALUES (?, ?, ?, ?)"
            .Parameters.Append .CreateParameter("LogTime", 7, 1, , logEntry(3))
            .Parameters.Append .CreateParameter("ModuleName", 200, 1, 255, logEntry(0))
            .Parameters.Append .CreateParameter("LogLevel", 3, 1, , logEntry(2))
            .Parameters.Append .CreateParameter("Message", 200, 1, 4000, logEntry(1))
            .Execute
        End With
        
        mConnectionPool.ReleaseConnection conn
    End If
    Exit Sub
    
ErrorHandler:
    If Not conn Is Nothing Then mConnectionPool.ReleaseConnection conn
    
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_DATABASE_QUERY_FAILED
        .Description = "データベースへのログ出力に失敗しました: " & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "WriteToDatabase"
    End With
    modError.HandleError errDetail
End Sub

Private Sub WriteToEventLog(ByVal logEntry As Variant)
    On Error Resume Next
    
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    
    Dim eventType As Integer
    Select Case logEntry(2)
        Case LOG_LEVEL_ERROR, LOG_LEVEL_FATAL
            eventType = 1 ' ERROR
        Case LOG_LEVEL_WARNING
            eventType = 2 ' WARNING
        Case Else
            eventType = 4 ' INFORMATION
    End Select
    
    objShell.LogEvent eventType, logEntry(1), mSettings.LogEventSource
    Set objShell = Nothing
End Sub

Private Sub WriteToConsole(ByVal logEntry As Variant)
    Debug.Print Format$(logEntry(3), "yyyy/mm/dd hh:nn:ss") & " " & _
                "[" & logEntry(0) & "] " & _
                IIf(logEntry(2) <> 0, "(Code: " & logEntry(2) & ") ", "") & _
                logEntry(1)
End Sub

' ======================
' テストサポート機能
' 警告: これらのメソッドは開発時のテスト目的でのみ使用し、
' 本番環境では使用しないでください。
' ======================
#If DEBUG Then
    ''' <summary>
    ''' ログキューのエントリ数を取得（テスト用）
    ''' </summary>
    Private Function GetQueueCount() As Long
        GetQueueCount = mLogQueue.Count
    End Function
    
    ''' <summary>
    ''' ログキューをクリア（テスト用）
    ''' </summary>
    Private Sub ClearQueue()
        Set mLogQueue = New Queue
    End Sub
    
    ''' <summary>
    ''' パフォーマンスモニターの参照を取得（テスト用）
    ''' </summary>
    Private Function GetPerformanceMonitor() As clsPerformanceMonitor
        Set GetPerformanceMonitor = mPerformanceMonitor
    End Function
    
    ''' <summary>
    ''' タイマー処理をシミュレート（テスト用）
    ''' </summary>
    Private Sub SimulateTimer()
        ProcessLogQueue
    End Sub
#End If