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
Private mDatabaseConfig As IDatabaseConfig
Private mConnectionPool As ConnectionPool
Private mLock As clsLock
Private mTimerID As LongPtr
Private mPerformanceMonitor As clsPerformanceMonitor
Private mIsInitialized As Boolean
Private Const MAX_RETRY_COUNT As Long = 3
Private Const RETRY_DELAY_MS As Long = 1000
Private Const EMERGENCY_LOG_PATH As String = "\PDI_emergency.log"

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
    Set mDatabaseConfig = Nothing
    Set mSettings = Nothing
    
    mIsInitialized = False
End Sub

' ======================
' 公開メソッド
' ======================
Public Sub Configure(ByVal settings As ILoggerSettings, ByVal dbConfig As IDatabaseConfig)
    Set mSettings = settings
    Set mDatabaseConfig = dbConfig
    
    ' コネクションプールの初期化
    If Not mSettings Is Nothing And Not mDatabaseConfig Is Nothing Then
        mConnectionPool.Initialize mDatabaseConfig.GetConnectionString()
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
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "ログ出力中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "ILogger_Log"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
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

Private Property Let ILogger_LogDestination(ByVal destination As LogDestination)
    If Not mSettings Is Nothing Then
        mSettings.LogDestination = destination
    End If
End Property

Private Property Get ILogger_LogDestination() As LogDestination
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
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "ログキュー処理中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "ProcessLogQueue"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    Resume Cleanup
End Sub

Private Sub WriteToFile(ByVal logEntry As Variant)
    On Error GoTo ErrorHandler
    
    Dim retryCount As Long
    fDim ileNum As Integer
    Dim logMessage As String
    
    logMessage = Format$(logEntry(3), "yyyy/mm/dd hh:nn:ss") & " " & _
                 "[" & logEntry(0) & "] " & _
                 IIf(logEntry(2) <> 0, "(Code: " & logEntry(2) & ") ", "") & _
                 logEntry(1)
    
    Do While retryCount < MAX_RETRY_COUNT
        On Error Resume Next
        fileNum = FreeFile
        Open mSettings.LogFilePath For Append As #fileNum
        Print #fileNum, logMessage
        Close #fileNum
        
        If Err.Number = 0 Then
            Exit Sub
        End If
        
        retryCount = retryCount + 1
        If retryCount < MAX_RETRY_COUNT Then
            Sleep RETRY_DELAY_MS
        End If
    Loop
    
    On Error GoTo ErrorHandler
    ' ファイルへの書き込みが失敗した場合、代替出力先を試行
    WriteToFallbackDestination logMessage, "ファイルログ出力失敗"
    Exit Sub
    
ErrorHandler:
    WriteToFallbackDestination logMessage, "ファイルログ出力で予期せぬエラー: " & Err.Description
End Sub

Private Sub WriteToFallbackDestination(ByVal message As String, ByVal errorContext As String)
    On Error Resume Next
   
    M' イベントログへの出力を試みる
    modWindowsAPI.WriteToEventLog "PDI Logger", errorContext & vbCrLf & message, EVENTLOG_WARNING_TYPE
  
    oc' 緊急用ファイルへの出力を試みる
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
  
    = Dim emergencyPath As String
    emergencyPath = Environ$("TEMP") & EMERGENCY_LOG_PATH
    
    Dim ts As Object
    Set ts = fso.OpenTextFile(emergencyPath, 8, True)
    If Not ts Is Nothing Then
        ts.WriteLine Now & ": " & errorContext
        ts.WriteLine message
        ts.WriteLine String(50, "-")
        ts.Close
    End If
    Set fso = Nothing
End Sub

Private Sub WriteToDatabase(ByVal logEntry As Variant)
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Set conn = Nothing
    Dim retryCount As Long
    
    Do While retryCount < MAX_RETRY_COUNT
        On Error Resume Next
        mPerformanceMonitor.StartMeasurement "Database Connection"
        Set conn = mConnectionPool.GetConnection()
    
    
        If Not conn Is Nothing Then
            mPerformanceMonitor.StartMeasurement "Database Insert"
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
        
    
            If Err.Number = 0 Then
                mConnectionPool.ReleaseConnection conn
                mPerformanceMonitor.EndMeasurement "Database Connection"
                mPerformanceMonitor.EndMeasurement "Database Insert"
                    End If
        End If
        
        retryCount = retryCount + 1
        If retryCount < MAX_RETRY_COUNT Then
            If Not conn Is Nothing Then
            Exit Sub
    
        mConnectionPool.ReleaseConnection conn
                Set conn = Nothing
            End If
    
        Sleep RETRY_DELAY_MS
        End If
    Loop
    
    On Error GoTo ErrorHandler
    ' データベースへの書き込みが失敗した場合、ファイルに出力
    Dim logMessage As String
    logMessage = Format$(logEntry(3), "yyyy/mm/dd hh:nn:ss") & " " & _
                 "[" & logEntry(0) & "] " & _
                 IIf(logEntry(2) <> 0, "(Code: " & logEntry(2) & ") ", "") & _
                 logEntry(1)
    
    WriteToFallbackDestination logMessage, "データベースログ出力失敗"
    Exit Sub
    
ErrorHandler:
    If Not conn Is Nothing Then
        mPerformanceMonitor.EndMeasurement "Database Connection"
        mPerformanceMonitor.EndMeasurement "Database Insert"
        mConnectionPool.ReleaseConnection conn
    End If
    
    WriteToFallbackDestination logMessage, "データベースログ出力で予期せぬエラー: " & Err.Description
End Sub

Private Sub WriteToEventLog(ByVal logEntry As Variant)
    On Error GoTo ErrorHandler
    
    Dim objShell As Object
    Set objShell = Nothing
    
    Try:
        ' WScript.Shellオブジェクトの作成
        Set objShell = CreateObject("WScript.Shell")
        If objShell Is Nothing Then
            Err.Raise ErrEventLogCreateFailed, "WriteToEventLog", "WScript.Shellオブジェクトの作成に失敗しました。"
        End If
        
        ' イベントの種別を決定
        Dim eventType As Integer
        Select Case logEntry(2)
            Case LOG_LEVEL_ERROR, LOG_LEVEL_FATAL
                eventType = 1 ' ERROR
            Case LOG_LEVEL_WARNING
                eventType = 2 ' WARNING
            Case Else
                eventType = 4 ' INFORMATION
        End Select
    
    
        ' イベントログへの書き込み
        If Len(mSettings.LogEventSource) = 0 Then
            Err.Raise ErrEventLogSourceInvalid, "WriteToEventLog", "イベントログソースが指定されていません。"
        End If
        
        objShell.LogEvent eventType, logEntry(1), mSettings.LogEventSource
    
    Cleanup:
        If Not objShell Is Nothing Then Set objShell = Nothing
        Exit Sub
    
    ErrorHandler:
        Dim errDetail As ErrorInfo
        With errDetail
            .Code = IIf(Err.Number = ErrEventLogCreateFailed, ErrEventLogCreateFailed, ErrEventLogWriteFailed)
            .Description = "イベントログへの書き込みに失敗しました: " & Err.Description
            .Category = ECSystem
            .Source = MODULE_NAME
            .ProcedureName = "WriteToEventLog"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        modError.HandleError errDetail
        GoTo Cleanup
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