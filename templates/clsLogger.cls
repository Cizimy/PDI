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
Private mLogQueue As IQueue
Private mLock As ILock
Private mTimerID As LongPtr
Private mLoggers As Object ' Dictionary of ILogger instances
Private mIsProcessing As Boolean
Private mWorkerThread As Object
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
    Set mLoggers = CreateObject("Scripting.Dictionary")
    Set mLogQueue = New QueueImpl ' QueueImplはIQueueを実装
    Set mPerformanceMonitor = New clsPerformanceMonitor
End Sub

Private Sub Class_Terminate()
    If mTimerID <> 0 Then
        KillTimer 0, mTimerID
        mTimerID = 0
    End If
    
    If Not mWorkerThread Is Nothing Then
        ' ワーカースレッドの終了を待機
        mIsProcessing = False
        While mWorkerThread.Status = 1 ' 実行中
            Sleep 100
        Wend
    End If
    
    Set mLogQueue = Nothing
    Set mLock = Nothing
    Set mPerformanceMonitor = Nothing
    Set mLoggers = Nothing
    Set mSettings = Nothing
    
    mIsInitialized = False
End Sub

' ======================
' 公開メソッド
' ======================
Public Sub Configure(ByVal settings As ILoggerSettings, _
                    ByVal queue As IQueue, _
                    ByVal lock As ILock, _
                    ByVal performanceMonitor As IPerformanceMonitor)
    Set mSettings = settings
    Set mLogQueue = queue
    Set mLock = lock
    Set mPerformanceMonitor = performanceMonitor
    
    ' ログ出力先の設定
    InitializeLoggers
    
    ' ワーカースレッドの開始
    StartWorkerThread
    
    mIsInitialized = True
End Sub

Public Sub SetLogger(ByVal destination As String, ByVal logger As ILogger)
    If mLoggers.Exists(destination) Then
        mLoggers.Remove destination
    End If
    mLoggers.Add destination, logger
End Sub

Private Sub InitializeLoggers()
    ' 既存のロガーをクリア
    mLoggers.RemoveAll
    
    ' 設定に基づいてロガーを初期化
    Dim destinations As Variant
    destinations = mSettings.GetLogDestinations()
    
    Dim i As Long
    For i = LBound(destinations) To UBound(destinations)
        SetLogger destinations(i), CreateLogger(destinations(i))
    Next i
    
    ' タイマーの設定
    If mTimerID <> 0 Then
        KillTimer 0, mTimerID
    End If
    mTimerID = SetTimer(0, 0, mSettings.TimerInterval, AddressOf TimerProc)
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
        
        ' 各ロガーでログを処理
        Dim destination As Variant
        For Each destination In mLoggers.Keys
            If mSettings.ShouldLog(destination, logEntry(2)) Then
                Dim logger As ILogger
                Set logger = mLoggers(destination)
                logger.Log CStr(logEntry(2)), CStr(logEntry(1)), CStr(logEntry(3)), CLng(logEntry(0))
            End If
        Next destination
        
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

Private Function CreateLogger(ByVal destination As String) As ILogger
    Select Case destination
        Case "File"
            Dim fileLogger As New FileLogger
            fileLogger.Initialize mSettings.GetFileOperations(), _
                                mLock, _
                                mPerformanceMonitor, _
                                mSettings.LogFilePath
            Set CreateLogger = fileLogger
            
        Case "Database"
            Dim dbLogger As New DatabaseLogger
            dbLogger.Initialize mSettings.GetConnectionPool(), _
                              mLock, _
                              mPerformanceMonitor, _
                              mSettings.GetDatabaseConfig(), _
                              mSettings.LogTableName
            Set CreateLogger = dbLogger
            
        Case "EventLog"
            Dim eventLogger As New EventLogLogger
            eventLogger.Initialize mSettings.GetEventLog(), _
                                 mLock, _
                                 mPerformanceMonitor, _
                                 mSettings.LogEventSource
            Set CreateLogger = eventLogger
            
        Case "Console"
            Dim consoleLogger As New ConsoleLogger
            consoleLogger.Initialize mPerformanceMonitor, True
            Set CreateLogger = consoleLogger
            
        Case Else
            Err.Raise vbObjectError + 1, "CreateLogger", "不明なログ出力先: " & destination
    End Select
End Function

' ワーカースレッドの作成と開始
Private Sub StartWorkerThread()
    Set mWorkerThread = CreateObject("WScript.Shell").Exec( _
        "cmd /c powershell -Command {" & _
        "    while ($true) {" & _
        "        if (-not $global:isProcessing) { break }" & _
        "        Start-Sleep -Milliseconds " & mSettings.TimerInterval & _
        "    }" & _
        "}")
    mIsProcessing = True
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