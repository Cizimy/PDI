VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "TestLogger"

' ======================
' メンバ変数
' ======================
Private mLogger As clsLogger
Private mSettings As ILoggerSettings
Private mQueue As IQueue
Private mLock As ILock
Private mPerformanceMonitor As IPerformanceMonitor
Private mTestFileLogger As FileLogger
Private mTestDatabaseLogger As DatabaseLogger
Private mTestEventLogLogger As EventLogLogger
Private mTestConsoleLogger As ConsoleLogger

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    ' テスト対象の初期化
    Set mLogger = New clsLogger
    Set mSettings = New DefaultLoggerSettings
    Set mQueue = New QueueImpl
    Set mLock = New clsLock
    Set mPerformanceMonitor = New clsPerformanceMonitor
    
    ' テスト用ロガーの初期化
    Set mTestFileLogger = New FileLogger
    Set mTestDatabaseLogger = New DatabaseLogger
    Set mTestEventLogLogger = New EventLogLogger
    Set mTestConsoleLogger = New ConsoleLogger
End Sub

Private Sub Class_Terminate()
    Set mLogger = Nothing
    Set mSettings = Nothing
    Set mQueue = Nothing
    Set mLock = Nothing
    Set mPerformanceMonitor = Nothing
    Set mTestFileLogger = Nothing
    Set mTestDatabaseLogger = Nothing
    Set mTestEventLogLogger = Nothing
    Set mTestConsoleLogger = Nothing
End Sub

' ======================
' テストメソッド
' ======================
Public Sub TestInitialization()
    On Error GoTo TestFail
    
    ' 初期化前の状態確認
    Debug.Assert mLogger.GetQueueCount() = 0
    
    ' 初期化
    mLogger.Configure mSettings, mQueue, mLock, mPerformanceMonitor
    
    ' 初期化後の状態確認
    Debug.Assert Not mLogger Is Nothing
    Debug.Print "TestInitialization: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestInitialization: Failed - " & Err.Description
End Sub

Public Sub TestLoggerAddition()
    On Error GoTo TestFail
    
    ' ロガーの初期化
    mLogger.Configure mSettings, mQueue, mLock, mPerformanceMonitor
    
    ' ロガーの追加
    mLogger.SetLogger "File", mTestFileLogger
    mLogger.SetLogger "Database", mTestDatabaseLogger
    mLogger.SetLogger "EventLog", mTestEventLogLogger
    mLogger.SetLogger "Console", mTestConsoleLogger
    
    ' 同じ出力先に対する上書き確認
    Dim newFileLogger As New FileLogger
    mLogger.SetLogger "File", newFileLogger
    
    Debug.Print "TestLoggerAddition: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestLoggerAddition: Failed - " & Err.Description
End Sub

Public Sub TestLogMessageQueuing()
    On Error GoTo TestFail
    
    ' ロガーの初期化
    mLogger.Configure mSettings, mQueue, mLock, mPerformanceMonitor
    
    ' ログメッセージの追加
    mLogger.Log "TestModule", "Test message 1", LOG_LEVEL_INFO
    mLogger.Log "TestModule", "Test message 2", LOG_LEVEL_WARNING
    mLogger.Log "TestModule", "Test message 3", LOG_LEVEL_ERROR
    
    ' キューの状態確認
    Debug.Assert mLogger.GetQueueCount() = 3
    
    Debug.Print "TestLogMessageQueuing: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestLogMessageQueuing: Failed - " & Err.Description
End Sub

Public Sub TestAsyncProcessing()
    On Error GoTo TestFail
    
    ' ロガーの初期化
    mLogger.Configure mSettings, mQueue, mLock, mPerformanceMonitor
    
    ' ログメッセージの追加
    mLogger.Log "TestModule", "Async test message", LOG_LEVEL_INFO
    
    ' 非同期処理のシミュレーション
    mLogger.SimulateTimer
    
    ' キューが空になっていることを確認
    Debug.Assert mLogger.GetQueueCount() = 0
    
    Debug.Print "TestAsyncProcessing: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestAsyncProcessing: Failed - " & Err.Description
End Sub

' ======================
' 全テストの実行
' ======================
Public Sub RunAllTests()
    TestInitialization
    TestLoggerAddition
    TestLogMessageQueuing
    TestAsyncProcessing
End Sub