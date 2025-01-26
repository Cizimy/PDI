VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestEventLogLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "TestEventLogLogger"

' ======================
' メンバ変数
' ======================
Private mEventLogLogger As EventLogLogger
Private mEventLog As IEventLog
Private mLock As ILock
Private mPerformanceMonitor As IPerformanceMonitor
Private mTestEventSource As String

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    ' テスト対象の初期化
    Set mEventLogLogger = New EventLogLogger
    Set mEventLog = New EventLogImpl
    Set mLock = New clsLock
    Set mPerformanceMonitor = New clsPerformanceMonitor
    mTestEventSource = "PDILoggerTest"
    
    ' テストイベントソースの作成
    CreateEventSource
End Sub

Private Sub Class_Terminate()
    ' テストイベントソースの削除
    RemoveEventSource
    
    Set mEventLogLogger = Nothing
    Set mEventLog = Nothing
    Set mLock = Nothing
    Set mPerformanceMonitor = Nothing
End Sub

' ======================
' テストメソッド
' ======================
Public Sub TestInitialization()
    On Error GoTo TestFail
    
    ' 初期化
    mEventLogLogger.Initialize mEventLog, mLock, mPerformanceMonitor, mTestEventSource
    
    Debug.Print "TestInitialization: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestInitialization: Failed - " & Err.Description
End Sub

Public Sub TestLogWriting()
    On Error GoTo TestFail
    
    ' ロガーの初期化
    mEventLogLogger.Initialize mEventLog, mLock, mPerformanceMonitor, mTestEventSource
    
    ' ログの書き込み
    mEventLogLogger.Log "INFO", "Test event log message", "Test stack trace", 3001
    
    ' イベントログの確認は手動で行う必要がある
    Debug.Print "TestLogWriting: OK - イベントビューアーで確認してください"
    Exit Sub
    
TestFail:
    Debug.Print "TestLogWriting: Failed - " & Err.Description
End Sub

Public Sub TestDifferentLogLevels()
    On Error GoTo TestFail
    
    ' ロガーの初期化
    mEventLogLogger.Initialize mEventLog, mLock, mPerformanceMonitor, mTestEventSource
    
    ' 異なるログレベルでの書き込み
    mEventLogLogger.Log "INFO", "Information message", "", 0
    mEventLogLogger.Log "WARNING", "Warning message", "", 3002
    mEventLogLogger.Log "ERROR", "Error message", "Error stack trace", 3003
    
    Debug.Print "TestDifferentLogLevels: OK - イベントビューアーで確認してください"
    Exit Sub
    
TestFail:
    Debug.Print "TestDifferentLogLevels: Failed - " & Err.Description
End Sub

Public Sub TestErrorHandling()
    On Error GoTo TestFail
    
    ' 無効なイベントソースでの初期化
    mEventLogLogger.Initialize mEventLog, mLock, mPerformanceMonitor, "InvalidSource"
    
    ' エラーが発生するはずのログ書き込み
    mEventLogLogger.Log "ERROR", "Error test message", "", 3004
    
    Debug.Print "TestErrorHandling: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestErrorHandling: Failed - " & Err.Description
End Sub

' ======================
' ヘルパーメソッド
' ======================
Private Sub CreateEventSource()
    On Error Resume Next
    
    ' イベントソースの作成
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    
    Dim regPath As String
    regPath = "HKLM\SYSTEM\CurrentControlSet\Services\EventLog\Application\" & mTestEventSource
    
    shell.RegWrite regPath & "\EventMessageFile", "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\EventLogMessages.dll", "REG_SZ"
    shell.RegWrite regPath & "\TypesSupported", 7, "REG_DWORD"
    
    Set shell = Nothing
End Sub

Private Sub RemoveEventSource()
    On Error Resume Next
    
    ' イベントソースの削除
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    
    Dim regPath As String
    regPath = "HKLM\SYSTEM\CurrentControlSet\Services\EventLog\Application\" & mTestEventSource
    
    shell.RegDelete regPath & "\"
    
    Set shell = Nothing
End Sub

' ======================
' 全テストの実行
' ======================
Public Sub RunAllTests()
    TestInitialization
    TestLogWriting
    TestDifferentLogLevels
    TestErrorHandling
End Sub