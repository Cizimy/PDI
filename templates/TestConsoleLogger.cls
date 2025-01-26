VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestConsoleLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "TestConsoleLogger"

' ======================
' メンバ変数
' ======================
Private mConsoleLogger As ConsoleLogger
Private mPerformanceMonitor As IPerformanceMonitor

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    ' テスト対象の初期化
    Set mConsoleLogger = New ConsoleLogger
    Set mPerformanceMonitor = New clsPerformanceMonitor
End Sub

Private Sub Class_Terminate()
    Set mConsoleLogger = Nothing
    Set mPerformanceMonitor = Nothing
End Sub

' ======================
' テストメソッド
' ======================
Public Sub TestInitialization()
    On Error GoTo TestFail
    
    ' 初期化
    mConsoleLogger.Initialize mPerformanceMonitor, True
    
    Debug.Print "TestInitialization: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestInitialization: Failed - " & Err.Description
End Sub

Public Sub TestLogWriting()
    On Error GoTo TestFail
    
    ' ロガーの初期化
    mConsoleLogger.Initialize mPerformanceMonitor, True
    
    ' ログの書き込み
    Debug.Print "=== TestLogWriting Start ==="
    mConsoleLogger.Log "INFO", "Test console message", "Test stack trace", 4001
    Debug.Print "=== TestLogWriting End ==="
    
    Debug.Print "TestLogWriting: OK - イミディエイトウィンドウで出力を確認してください"
    Exit Sub
    
TestFail:
    Debug.Print "TestLogWriting: Failed - " & Err.Description
End Sub

Public Sub TestDifferentLogLevels()
    On Error GoTo TestFail
    
    ' ロガーの初期化
    mConsoleLogger.Initialize mPerformanceMonitor, True
    
    ' 異なるログレベルでの書き込み
    Debug.Print "=== TestDifferentLogLevels Start ==="
    mConsoleLogger.Log "INFO", "Information message", "", 0
    mConsoleLogger.Log "WARNING", "Warning message", "", 4002
    mConsoleLogger.Log "ERROR", "Error message", "Error stack trace", 4003
    Debug.Print "=== TestDifferentLogLevels End ==="
    
    Debug.Print "TestDifferentLogLevels: OK - イミディエイトウィンドウで出力を確認してください"
    Exit Sub
    
TestFail:
    Debug.Print "TestDifferentLogLevels: Failed - " & Err.Description
End Sub

Public Sub TestPerformanceMonitoring()
    On Error GoTo TestFail
    
    ' ロガーの初期化
    mConsoleLogger.Initialize mPerformanceMonitor, True
    
    ' パフォーマンス計測付きのログ出力
    Debug.Print "=== TestPerformanceMonitoring Start ==="
    Dim i As Long
    For i = 1 To 100
        mConsoleLogger.Log "INFO", "Performance test message " & i, "", i
    Next i
    Debug.Print "=== TestPerformanceMonitoring End ==="
    
    ' パフォーマンス情報の確認
    Dim measurements As Collection
    Set measurements = mPerformanceMonitor.GetMeasurements("ConsoleLogger_Log")
    Debug.Assert Not measurements Is Nothing
    Debug.Assert measurements.Count > 0
    
    Debug.Print "TestPerformanceMonitoring: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestPerformanceMonitoring: Failed - " & Err.Description
End Sub

' ======================
' 全テストの実行
' ======================
Public Sub RunAllTests()
    TestInitialization
    TestLogWriting
    TestDifferentLogLevels
    TestPerformanceMonitoring
End Sub