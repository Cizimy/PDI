VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestFileLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "TestFileLogger"

' ======================
' メンバ変数
' ======================
Private mFileLogger As FileLogger
Private mFileOperations As IFileOperations
Private mLock As ILock
Private mPerformanceMonitor As IPerformanceMonitor
Private mTestFilePath As String

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    ' テスト対象の初期化
    Set mFileLogger = New FileLogger
    Set mFileOperations = New FileSystemOperations
    Set mLock = New clsLock
    Set mPerformanceMonitor = New clsPerformanceMonitor
    mTestFilePath = "test_log.txt"
End Sub

Private Sub Class_Terminate()
    ' テストファイルの削除
    On Error Resume Next
    mFileOperations.DeleteFile mTestFilePath
    On Error GoTo 0
    
    Set mFileLogger = Nothing
    Set mFileOperations = Nothing
    Set mLock = Nothing
    Set mPerformanceMonitor = Nothing
End Sub

' ======================
' テストメソッド
' ======================
Public Sub TestInitialization()
    On Error GoTo TestFail
    
    ' 初期化
    mFileLogger.Initialize mFileOperations, mLock, mPerformanceMonitor, mTestFilePath
    
    Debug.Print "TestInitialization: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestInitialization: Failed - " & Err.Description
End Sub

Public Sub TestLogWriting()
    On Error GoTo TestFail
    
    ' ロガーの初期化
    mFileLogger.Initialize mFileOperations, mLock, mPerformanceMonitor, mTestFilePath
    
    ' ログの書き込み
    mFileLogger.Log "INFO", "Test message", "Test stack trace", 1001
    
    ' ファイルの存在確認
    Debug.Assert mFileOperations.FileExists(mTestFilePath)
    
    ' ファイルの内容確認
    Dim content As String
    content = mFileOperations.ReadAllText(mTestFilePath)
    Debug.Assert InStr(content, "Test message") > 0
    Debug.Assert InStr(content, "Test stack trace") > 0
    Debug.Assert InStr(content, "1001") > 0
    
    Debug.Print "TestLogWriting: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestLogWriting: Failed - " & Err.Description
End Sub

Public Sub TestConcurrentWriting()
    On Error GoTo TestFail
    
    ' ロガーの初期化
    mFileLogger.Initialize mFileOperations, mLock, mPerformanceMonitor, mTestFilePath
    
    ' 複数のログメッセージを同時に書き込み
    Dim i As Long
    For i = 1 To 10
        mFileLogger.Log "INFO", "Concurrent message " & i, "", i
    Next i
    
    ' ファイルの内容確認
    Dim content As String
    content = mFileOperations.ReadAllText(mTestFilePath)
    For i = 1 To 10
        Debug.Assert InStr(content, "Concurrent message " & i) > 0
    Next i
    
    Debug.Print "TestConcurrentWriting: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestConcurrentWriting: Failed - " & Err.Description
End Sub

Public Sub TestErrorHandling()
    On Error GoTo TestFail
    
    ' 無効なファイルパスでの初期化
    mFileLogger.Initialize mFileOperations, mLock, mPerformanceMonitor, "|invalid|path"
    
    ' エラーが発生するはずのログ書き込み
    mFileLogger.Log "ERROR", "Error test message", "", 1002
    
    Debug.Print "TestErrorHandling: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestErrorHandling: Failed - " & Err.Description
End Sub

' ======================
' 全テストの実行
' ======================
Public Sub RunAllTests()
    TestInitialization
    TestLogWriting
    TestConcurrentWriting
    TestErrorHandling
End Sub