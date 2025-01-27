VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestDatabaseLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "TestDatabaseLogger"

' ======================
' メンバ変数
' ======================
Private mDatabaseLogger As DatabaseLogger
Private mConnectionPool As IConnectionPool
Private mLock As ILock
Private mPerformanceMonitor As IPerformanceMonitor
Private mDatabaseConfig As IDatabaseConfig
Private mTestTableName As String

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    ' テスト対象の初期化
    Set mDatabaseLogger = New DatabaseLogger
    Set mConnectionPool = New ConnectionPool
    Set mLock = New clsLock
    Set mPerformanceMonitor = New clsPerformanceMonitor
    Set mDatabaseConfig = New ConfigImpl
    mTestTableName = "TestLog"
    
    ' テストテーブルの作成
    CreateTestTable
End Sub

Private Sub Class_Terminate()
    ' テストテーブルの削除
    DropTestTable
    
    Set mDatabaseLogger = Nothing
    Set mConnectionPool = Nothing
    Set mLock = Nothing
    Set mPerformanceMonitor = Nothing
    Set mDatabaseConfig = Nothing
End Sub

' ======================
' テストメソッド
' ======================
Public Sub TestInitialization()
    On Error GoTo TestFail
    
    ' 初期化
    mDatabaseLogger.Initialize mConnectionPool, mLock, mPerformanceMonitor, mDatabaseConfig, mTestTableName
    
    Debug.Print "TestInitialization: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestInitialization: Failed - " & Err.Description
End Sub

Public Sub TestLogWriting()
    On Error GoTo TestFail
    
    ' ロガーの初期化
    mDatabaseLogger.Initialize mConnectionPool, mLock, mPerformanceMonitor, mDatabaseConfig, mTestTableName
    
    ' ログの書き込み
    mDatabaseLogger.Log "INFO", "Test database message", "Test stack trace", 2001
    
    ' データベースの内容確認
    Dim conn As Object
    Set conn = mConnectionPool.AcquireConnection()
    
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM " & mTestTableName & " WHERE Message = 'Test database message'", conn
    
    Debug.Assert Not rs.EOF
    Debug.Assert rs.Fields("LogLevel").value = "INFO"
    Debug.Assert rs.Fields("StackTrace").value = "Test stack trace"
    Debug.Assert rs.Fields("ErrorCode").value = 2001
    
    rs.Close
    mConnectionPool.ReleaseConnection conn
    
    Debug.Print "TestLogWriting: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestLogWriting: Failed - " & Err.Description
End Sub

Public Sub TestConcurrentWriting()
    On Error GoTo TestFail
    
    ' ロガーの初期化
    mDatabaseLogger.Initialize mConnectionPool, mLock, mPerformanceMonitor, mDatabaseConfig, mTestTableName
    
    ' 複数のログメッセージを同時に書き込み
    Dim i As Long
    For i = 1 To 10
        mDatabaseLogger.Log "INFO", "Concurrent DB message " & i, "", i
    Next i
    
    ' データベースの内容確認
    Dim conn As Object
    Set conn = mConnectionPool.AcquireConnection()
    
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT COUNT(*) AS MessageCount FROM " & mTestTableName & _
           " WHERE Message LIKE 'Concurrent DB message%'", conn
    
    Debug.Assert rs.Fields("MessageCount").value = 10
    
    rs.Close
    mConnectionPool.ReleaseConnection conn
    
    Debug.Print "TestConcurrentWriting: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestConcurrentWriting: Failed - " & Err.Description
End Sub

Public Sub TestErrorHandling()
    On Error GoTo TestFail
    
    ' 無効なテーブル名での初期化
    mDatabaseLogger.Initialize mConnectionPool, mLock, mPerformanceMonitor, mDatabaseConfig, "InvalidTable"
    
    ' エラーが発生するはずのログ書き込み
    mDatabaseLogger.Log "ERROR", "Error test message", "", 2002
    
    Debug.Print "TestErrorHandling: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestErrorHandling: Failed - " & Err.Description
End Sub

' ======================
' ヘルパーメソッド
' ======================
Private Sub CreateTestTable()
    On Error Resume Next
    
    Dim conn As Object
    Set conn = mConnectionPool.AcquireConnection()
    
    ' テストテーブルの作成
    conn.Execute "CREATE TABLE " & mTestTableName & " (" & _
                "ID IDENTITY PRIMARY KEY, " & _
                "LogTime DATETIME, " & _
                "LogLevel VARCHAR(50), " & _
                "Message VARCHAR(4000), " & _
                "StackTrace VARCHAR(4000), " & _
                "ErrorCode INT)"
    
    mConnectionPool.ReleaseConnection conn
End Sub

Private Sub DropTestTable()
    On Error Resume Next
    
    Dim conn As Object
    Set conn = mConnectionPool.AcquireConnection()
    
    ' テストテーブルの削除
    conn.Execute "DROP TABLE " & mTestTableName
    
    mConnectionPool.ReleaseConnection conn
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