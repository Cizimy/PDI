VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestConnectionPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "TestConnectionPool"

' モックオブジェクト
Private mMockConfig As MockDatabaseConfig
Private mMockPerformanceMonitor As MockPerformanceMonitor
Private mMockLock As MockLock
Private mMockLogger As MockLogger

' テスト対象
Private mConnectionPool As ConnectionPool

'*******************************************************************************
' テスト初期化
'*******************************************************************************
Private Sub Class_Initialize()
    ' モックオブジェクトの初期化
    Set mMockConfig = New MockDatabaseConfig
    Set mMockPerformanceMonitor = New MockPerformanceMonitor
    Set mMockLock = New MockLock
    Set mMockLogger = New MockLogger
    
    ' テスト対象の初期化
    Set mConnectionPool = New ConnectionPool
End Sub

'*******************************************************************************
' テストクリーンアップ
'*******************************************************************************
Private Sub Class_Terminate()
    Set mConnectionPool = Nothing
    Set mMockLogger = Nothing
    Set mMockLock = Nothing
    Set mMockPerformanceMonitor = Nothing
    Set mMockConfig = Nothing
End Sub

'*******************************************************************************
' テストケース：初期化時に依存オブジェクトが正しく設定されること
'*******************************************************************************
Public Sub TestInitialize()
    On Error GoTo TestFail
    
    ' 準備
    mMockConfig.SetConnectionString "Provider=SQLOLEDB;Data Source=TestServer;"
    mMockConfig.SetMaxPoolSize 5
    mMockConfig.SetRetryCount 3
    mMockConfig.SetRetryInterval 1000
    
    ' 実行
    mConnectionPool.Initialize mMockConfig, mMockPerformanceMonitor, mMockLock, mMockLogger
    
    ' 検証
    Debug.Assert mMockConfig.GetConnectionStringCallCount > 0
    Debug.Print "TestInitialize: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestInitialize: Failed - " & Err.Description
End Sub

'*******************************************************************************
' テストケース：不正な接続文字列の場合、エラーログが出力されること
'*******************************************************************************
Public Sub TestInvalidConnectionString()
    On Error GoTo TestFail
    
    ' 準備
    mMockConfig.SetConnectionString "Invalid Connection String"
    mConnectionPool.Initialize mMockConfig, mMockPerformanceMonitor, mMockLock, mMockLogger
    
    ' 実行
    Dim conn As Object
    Set conn = mConnectionPool.AcquireConnection
    
    ' 検証
    Debug.Assert mMockLogger.ErrorLogCount > 0
    Debug.Assert conn Is Nothing
    Debug.Print "TestInvalidConnectionString: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestInvalidConnectionString: Failed - " & Err.Description
End Sub

'*******************************************************************************
' テストケース：接続失敗時にリトライ処理が行われること
'*******************************************************************************
Public Sub TestConnectionRetry()
    On Error GoTo TestFail
    
    ' 準備
    mMockConfig.SetConnectionString "Provider=SQLOLEDB;Data Source=TestServer;"
    mMockConfig.SetRetryCount 3
    mMockConfig.SetRetryInterval 100
    mConnectionPool.Initialize mMockConfig, mMockPerformanceMonitor, mMockLock, mMockLogger
    
    ' 実行
    Dim conn As Object
    Set conn = mConnectionPool.AcquireConnection
    
    ' 検証
    Debug.Assert mMockConfig.GetRetryCountCallCount > 0
    Debug.Assert mMockConfig.GetRetryIntervalCallCount > 0
    Debug.Print "TestConnectionRetry: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestConnectionRetry: Failed - " & Err.Description
End Sub

'*******************************************************************************
' テストケース：プールサイズが設定値を超えないこと
'*******************************************************************************
Public Sub TestPoolSizeLimit()
    On Error GoTo TestFail
    
    ' 準備
    mMockConfig.SetConnectionString "Provider=SQLOLEDB;Data Source=TestServer;"
    mMockConfig.SetMaxPoolSize 2
    mConnectionPool.Initialize mMockConfig, mMockPerformanceMonitor, mMockLock, mMockLogger
    
    ' 実行
    Dim conn1 As Object, conn2 As Object, conn3 As Object
    Set conn1 = mConnectionPool.AcquireConnection
    Set conn2 = mConnectionPool.AcquireConnection
    Set conn3 = mConnectionPool.AcquireConnection ' これは失敗するはず
    
    ' 検証
    Debug.Assert Not conn1 Is Nothing
    Debug.Assert Not conn2 Is Nothing
    Debug.Assert conn3 Is Nothing
    Debug.Assert mConnectionPool.ActiveConnections = 2
    Debug.Print "TestPoolSizeLimit: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestPoolSizeLimit: Failed - " & Err.Description
End Sub

'*******************************************************************************
' テストケース：解放された接続が再利用可能になること
'*******************************************************************************
Public Sub TestConnectionReuse()
    On Error GoTo TestFail
    
    ' 準備
    mMockConfig.SetConnectionString "Provider=SQLOLEDB;Data Source=TestServer;"
    mMockConfig.SetMaxPoolSize 1
    mConnectionPool.Initialize mMockConfig, mMockPerformanceMonitor, mMockLock, mMockLogger
    
    ' 実行
    Dim conn1 As Object, conn2 As Object
    Set conn1 = mConnectionPool.AcquireConnection
    mConnectionPool.ReleaseConnection conn1
    Set conn2 = mConnectionPool.AcquireConnection
    
    ' 検証
    Debug.Assert Not conn2 Is Nothing
    Debug.Assert mConnectionPool.ActiveConnections = 1
    Debug.Print "TestConnectionReuse: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestConnectionReuse: Failed - " & Err.Description
End Sub

'*******************************************************************************
' 全テストの実行
'*******************************************************************************
Public Sub RunAllTests()
    TestInitialize
    TestInvalidConnectionString
    TestConnectionRetry
    TestPoolSizeLimit
    TestConnectionReuse
End Sub