VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestSemaphoreLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "TestSemaphoreLock"
Private testLock As clsSemaphoreLock

'@TestInitialize
Private Sub TestInitialize()
    Set testLock = New clsSemaphoreLock
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not testLock Is Nothing Then
        testLock.ForceRelease
        Set testLock = Nothing
    End If
End Sub

'@TestMethod("SemaphoreLock")
Public Sub TestAcquireSuccess()
    On Error GoTo TestFail
    
    'Arrange
    Dim result As Boolean
    
    'Act
    result = testLock.Acquire()
    
    'Assert
    Debug.Assert result = True
    Debug.Assert testLock.ValidateLockState() = True
    
    Exit Sub
TestFail:
    Debug.Print "TestAcquireSuccess failed: " & Err.Description
End Sub

'@TestMethod("SemaphoreLock")
Public Sub TestAcquireWithTimeout()
    On Error GoTo TestFail
    
    'Arrange
    Dim result As Boolean
    Dim timeout As Long
    timeout = 1000 ' 1秒
    
    'Act
    result = testLock.Acquire(timeout)
    
    'Assert
    Debug.Assert result = True
    Debug.Assert testLock.ValidateLockState() = True
    
    Exit Sub
TestFail:
    Debug.Print "TestAcquireWithTimeout failed: " & Err.Description
End Sub

'@TestMethod("SemaphoreLock")
Public Sub TestReleaseSuccess()
    On Error GoTo TestFail
    
    'Arrange
    testLock.Acquire
    
    'Act
    testLock.Release
    
    'Assert
    Debug.Assert testLock.ValidateLockState() = False
    
    Exit Sub
TestFail:
    Debug.Print "TestReleaseSuccess failed: " & Err.Description
End Sub

'@TestMethod("SemaphoreLock")
Public Sub TestConfigure()
    On Error GoTo TestFail
    
    'Arrange
    Const NEW_SEMAPHORE_NAME As String = "TestSemaphore"
    
    'Act
    testLock.Configure NEW_SEMAPHORE_NAME
    
    'Assert
    Debug.Assert testLock.SemaphoreName = NEW_SEMAPHORE_NAME
    
    Exit Sub
TestFail:
    Debug.Print "TestConfigure failed: " & Err.Description
End Sub

'@TestMethod("SemaphoreLock")
Public Sub TestMultipleAcquire()
    On Error GoTo TestFail
    
    'Arrange
    Dim secondLock As clsSemaphoreLock
    Set secondLock = New clsSemaphoreLock
    
    'Act
    testLock.Acquire
    Dim result As Boolean
    result = secondLock.Acquire(100) ' 100msタイムアウト
    
    'Assert
    Debug.Assert result = False ' 2番目のロックは失敗するはず
    
    'Cleanup
    Set secondLock = Nothing
    
    Exit Sub
TestFail:
    Debug.Print "TestMultipleAcquire failed: " & Err.Description
End Sub