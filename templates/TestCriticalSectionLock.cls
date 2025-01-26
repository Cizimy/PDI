VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCriticalSectionLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "TestCriticalSectionLock"
Private testLock As clsCriticalSectionLock

'@TestInitialize
Private Sub TestInitialize()
    Set testLock = New clsCriticalSectionLock
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not testLock Is Nothing Then
        testLock.ForceRelease
        Set testLock = Nothing
    End If
End Sub

'@TestMethod("CriticalSectionLock")
Public Sub TestAcquireSuccess()
    On Error GoTo TestFail
    
    'Arrange
    Dim result As Boolean
    
    'Act
    result = testLock.Acquire()
    
    'Assert
    Debug.Assert result = True
    Debug.Assert testLock.ValidateLockState() = True
    Debug.Assert testLock.IsInitialized() = True
    
    Exit Sub
TestFail:
    Debug.Print "TestAcquireSuccess failed: " & Err.Description
End Sub

'@TestMethod("CriticalSectionLock")
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

'@TestMethod("CriticalSectionLock")
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

'@TestMethod("CriticalSectionLock")
Public Sub TestMultipleAcquire()
    On Error GoTo TestFail
    
    'Arrange
    Dim secondLock As clsCriticalSectionLock
    Set secondLock = New clsCriticalSectionLock
    
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

'@TestMethod("CriticalSectionLock")
Public Sub TestInitialization()
    On Error GoTo TestFail
    
    'Assert
    Debug.Assert testLock.IsInitialized() = True
    
    Exit Sub
TestFail:
    Debug.Print "TestInitialization failed: " & Err.Description
End Sub