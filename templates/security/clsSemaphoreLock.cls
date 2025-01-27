VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSemaphoreLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILock

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "clsSemaphoreLock"

' ======================
' Win32 API宣言
' ======================
Private Declare Function CreateSemaphore Lib "kernel32" Alias "CreateSemaphoreA" ( _
    ByVal lpSemaphoreAttributes As Long, _
    ByVal lInitialCount As Long, _
    ByVal lMaximumCount As Long, _
    ByVal lpName As String) As Long

Private Declare Function ReleaseSemaphore Lib "kernel32" ( _
    ByVal hSemaphore As Long, _
    ByVal lReleaseCount As Long, _
    ByRef lpPreviousCount As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long

' ======================
' 定数定義
' ======================
Private Const WAIT_OBJECT_0 As Long = 0
Private Const WAIT_TIMEOUT As Long = &H102&
Private Const INFINITE As Long = -1

' ======================
' メンバ変数
' ======================
Private semaphoreHandle As Long
Private semaphoreName As String
Private isLocked As Boolean
Private config As IAppConfig

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set config = modConfig.GetAppConfig
    semaphoreName = config.GetSetting("SemaphoreName", "Global\MyApp_ModCommon_Semaphore")
    isLocked = False
    InitializeSemaphore
End Sub

Private Sub Class_Terminate()
    If isLocked Then ILock_Release
    If semaphoreHandle <> 0 Then CloseHandle semaphoreHandle
    Set config = Nothing
End Sub

' ======================
' ILock実装
' ======================
Private Function ILock_Acquire(Optional ByVal timeout As Long = -1) As Boolean
    On Error GoTo ErrorHandler
    
    If semaphoreHandle = 0 Then
        If Not InitializeSemaphore Then Exit Function
    End If
    
    Dim waitResult As Long
    waitResult = WaitForSingleObject(semaphoreHandle, timeout)
    
    Select Case waitResult
        Case WAIT_OBJECT_0
            isLocked = True
            ILock_Acquire = True
        Case WAIT_TIMEOUT
            ' タイムアウト - 何もしない
        Case Else
            LogError "セマフォの取得に失敗しました。"
    End Select
    Exit Function

ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrLockAcquireFailed
        .Description = "セマフォの取得に失敗しました: " & Err.Description
        .Category = ECSystem
        .Source = MODULE_NAME
        .ProcedureName = "ILock_Acquire"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errInfo
End Function

Private Sub ILock_Release()
    If Not isLocked Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Dim previousCount As Long
    If ReleaseSemaphore(semaphoreHandle, 1, previousCount) = 0 Then
        LogError "セマフォの解放に失敗しました。"
    Else
        isLocked = False
    End If
    Exit Sub

ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrLockReleaseFailed
        .Description = "セマフォの解放に失敗しました: " & Err.Description
        .Category = ECSystem
        .Source = MODULE_NAME
        .ProcedureName = "ILock_Release"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errInfo
End Sub

' ======================
' パブリックメソッド
' ======================
Public Sub Configure(ByVal newSemaphoreName As String)
    If isLocked Then ILock_Release
    If semaphoreHandle <> 0 Then CloseHandle semaphoreHandle
    semaphoreName = newSemaphoreName
    InitializeSemaphore
End Sub

' ======================
' プライベートメソッド
' ======================
Private Function InitializeSemaphore() As Boolean
    semaphoreHandle = CreateSemaphore(0&, 1, 1, semaphoreName)
    InitializeSemaphore = (semaphoreHandle <> 0)
End Function

Private Sub LogError(ByVal message As String)
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrUnexpected
        .Description = message
        .Category = ECSystem
        .Source = MODULE_NAME
        .ProcedureName = "Unknown"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errInfo
End Sub

' ======================
' テストサポート機能
' ======================
#If DEBUG Then
    Public Function ValidateLockState() As Boolean
        ValidateLockState = isLocked
    End Function
    
    Public Property Get SemaphoreName() As String
        SemaphoreName = semaphoreName
    End Property
    
    Public Sub ForceRelease()
        If isLocked Then ILock_Release
    End Sub
#End If