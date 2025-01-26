VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCriticalSectionLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILock

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "clsCriticalSectionLock"

' ======================
' Win32 API宣言
' ======================
Private Type CRITICAL_SECTION
    DebugInfo As Long
    LockCount As Long
    RecursionCount As Long
    OwningThread As Long
    LockSemaphore As Long
    SpinCount As Long
End Type

Private Declare Sub InitializeCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Private Declare Sub DeleteCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Private Declare Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Private Declare Function TryEnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION) As Long
Private Declare Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Private Declare Function GetTickCount Lib "kernel32" () As Long

' ======================
' メンバ変数
' ======================
Private criticalSection As CRITICAL_SECTION
Private isLocked As Boolean
Private isInitialized As Boolean

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    InitializeCriticalSection criticalSection
    isInitialized = True
    isLocked = False
End Sub

Private Sub Class_Terminate()
    If isLocked Then ILock_Release
    If isInitialized Then DeleteCriticalSection criticalSection
End Sub

' ======================
' ILock実装
' ======================
Private Function ILock_Acquire(Optional ByVal timeout As Long = -1) As Boolean
    On Error GoTo ErrorHandler
    
    If Not isInitialized Then Exit Function
    
    ' タイムアウトが指定されていない場合は即座に取得を試みる
    If timeout = -1 Then
        EnterCriticalSection criticalSection
        isLocked = True
        ILock_Acquire = True
        Exit Function
    End If
    
    ' タイムアウト付きの取得
    Dim startTime As Long
    startTime = GetTickCount()
    
    Do
        If TryEnterCriticalSection(criticalSection) <> 0 Then
            isLocked = True
            ILock_Acquire = True
            Exit Function
        End If
        
        ' タイムアウトチェック
        If timeout > 0 Then
            If GetTickCount() - startTime >= timeout Then Exit Do
        End If
        
        DoEvents ' 他のスレッドに実行機会を与える
    Loop
    
    Exit Function

ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrLockAcquireFailed
        .Description = "クリティカルセクションの取得に失敗しました: " & Err.Description
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
    
    LeaveCriticalSection criticalSection
    isLocked = False
    Exit Sub

ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrLockReleaseFailed
        .Description = "クリティカルセクションの解放に失敗しました: " & Err.Description
        .Category = ECSystem
        .Source = MODULE_NAME
        .ProcedureName = "ILock_Release"
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
    
    Public Sub ForceRelease()
        If isLocked Then ILock_Release
    End Sub
    
    Public Function IsInitialized() As Boolean
        IsInitialized = isInitialized
    End Function
#End If