VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMutexLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILock

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "clsMutexLock"

' ======================
' メンバ変数
' ======================
Private mutex As IMutex
Private mutexName As String
Private isLocked As Boolean
Private config As IAppConfig

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mutex = New MutexImpl
    Set config = modConfig.GetAppConfig
    ' ミューテックス名を設定から取得
    mutexName = config.GetSetting("MutexName", "Global\MyApp_ModCommon_Mutex")
    isLocked = False
End Sub

Private Sub Class_Terminate()
    If isLocked Then ILock_Release
    Set mutex = Nothing
    Set config = Nothing
End Sub

' ======================
' ILock実装
' ======================
Private Function ILock_Acquire(Optional ByVal timeout As Long = -1) As Boolean
    On Error GoTo ErrorHandler
    
    ' ミューテックスの作成
    If Not mutex.CreateMutex(False, mutexName) Then
        LogError "ミューテックスの作成に失敗しました。"
        Exit Function
    End If
    
    ' ロックの取得を試みる
    If mutex.WaitForSingleObject(timeout) Then
        isLocked = True
        ILock_Acquire = True
    End If
    Exit Function

ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrLockAcquireFailed
        .Description = "ロックの取得に失敗しました: " & Err.Description
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
    
    If Not mutex.ReleaseMutex() Then
        LogError "ロックの解放に失敗しました。"
    End If
    isLocked = False
    Exit Sub

ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrLockReleaseFailed
        .Description = "ロックの解放に失敗しました: " & Err.Description
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
Public Sub Configure(ByVal newMutexName As String)
    mutexName = newMutexName
End Sub

' ======================
' プライベートメソッド
' ======================
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
    
    Public Property Get MutexName() As String
        MutexName = mutexName
    End Property
    
    Public Sub ForceRelease()
        If isLocked Then ILock_Release
    End Sub
#End If