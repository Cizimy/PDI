VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' 定数定義
Private Const INFINITE As Long = -1
Private Const WAIT_OBJECT_0 As Long = 0
Private Const ERROR_ALREADY_EXISTS As Long = 183&
Private Const MUTEX_NAME As String = "Global\MyApp_ModCommon_Mutex"

' メンバ変数
Private mMutexHandle As LongPtr
Private mIsLocked As Boolean

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    mMutexHandle = modWindowsAPI.CreateMutex(0&, 0&, MUTEX_NAME)
    mIsLocked = False
End Sub

Private Sub Class_Terminate()
    If mIsLocked Then modWindowsAPI.ReleaseMutex mMutexHandle
    If mMutexHandle <> 0 Then modWindowsAPI.CloseHandle mMutexHandle
End Sub

' ======================
' パブリックメソッド
' ======================
Public Sub AcquireLock()
    If mMutexHandle = 0 Then
        Err.Raise vbObjectError + 1, "clsLock", "Mutexの作成に失敗しました。"
    End If
    
    ' ロックの取得を試みる
    Dim result As Long
    result = modWindowsAPI.WaitForSingleObject(mMutexHandle, INFINITE)
    
    If result = WAIT_OBJECT_0 Then
        mIsLocked = True
    Else
        Err.Raise vbObjectError + 2, "clsLock", "ロックの取得に失敗しました。"
    End If
End Sub

Public Sub ReleaseLock()
    If mIsLocked Then
        If modWindowsAPI.ReleaseMutex(mMutexHandle) = 0 Then
            Err.Raise vbObjectError + 3, "clsLock", "ロックの解放に失敗しました。"
        End If
        mIsLocked = False
    End If
End Sub

' ======================
' プロパティ
' ======================
Public Property Get IsLocked() As Boolean
    IsLocked = mIsLocked
End Property

' ======================
' テストサポート機能
' ======================
#If DEBUG Then
    Public Function ValidateLockState() As Boolean
        ValidateLockState = (mMutexHandle <> 0)
    End Function
    
    Public Sub ForceRelease()
        If mMutexHandle <> 0 Then
            If mIsLocked Then modWindowsAPI.ReleaseMutex mMutexHandle
            modWindowsAPI.CloseHandle mMutexHandle
            mMutexHandle = 0
            mIsLocked = False
        End If
    End Sub
#End If