VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILock

' モック用の内部状態
Private mIsLocked As Boolean
Private mAcquireLockCallCount As Long
Private mReleaseLockCallCount As Long
Private mShouldFailAcquire As Boolean
Private mShouldFailRelease As Boolean

'*******************************************************************************
' 初期化
'*******************************************************************************
Private Sub Class_Initialize()
    mIsLocked = False
    mAcquireLockCallCount = 0
    mReleaseLockCallCount = 0
    mShouldFailAcquire = False
    mShouldFailRelease = False
End Sub

'*******************************************************************************
' ILock インターフェースの実装
'*******************************************************************************
Private Sub ILock_AcquireLock()
    mAcquireLockCallCount = mAcquireLockCallCount + 1
    
    If mShouldFailAcquire Then
        Err.Raise vbObjectError + 1001, "MockLock", "ロックの取得に失敗しました"
    End If
    
    If mIsLocked Then
        Err.Raise vbObjectError + 1002, "MockLock", "既にロックされています"
    End If
    
    mIsLocked = True
End Sub

Private Sub ILock_ReleaseLock()
    mReleaseLockCallCount = mReleaseLockCallCount + 1
    
    If mShouldFailRelease Then
        Err.Raise vbObjectError + 1003, "MockLock", "ロックの解放に失敗しました"
    End If
    
    If Not mIsLocked Then
        Err.Raise vbObjectError + 1004, "MockLock", "ロックされていません"
    End If
    
    mIsLocked = False
End Sub

'*******************************************************************************
' モック用のヘルパーメソッド
'*******************************************************************************
Public Property Get IsLocked() As Boolean
    IsLocked = mIsLocked
End Property

Public Property Get AcquireLockCallCount() As Long
    AcquireLockCallCount = mAcquireLockCallCount
End Property

Public Property Get ReleaseLockCallCount() As Long
    ReleaseLockCallCount = mReleaseLockCallCount
End Property

Public Sub SetShouldFailAcquire(ByVal value As Boolean)
    mShouldFailAcquire = value
End Sub

Public Sub SetShouldFailRelease(ByVal value As Boolean)
    mShouldFailRelease = value
End Sub

Public Sub ResetCallCounts()
    mAcquireLockCallCount = 0
    mReleaseLockCallCount = 0
End Sub

Public Sub Reset()
    mIsLocked = False
    ResetCallCounts
    mShouldFailAcquire = False
    mShouldFailRelease = False
End Sub

'*******************************************************************************
' テスト用のヘルパーメソッド
'*******************************************************************************
Public Function IsLockBalanced() As Boolean
    IsLockBalanced = (mAcquireLockCallCount = mReleaseLockCallCount) And (Not mIsLocked)
End Function

Public Function GetUnbalancedLockCount() As Long
    GetUnbalancedLockCount = mAcquireLockCallCount - mReleaseLockCallCount
End Function