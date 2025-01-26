VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StackImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

Implements IStack

'==========================
' 定数
'==========================
Private Const DEFAULT_CAPACITY As Long = 10
Private Const MODULE_NAME As String = "StackImpl"

'==========================
' メンバ変数
'==========================
Private mItems() As String
Private mCount As Long
Private mMaxSize As Long
Private mLock As clsLock

'==========================
' 初期化・終了処理
'==========================
Private Sub Class_Initialize()
    mMaxSize = DEFAULT_CAPACITY
    ReDim mItems(0 To mMaxSize - 1)
    Set mLock = New clsLock
End Sub

Private Sub Class_Terminate()
    Set mLock = Nothing
End Sub

'==========================
' プロパティ
'==========================
Public Property Let MaxSize(ByVal value As Long)
    If value <= 0 Then
        Err.Raise vbObjectError + 1, MODULE_NAME, "最大サイズは1以上である必要があります。"
    End If
    mMaxSize = value
    ReDim Preserve mItems(0 To mMaxSize - 1)
End Property

'==========================
' IStack実装
'==========================
Private Function IStack_Push(ByVal Item As String) As Boolean
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    If mCount >= mMaxSize Then
        IStack_Push = False
        mLock.ReleaseLock
        Exit Function
    End If
    
    mItems(mCount) = Item
    mCount = mCount + 1
    IStack_Push = True
    
    mLock.ReleaseLock
    Exit Function

ErrorHandler:
    mLock.ReleaseLock
    Err.Raise Err.Number, MODULE_NAME, "Push操作でエラーが発生しました: " & Err.Description
End Function

Private Function IStack_Pop() As String
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    If mCount = 0 Then
        IStack_Pop = ""
        mLock.ReleaseLock
        Exit Function
    End If
    
    mCount = mCount - 1
    IStack_Pop = mItems(mCount)
    
    mLock.ReleaseLock
    Exit Function

ErrorHandler:
    mLock.ReleaseLock
    Err.Raise Err.Number, MODULE_NAME, "Pop操作でエラーが発生しました: " & Err.Description
End Function

Private Function IStack_Peek() As String
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    If mCount = 0 Then
        IStack_Peek = ""
        mLock.ReleaseLock
        Exit Function
    End If
    
    IStack_Peek = mItems(mCount - 1)
    
    mLock.ReleaseLock
    Exit Function

ErrorHandler:
    mLock.ReleaseLock
    Err.Raise Err.Number, MODULE_NAME, "Peek操作でエラーが発生しました: " & Err.Description
End Function

Private Function IStack_IsEmpty() As Boolean
    IStack_IsEmpty = (mCount = 0)
End Function

Private Property Get IStack_Count() As Long
    IStack_Count = mCount
End Property

Private Sub IStack_Clear()
    mLock.AcquireLock
    mCount = 0
    mLock.ReleaseLock
End Sub

'==========================
' テストサポート（開発環境用）
'==========================
#If DEBUG Then
    Public Function ValidateState() As Boolean
        ValidateState = (mCount >= 0 And mCount <= mMaxSize)
    End Function
    
    Public Property Get Capacity() As Long
        Capacity = mMaxSize
    End Property
#End If