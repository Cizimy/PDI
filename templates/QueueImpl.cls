VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueueImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements IQueue

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "QueueImpl"

' ======================
' メンバ変数
' ======================
Private mItems As Collection

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mItems = New Collection
End Sub

Private Sub Class_Terminate()
    Set mItems = Nothing
End Sub

' ======================
' IQueue実装
' ======================
Private Sub IQueue_Enqueue(ByVal item As Variant)
    mItems.Add item
End Sub

Private Function IQueue_Dequeue() As Variant
    If mItems.Count = 0 Then
        Err.Raise vbObjectError + 1, MODULE_NAME, "キューが空です"
    End If
    
    If IsObject(mItems(1)) Then
        Set IQueue_Dequeue = mItems(1)
    Else
        IQueue_Dequeue = mItems(1)
    End If
    mItems.Remove 1
End Function

Private Function IQueue_IsEmpty() As Boolean
    IQueue_IsEmpty = (mItems.Count = 0)
End Function

Private Property Get IQueue_Count() As Long
    IQueue_Count = mItems.Count
End Property

Private Function IQueue_Peek() As Variant
    If mItems.Count = 0 Then
        Err.Raise vbObjectError + 1, MODULE_NAME, "キューが空です"
    End If
    
    If IsObject(mItems(1)) Then
        Set IQueue_Peek = mItems(1)
    Else
        IQueue_Peek = mItems(1)
    End If
End Function

Private Sub IQueue_Clear()
    Set mItems = New Collection
End Sub