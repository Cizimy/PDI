VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StackTraceImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

Implements IStackTrace

'==========================
' 定数
'==========================
Private Const MODULE_NAME As String = "StackTraceImpl"
Private Const CONFIG_KEY_STACK_DEPTH As String = "MaxStackTraceDepth"

'==========================
' メンバ変数
'==========================
Private mStack As IStack
Private mConfig As IAppConfig
Private mLock As clsLock

'==========================
' 初期化・終了処理
'==========================
Private Sub Class_Initialize()
    Set mStack = New StackImpl
    Set mConfig = modConfig.GetAppConfig
    Set mLock = New clsLock
    
    ' スタックの最大サイズを設定から取得
    Dim maxDepth As Long
    maxDepth = mConfig.GetSetting(CONFIG_KEY_STACK_DEPTH, 10)  ' デフォルト値は10
    StackImpl(mStack).MaxSize = maxDepth
End Sub

Private Sub Class_Terminate()
    Set mStack = Nothing
    Set mConfig = Nothing
    Set mLock = Nothing
End Sub

'==========================
' IStackTrace実装
'==========================
Private Function IStackTrace_PushStackEntry(ByVal ClassName As String, ByVal MethodName As String, Optional ByVal Parameters As Variant) As Boolean
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    Dim entry As String
    entry = FormatStackEntry(ClassName, MethodName, Parameters)
    IStackTrace_PushStackEntry = mStack.Push(entry)
    
    mLock.ReleaseLock
    Exit Function

ErrorHandler:
    mLock.ReleaseLock
    Err.Raise Err.Number, MODULE_NAME, "PushStackEntry操作でエラーが発生しました: " & Err.Description
End Function

Private Function IStackTrace_PopStackEntry() As String
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    IStackTrace_PopStackEntry = mStack.Pop()
    
    mLock.ReleaseLock
    Exit Function

ErrorHandler:
    mLock.ReleaseLock
    Err.Raise Err.Number, MODULE_NAME, "PopStackEntry操作でエラーが発生しました: " & Err.Description
End Function

Private Function IStackTrace_GetStackTrace() As String
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    Dim tempStack As New StackImpl
    Dim trace As String
    Dim entry As String
    
    ' オリジナルスタックの内容を保持しながらトレースを生成
    Do While Not mStack.IsEmpty
        entry = mStack.Pop()
        tempStack.Push entry
        trace = entry & vbCrLf & trace
    Loop
    
    ' オリジナルスタックを復元
    Do While Not tempStack.IsEmpty
        mStack.Push tempStack.Pop()
    Loop
    
    IStackTrace_GetStackTrace = trace
    
    mLock.ReleaseLock
    Exit Function

ErrorHandler:
    mLock.ReleaseLock
    Err.Raise Err.Number, MODULE_NAME, "GetStackTrace操作でエラーが発生しました: " & Err.Description
End Function

Private Function IStackTrace_StackDepth() As Long
    IStackTrace_StackDepth = mStack.Count
End Function

Private Sub IStackTrace_ClearStack()
    mLock.AcquireLock
    mStack.Clear
    mLock.ReleaseLock
End Sub

Private Function IStackTrace_IsEmpty() As Boolean
    IStackTrace_IsEmpty = mStack.IsEmpty()
End Function

Private Function IStackTrace_Peek() As String
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    IStackTrace_Peek = mStack.Peek()
    
    mLock.ReleaseLock
    Exit Function

ErrorHandler:
    mLock.ReleaseLock
    Err.Raise Err.Number, MODULE_NAME, "Peek操作でエラーが発生しました: " & Err.Description
End Function

'==========================
' プライベートメソッド
'==========================
Private Function FormatStackEntry(ByVal ClassName As String, ByVal MethodName As String, ByVal Parameters As Variant) As String
    Dim result As String
    result = ClassName & "." & MethodName
    
    ' パラメータが指定されている場合は追加
    If Not IsMissing(Parameters) Then
        result = result & "("
        
        If IsArray(Parameters) Then
            Dim i As Long
            For i = LBound(Parameters) To UBound(Parameters)
                If i > LBound(Parameters) Then result = result & ", "
                result = result & CStr(Parameters(i))
            Next i
        Else
            result = result & CStr(Parameters)
        End If
        
        result = result & ")"
    End If
    
    FormatStackEntry = result
End Function

'==========================
' テストサポート（開発環境用）
'==========================
#If DEBUG Then
    Public Function ValidateState() As Boolean
        ValidateState = StackImpl(mStack).ValidateState()
    End Function
    
    Public Property Get MaxStackDepth() As Long
        MaxStackDepth = StackImpl(mStack).Capacity
    End Property
#End If