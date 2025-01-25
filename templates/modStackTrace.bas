Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "modStackTrace"

' ======================
' 定数定義
' ======================
Private Const MAX_STACK_TRACE_DEPTH As Long = 10 ' スタックトレースの最大深さ

' ======================
' プライベート変数
' ======================
Private mStack As Collection
Private mIsInitialized As Boolean

' ======================
' 初期化・終了処理
' ======================
Public Property Get IsInitialized() As Boolean
    IsInitialized = mIsInitialized
End Property

Public Sub InitializeModule()
    If mIsInitialized Then Exit Sub
    
    Set mStack = New Collection
    mIsInitialized = True
End Sub

Public Sub TerminateModule()
    If Not mIsInitialized Then Exit Sub
    
    Set mStack = Nothing
    mIsInitialized = False
End Sub

' ======================
' パブリックメソッド
' ======================
Public Sub PushStackEntry(ByVal ModuleName As String, ByVal ProcedureName As String)
    If Not mIsInitialized Then InitializeModule
    
    On Error Resume Next
    
    If mStack.Count < MAX_STACK_TRACE_DEPTH Then
        mStack.Add ModuleName & "." & ProcedureName
    End If
End Sub

Public Function PopStackEntry() As String
    If Not mIsInitialized Then Exit Function
    
    On Error Resume Next
    
    If mStack.Count > 0 Then
        PopStackEntry = mStack(mStack.Count)
        mStack.Remove mStack.Count
    End If
End Function

Public Function GetStackTrace() As String
    If Not mIsInitialized Then Exit Function
    
    On Error Resume Next
    
    Dim i As Long
    Dim trace As String
    
    For i = mStack.Count To 1 Step -1
        trace = trace & "  " & mStack(i) & vbCrLf
    Next i
    
    GetStackTrace = trace
End Function

Public Property Get StackDepth() As Long
    If Not mIsInitialized Then Exit Property
    StackDepth = mStack.Count
End Property

' ======================
' テストサポート機能（開発環境専用）
' 警告: これらのメソッドは開発時のテスト目的でのみ使用し、
' 本番環境では使用しないでください。
' ======================
#If DEBUG Then
    ''' <summary>
    ''' スタックの内容をクリア（テスト用）
    ''' </summary>
    Private Sub ClearStack()
        If Not mIsInitialized Then Exit Sub
        Set mStack = New Collection
    End Sub
    
    ''' <summary>
    ''' スタックの状態が有効かどうかを検証（テスト用）
    ''' </summary>
    ''' <returns>スタックの深さが最大値以下の場合True</returns>
    Private Function ValidateStackState() As Boolean
        If Not mIsInitialized Then Exit Function
        ValidateStackState = (mStack.Count <= MAX_STACK_TRACE_DEPTH)
    End Function
    
    ''' <summary>
    ''' モジュールの状態を初期化（テスト用）
    ''' </summary>
    Private Sub ResetModule()
        TerminateModule
        InitializeModule
    End Sub
#End If