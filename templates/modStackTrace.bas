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
Private stack As Collection
Private isInitialized As Boolean

' ======================
' 初期化・終了処理
' ======================
Public Property Get IsInitialized() As Boolean
    IsInitialized = isInitialized
End Property

Public Sub InitializeModule()
    If isInitialized Then Exit Sub
    
    Set stack = New Collection
    isInitialized = True
End Sub

Public Sub TerminateModule()
    If Not isInitialized Then Exit Sub
    
    Set stack = Nothing
    isInitialized = False
End Sub

' ======================
' パブリックメソッド
' ======================
Public Sub PushStackEntry(ByVal ModuleName As String, ByVal ProcedureName As String)
    If Not isInitialized Then InitializeModule
    
    On Error GoTo ErrorHandler
    
    If stack.Count < MAX_STACK_TRACE_DEPTH Then
        stack.Add ModuleName & "." & ProcedureName
    End If
    Exit Sub

ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrStackTracePushFailed
        .Description = "スタックトレースへのエントリ追加に失敗しました。"
        .Category = ECSystem
        .Source = MODULE_NAME
        .ProcedureName = "PushStackEntry"
        .StackTrace = "モジュール: " & ModuleName & ", プロシージャ: " & ProcedureName
        .OccurredAt = Now
    End With
    modError.HandleError errInfo
End Sub

Public Function PopStackEntry() As String
    If Not isInitialized Then Exit Function
    
    On Error GoTo ErrorHandler
    
    Dim result As String
    If stack.Count > 0 Then
        result = stack(stack.Count)
        stack.Remove stack.Count
        PopStackEntry = result
    End If
    Exit Function

ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrStackTracePopFailed
        .Description = "スタックトレースからのエントリ取得に失敗しました。"
        .Category = ECSystem
        .Source = MODULE_NAME
        .ProcedureName = "PopStackEntry"
        .OccurredAt = Now
    End With
    modError.HandleError errInfo
End Function

Public Function GetStackTrace() As String
    If Not isInitialized Then Exit Function
    
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim trace As String
    
    For i = stack.Count To 1 Step -1
        trace = trace & "  " & stack(i) & vbCrLf
    Next i
    
    GetStackTrace = trace
    Exit Function

ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrStackTraceGetFailed
        .Description = "スタックトレース文字列の生成に失敗しました。"
        .Category = ECSystem
        .Source = MODULE_NAME
        .ProcedureName = "GetStackTrace"
        .OccurredAt = Now
    End With
    modError.HandleError errInfo
    GetStackTrace = "スタックトレースの取得に失敗しました。"
End Function

Public Property Get StackDepth() As Long
    If Not isInitialized Then Exit Property
    StackDepth = stack.Count
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
        If Not isInitialized Then Exit Sub
        Set stack = New Collection
    End Sub
    
    ''' <summary>
    ''' スタックの状態が有効かどうかを検証（テスト用）
    ''' </summary>
    ''' <returns>スタックの深さが最大値以下の場合True</returns>
    Private Function ValidateStackState() As Boolean
        If Not isInitialized Then Exit Function
        ValidateStackState = (stack.Count <= MAX_STACK_TRACE_DEPTH)
    End Function
    
    ''' <summary>
    ''' モジュールの状態を初期化（テスト用）
    ''' </summary>
    Private Sub ResetModule()
        TerminateModule
        InitializeModule
    End Sub
#End If