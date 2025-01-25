VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCallStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "clsCallStack"

' ======================
' 定数定義
' ======================
Private Const MAX_STACK_TRACE_DEPTH As Long = 10 ' スタックトレースの最大深さ

' ======================
' メンバ変数
' ======================
Private mStack As Collection
Private mPerformanceMonitor As clsPerformanceMonitor
Private mIsInitialized As Boolean

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mStack = New Collection
    Set mPerformanceMonitor = New clsPerformanceMonitor
    mIsInitialized = True
End Sub

Private Sub Class_Terminate()
    Set mStack = Nothing
    Set mPerformanceMonitor = Nothing
    mIsInitialized = False
End Sub

' ======================
' パブリックメソッド
' ======================

''' <summary>
''' 呼び出し履歴にエントリを追加
''' </summary>
''' <param name="ModuleName">モジュール名</param>
''' <param name="ProcedureName">プロシージャ名</param>
''' <remarks>
''' スタックトレースの最大深さを超えた場合、エラーを発生させずに無視します
''' </remarks>
Public Sub Push(ByVal ModuleName As String, ByVal ProcedureName As String)
    If Not mIsInitialized Then Exit Sub
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.StartMeasurement "CallStack_Push"
    End If
    
    On Error GoTo ErrorHandler
    
    If mStack.Count < MAX_STACK_TRACE_DEPTH Then
        mStack.Add ModuleName & "." & ProcedureName
    End If
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "CallStack_Push"
    End If
    Exit Sub

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "Push操作中にエラーが発生しました: " & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "Push"
    End With
    modError.HandleError errDetail
    Resume Next
End Sub

''' <summary>
''' 呼び出し履歴から最後のエントリを取得して削除
''' </summary>
''' <returns>最後に追加されたプロシージャの完全修飾名、またはスタックが空の場合は空文字列</returns>
Public Function Pop() As String
    If Not mIsInitialized Then Exit Function
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.StartMeasurement "CallStack_Pop"
    End If
    
    On Error GoTo ErrorHandler
    
    If mStack.Count > 0 Then
        Pop = mStack(mStack.Count)
        mStack.Remove mStack.Count
    End If
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "CallStack_Pop"
    End If
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "Pop操作中にエラーが発生しました: " & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "Pop"
    End With
    modError.HandleError errDetail
    Resume Next
End Function

''' <summary>
''' 現在のスタックトレースを文字列として取得
''' </summary>
''' <returns>スタックトレースの文字列表現</returns>
Public Property Get StackTrace() As String
    If Not mIsInitialized Then Exit Property
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.StartMeasurement "CallStack_GetTrace"
    End If
    
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim trace As String
    
    For i = mStack.Count To 1 Step -1
        trace = trace & "  " & mStack(i) & vbCrLf
    Next i
    
    StackTrace = trace
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "CallStack_GetTrace"
    End If
    Exit Property

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "スタックトレース取得中にエラーが発生しました: " & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "StackTrace"
    End With
    modError.HandleError errDetail
    Resume Next
End Property

''' <summary>
''' スタック内のエントリ数を取得
''' </summary>
''' <returns>現在のスタックの深さ</returns>
Public Property Get Count() As Long
    Count = mStack.Count
End Property

' ======================
' テストサポート機能
' ======================
#If DEBUG Then
    ''' <summary>
    ''' スタックの状態を検証
    ''' </summary>
    ''' <returns>スタックが有効な状態の場合True</returns>
    Public Function ValidateStackState() As Boolean
        ValidateStackState = (mStack.Count <= MAX_STACK_TRACE_DEPTH)
    End Function
    
    ''' <summary>
    ''' スタックをクリア
    ''' </summary>
    Public Sub ClearStack()
        Set mStack = New Collection
    End Sub
    
    ''' <summary>
    ''' パフォーマンスモニターの参照を取得
    ''' </summary>
    Public Function GetPerformanceMonitor() As clsPerformanceMonitor
        Set GetPerformanceMonitor = mPerformanceMonitor
    End Function
    
    ''' <summary>
    ''' モジュールをリセット
    ''' </summary>
    Public Sub ResetModule()
        Set mStack = Nothing
        Set mPerformanceMonitor = Nothing
        mIsInitialized = False
        
        Set mStack = New Collection
        Set mPerformanceMonitor = New clsPerformanceMonitor
        mIsInitialized = True
    End Sub
#End If