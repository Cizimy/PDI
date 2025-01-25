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
' メンバ変数
' ======================
Private mPerformanceMonitor As clsPerformanceMonitor
Private mIsInitialized As Boolean

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mPerformanceMonitor = New clsPerformanceMonitor
    modStackTrace.InitializeModule
    mIsInitialized = True
End Sub

Private Sub Class_Terminate()
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
    
    modStackTrace.PushStackEntry ModuleName, ProcedureName
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "CallStack_Push"
    End If
    Exit Sub

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = modErrorCodes.ErrUnexpected
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
    
    Pop = modStackTrace.PopStackEntry()
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "CallStack_Pop"
    End If
    Exit Function

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = modErrorCodes.ErrUnexpected
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
    
    StackTrace = modStackTrace.GetStackTrace()
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "CallStack_GetTrace"
    End If
    Exit Property

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = modErrorCodes.ErrUnexpected
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
    Count = modStackTrace.StackDepth
End Property

' ======================
' テストサポート機能（開発環境専用）
' 警告: これらのメソッドは開発時のテスト目的でのみ使用し、
' 本番環境では使用しないでください。
' ======================
#If DEBUG Then
    ''' <summary>
    ''' スタックの状態を検証（テスト用）
    ''' </summary>
    ''' <returns>スタックが有効な状態の場合True</returns>
    Private Function ValidateStackState() As Boolean
        ValidateStackState = modStackTrace.ValidateStackState()
    End Function
    
    ''' <summary>
    ''' スタックをクリア（テスト用）
    ''' </summary>
    Private Sub ClearStack()
        modStackTrace.ClearStack
    End Sub
    
    ''' <summary>
    ''' パフォーマンスモニターの参照を取得（テスト用）
    ''' </summary>
    Private Function GetPerformanceMonitor() As clsPerformanceMonitor
        Set GetPerformanceMonitor = mPerformanceMonitor
    End Function
    
    ''' <summary>
    ''' モジュールをリセット（テスト用）
    ''' </summary>
    Private Sub ResetModule()
        Set mPerformanceMonitor = Nothing
        mIsInitialized = False
        
        Set mPerformanceMonitor = New clsPerformanceMonitor
        modStackTrace.ResetModule
        mIsInitialized = True
    End Sub
#End If