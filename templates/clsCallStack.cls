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

'==========================
' モジュール情報
'==========================
Private Const MODULE_NAME As String = "clsCallStack"

'==========================
' メンバ変数
'==========================
Private mStackTrace As IStackTrace
Private mLock As clsLock

'==========================
' 初期化・終了処理
'==========================
Private Sub Class_Initialize()
    Set mStackTrace = New StackTraceImpl
    Set mLock = New clsLock
End Sub

Private Sub Class_Terminate()
    Set mStackTrace = Nothing
    Set mLock = Nothing
End Sub

'==========================
' パブリックメソッド
'==========================

''' <summary>
''' 呼び出し履歴にエントリを追加
''' </summary>
''' <param name="ModuleName">モジュール名</param>
''' <param name="ProcedureName">プロシージャ名</param>
''' <remarks>
''' スタックトレースの最大深さを超えた場合、エラーを発生させずに無視します
''' </remarks>
Public Sub Push(ByVal ModuleName As String, ByVal ProcedureName As String)
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    mStackTrace.PushStackEntry ModuleName, ProcedureName
    
    mLock.ReleaseLock
    Exit Sub

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = modErrorCodes.ErrUnexpected
        .Description = "Push操作中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "Push"
        .StackTrace = mStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    mLock.ReleaseLock
    Resume Next
End Sub

''' <summary>
''' 呼び出し履歴から最後のエントリを取得して削除
''' </summary>
''' <returns>最後に追加されたプロシージャの完全修飾名、またはスタックが空の場合は空文字列</returns>
Public Function Pop() As String
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    Pop = mStackTrace.PopStackEntry()
    
    mLock.ReleaseLock
    Exit Function

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = modErrorCodes.ErrUnexpected
        .Description = "Pop操作中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "Pop"
        .StackTrace = mStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    mLock.ReleaseLock
    Resume Next
End Function

''' <summary>
''' 現在のスタックトレースを文字列として取得
''' </summary>
''' <returns>スタックトレースの文字列表現</returns>
Public Property Get StackTrace() As String
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    StackTrace = mStackTrace.GetStackTrace()
    
    mLock.ReleaseLock
    Exit Property

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = modErrorCodes.ErrUnexpected
        .Description = "スタックトレース取得中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "StackTrace"
        .StackTrace = mStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    mLock.ReleaseLock
    Resume Next
End Property

''' <summary>
''' スタック内のエントリ数を取得
''' </summary>
''' <returns>現在のスタックの深さ</returns>
Public Property Get Count() As Long
    Count = mStackTrace.StackDepth
End Property

'==========================
' テストサポート（開発環境用）
'==========================
#If DEBUG Then
    Public Sub SetStackTrace(ByVal stackTrace As IStackTrace)
        Set mStackTrace = stackTrace
    End Sub
    
    Public Function GetStackTrace() As IStackTrace
        Set GetStackTrace = mStackTrace
    End Function
#End If