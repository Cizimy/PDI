Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "modError"

' ======================
' プライベート変数
' ======================
Private mErrorHandlers As Collection
Private mIsInitialized As Boolean

' ======================
' 初期化・終了処理
' ======================
Public Property Get IsInitialized() As Boolean
    IsInitialized = mIsInitialized
End Property

Public Sub InitializeModule()
    If mIsInitialized Then Exit Sub
    
    Set mErrorHandlers = New Collection
    RegisterDefaultHandlers
    
    mIsInitialized = True
End Sub

Public Sub TerminateModule()
    If Not mIsInitialized Then Exit Sub
    
    Set mErrorHandlers = Nothing
    mIsInitialized = False
End Sub

' ======================
' エラーハンドリング
' ======================
Public Sub HandleError(ByRef errInfo As ErrorInfo)
    If Not mIsInitialized Then InitializeModule
    
    ' エラー情報の補完
    With errInfo
        If .OccurredAt = #12:00:00 AM# Then .OccurredAt = Now
        If .Category = 0 Then .Category = modErrorCodes.GetErrorCategory(.Code)
        If Len(.StackTrace) = 0 Then .StackTrace = modStackTrace.GetStackTrace()
    End With
    
    ' エラーハンドラの取得
    Dim handler As IErrorHandler
    Set handler = GetErrorHandler(errInfo.Code)
    
    ' エラーハンドラによる処理
    Dim proceed As Boolean
    proceed = handler.HandleError(errInfo)
    
    ' エラー処理の結果に基づいて処理を継続するかどうかを判断
    If Not proceed Then
        Err.Raise errInfo.Code, errInfo.Source, errInfo.Description
    End If
End Sub

' ======================
' プライベートメソッド
' ======================
Private Sub RegisterDefaultHandlers()
    ' データベース接続エラー用ハンドラ
    RegisterErrorHandler ErrDbConnectionFailed, New DatabaseConnectionErrorHandler
    
    ' ファイル不在エラー用ハンドラ
    RegisterErrorHandler ErrFileNotFound, New FileNotFoundErrorHandler
    
    ' 入力検証エラー用ハンドラ
    RegisterErrorHandler ErrInvalidInput, New InvalidInputErrorHandler
End Sub

Private Function GetErrorHandler(ByVal errorCode As ErrorCode) As IErrorHandler
    On Error Resume Next
    Set GetErrorHandler = mErrorHandlers(CStr(errorCode))
    If Err.Number <> 0 Then
        ' 該当するハンドラが見つからない場合は、エラーカテゴリに基づいてデフォルトハンドラを返す
        Set GetErrorHandler = GetDefaultHandlerForCategory(modErrorCodes.GetErrorCategory(errorCode))
    End If
    On Error GoTo 0
End Function

Private Function GetDefaultHandlerForCategory(ByVal category As ErrorCodeCategory) As IErrorHandler
    Select Case category
        Case ECDatabase
            Set GetDefaultHandlerForCategory = New DatabaseConnectionErrorHandler
        Case ECFileIO
            Set GetDefaultHandlerForCategory = New FileNotFoundErrorHandler
        Case Else
            Set GetDefaultHandlerForCategory = New InvalidInputErrorHandler
    End Select
End Function

' ======================
' パブリックメソッド
' ======================
Public Sub RegisterErrorHandler(ByVal errorCode As ErrorCode, ByVal handler As IErrorHandler)
    If Not mIsInitialized Then InitializeModule
    
    On Error Resume Next
    mErrorHandlers.Remove CStr(errorCode)
    On Error GoTo 0
    
    mErrorHandlers.Add handler, CStr(errorCode)
End Sub

Public Sub UnregisterErrorHandler(ByVal errorCode As ErrorCode)
    If Not mIsInitialized Then Exit Sub
    
    On Error Resume Next
    mErrorHandlers.Remove CStr(errorCode)
    On Error GoTo 0
End Sub

' ======================
' テストサポート機能（開発環境専用）
' 警告: これらのメソッドは開発時のテスト目的でのみ使用し、
' 本番環境では使用しないでください。
' ======================
#If DEBUG Then
    Private Function GetRegisteredHandlerCount() As Long
        GetRegisteredHandlerCount = mErrorHandlers.Count
    End Function
    
    Private Sub ClearHandlers()
        Set mErrorHandlers = New Collection
    End Sub
    
    Private Sub ResetModule()
        TerminateModule
        InitializeModule
    End Sub
#End If