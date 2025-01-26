Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "modError"

' ======================
' プライベート変数
' ======================
Private errorHandlers As Collection
Private isInitialized As Boolean
Private mLock As clsLock

' ======================
' 初期化・終了処理
' ======================
Public Property Get IsInitialized() As Boolean
    IsInitialized = isInitialized
End Property

Public Sub InitializeModule()
    If isInitialized Then Exit Sub
    
    Set errorHandlers = New Collection
    Set mLock = New clsLock
    RegisterDefaultHandlers
    
    isInitialized = True
End Sub

Public Sub TerminateModule()
    If Not isInitialized Then Exit Sub
    
    Set errorHandlers = Nothing
    Set mLock = Nothing
    isInitialized = False
End Sub

' ======================
' エラーハンドリング
' ======================
Public Sub HandleError(ByRef errInfo As ErrorInfo)
    If Not isInitialized Then InitializeModule
    
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    ' エラー情報の補完
    With errInfo
        If .OccurredAt = #12:00:00 AM# Then .OccurredAt = Now
        If .Category = 0 Then .Category = modErrorCodes.GetErrorCategory(.Code)
        If Len(.StackTrace) = 0 Then .StackTrace = modStackTrace.GetStackTrace()
    End With
    
    ' エラーハンドラの取得
    Dim handler As IErrorHandler
    Set handler = GetErrorHandler(errInfo.Code)
    
    mLock.ReleaseLock
    
    ' エラーハンドラによる処理
    Dim proceed As Boolean
    proceed = handler.HandleError(errInfo)
    
    ' エラー処理の結果に基づいて処理を継続するかどうかを判断
    If Not proceed Then
        Err.Raise errInfo.Code, errInfo.Source, errInfo.Description
    End If
    Exit Sub

ErrorHandler:
    If Not mLock Is Nothing Then mLock.ReleaseLock
    Err.Raise Err.Number, Err.Source, "HandleError中にエラーが発生しました: " & Err.Description
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
    Dim handler As IErrorHandler
    
    On Error Resume Next
    Set handler = errorHandlers(CStr(errorCode))
    If Err.Number <> 0 Then
        ' 該当するハンドラが見つからない場合は、エラーカテゴリに基づいてデフォルトハンドラを返す
        Set handler = GetDefaultHandlerForCategory(modErrorCodes.GetErrorCategory(errorCode))
    End If
    On Error GoTo 0
    
    Set GetErrorHandler = handler
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
    If Not isInitialized Then InitializeModule
    
    mLock.AcquireLock
    On Error Resume Next
    errorHandlers.Remove CStr(errorCode)
    On Error GoTo 0
    
    errorHandlers.Add handler, CStr(errorCode)
    mLock.ReleaseLock
End Sub

Public Sub UnregisterErrorHandler(ByVal errorCode As ErrorCode)
    If Not isInitialized Then Exit Sub
    
    mLock.AcquireLock
    On Error Resume Next
    errorHandlers.Remove CStr(errorCode)
    On Error GoTo 0
    mLock.ReleaseLock
End Sub

' ======================
' テストサポート機能（開発環境専用）
' 警告: これらのメソッドは開発時のテスト目的でのみ使用し、
' 本番環境では使用しないでください。
' ======================
#If DEBUG Then
    Private Function GetRegisteredHandlerCount() As Long
        mLock.AcquireLock
        GetRegisteredHandlerCount = errorHandlers.Count
        mLock.ReleaseLock
    End Function
    
    Private Sub ClearHandlers()
        mLock.AcquireLock
        Set errorHandlers = New Collection
        mLock.ReleaseLock
    End Sub
    
    Private Sub ResetModule()
        TerminateModule
        InitializeModule
    End Sub
#End If