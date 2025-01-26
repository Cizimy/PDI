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
Private Const MAX_ERROR_RECURSION As Long = 3
Private errorRecursionCount As Long

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
    errorRecursionCount = 0
    RegisterDefaultHandlers
    
    isInitialized = True
End Sub

Public Sub TerminateModule()
    If Not isInitialized Then Exit Sub
    
    Set errorHandlers = Nothing
    Set mLock = Nothing
    errorRecursionCount = 0
    isInitialized = False
End Sub

' ======================
' エラーハンドリング
' ======================
Public Sub HandleError(ByRef errInfo As ErrorInfo)
    If Not isInitialized Then InitializeModule
    
    ' エラーの再帰を防ぐ
    errorRecursionCount = errorRecursionCount + 1
    If errorRecursionCount > MAX_ERROR_RECURSION Then
        EmergencyErrorLog "エラー処理の再帰回数が上限を超えました。処理を中断します。"
        Exit Sub
    End If
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
    errorRecursionCount = errorRecursionCount - 1
    Exit Sub

ErrorHandler:
    If Not mLock Is Nothing Then mLock.ReleaseLock
    EmergencyErrorLog "HandleError中にエラーが発生しました: " & Err.Description
    Exit Sub
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

' ======================
' エラーログ出力
' ======================
Private Sub EmergencyErrorLog(ByVal message As String)
    On Error Resume Next
    
    ' イベントログへの出力を試みる
    WriteToEventLog message
    
    ' ファイルへの出力を試みる
    WriteToEmergencyFile message
End Sub

Private Sub WriteToEventLog(ByVal message As String)
    ' Windowsイベントログへの出力
    modWindowsAPI.WriteToEventLog "PDI Error", message, EVENTLOG_ERROR_TYPE
End Sub

Private Sub WriteToEmergencyFile(ByVal message As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.OpenTextFile(Environ$("TEMP") & "\PDI_emergency.log", 8, True).WriteLine Now & ": " & message
End Sub