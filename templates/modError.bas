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
Private mPerformanceMonitor As clsPerformanceMonitor

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
    Set mPerformanceMonitor = New clsPerformanceMonitor
    errorRecursionCount = 0
    RegisterDefaultHandlers
    
    isInitialized = True
End Sub

Public Sub TerminateModule()
    If Not isInitialized Then Exit Sub
    
    Set errorHandlers = Nothing
    Set mLock = Nothing
    Set mPerformanceMonitor = Nothing
    errorRecursionCount = 0
    isInitialized = False
End Sub

' ======================
' エラーハンドリング
' ======================
Private Type ErrorContext
    Info As ErrorInfo
    Handler As IErrorHandler
    IsLocked As Boolean
    IsEmergency As Boolean
End Type

Private Function TryHandleError(ByRef context As ErrorContext) As Boolean
    On Error GoTo ErrorHandler
    
    ' パフォーマンス計測開始
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.StartMeasurement "TryHandleError"
    End If
    
    ' エラー情報の補完
    With context.Info
        If .OccurredAt = #12:00:00 AM# Then .OccurredAt = Now
        If .Category = 0 Then .Category = modErrorCodes.GetErrorCategory(.Code)
        If Len(.StackTrace) = 0 Then .StackTrace = modStackTrace.GetStackTrace()
    End With
    
    ' エラーハンドラの取得
    Set context.Handler = GetErrorHandler(context.Info.Code)
    
    ' パフォーマンス計測終了
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "TryHandleError"
    End If
    
    TryHandleError = True
    Exit Function
    
ErrorHandler:
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "TryHandleError"
    End If
    TryHandleError = False
End Function

Public Sub HandleError(ByRef errInfo As ErrorInfo)
    If Not isInitialized Then InitializeModule
    
    ' パフォーマンス計測開始
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.StartMeasurement "HandleError"
    End If
    
    Dim context As ErrorContext
    context.Info = errInfo
    
    ' エラーの再帰を防ぐ
    errorRecursionCount = errorRecursionCount + 1
    If errorRecursionCount > MAX_ERROR_RECURSION Then
        EmergencyErrorLog "エラー処理の再帰回数が上限を超えました。処理を中断します。"
        context.IsEmergency = True
        GoTo Cleanup
    End If

    ' ロック取得
    On Error Resume Next
    mLock.AcquireLock
    context.IsLocked = (Err.Number = 0)
    On Error GoTo 0

    ' エラー処理のメイン部分
    If TryHandleError(context) Then
        If Not context.Handler Is Nothing Then
            Dim proceed As Boolean
            proceed = context.Handler.HandleError(context.Info)
            
            ' エラー処理の結果に基づいて処理を継続するかどうかを判断
            If Not proceed Then
                context.IsEmergency = True
                GoTo Cleanup
            End If
        End If
    Else
        context.IsEmergency = True
    End If

Cleanup:
    ' クリーンアップ処理
    If context.IsLocked Then
        mLock.ReleaseLock
    End If

    If context.IsEmergency Then
        EmergencyErrorLog "HandleError中にエラーが発生しました: " & Err.Description
    End If
    
    ' パフォーマンス計測終了
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "HandleError"
    End If

    errorRecursionCount = errorRecursionCount - 1
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
' テストサポート機能
' ======================
#If DEBUG Then
    ' === エラー処理テスト ===
    Public Sub TestErrorHandling()
        Dim testError As ErrorInfo
        With testError
            .Code = ErrUnexpected
            .Description = "テスト用エラー"
            .Category = ECGeneral
            .Source = MODULE_NAME
            .ProcedureName = "TestErrorHandling"
            .StackTrace = ""
            .OccurredAt = Now
        End With
        
        mPerformanceMonitor.StartMeasurement "ErrorHandlingTest"
        HandleError testError
        mPerformanceMonitor.EndMeasurement "ErrorHandlingTest"
        
        Debug.Print "テスト実行時間: " & _
                   mPerformanceMonitor.GetMeasurement("ErrorHandlingTest")
    End Sub
    
    ' === 再帰制御テスト ===
    Public Sub TestErrorRecursion()
        Dim i As Long
        For i = 1 To MAX_ERROR_RECURSION + 1
            Dim testError As ErrorInfo
            With testError
                .Code = ErrUnexpected
                .Description = "再帰テスト" & i
                .Category = ECGeneral
                .Source = MODULE_NAME
                .ProcedureName = "TestErrorRecursion"
                .StackTrace = ""
                .OccurredAt = Now
            End With
            
            mPerformanceMonitor.StartMeasurement "RecursionTest_" & i
            HandleError testError
            mPerformanceMonitor.EndMeasurement "RecursionTest_" & i
            
            Debug.Print "再帰テスト" & i & "実行時間: " & _
                       mPerformanceMonitor.GetMeasurement("RecursionTest_" & i)
        Next i
    End Sub
    
    ' === リソース管理テスト ===
    Public Sub TestResourceManagement()
        Dim lockCountBefore As Long
        lockCountBefore = GetActiveLockCount()
        
        Dim testError As ErrorInfo
        With testError
            .Code = ErrUnexpected
            .Description = "リソース管理テスト"
            .Category = ECGeneral
            .Source = MODULE_NAME
            .ProcedureName = "TestResourceManagement"
            .StackTrace = ""
            .OccurredAt = Now
        End With
        
        mPerformanceMonitor.StartMeasurement "ResourceTest"
        
        On Error Resume Next
        HandleError testError
        
        mPerformanceMonitor.EndMeasurement "ResourceTest"
        
        Dim lockCountAfter As Long
        lockCountAfter = GetActiveLockCount()
        
        Debug.Print "リソース管理テスト実行時間: " & _
                   mPerformanceMonitor.GetMeasurement("ResourceTest")
        
        If lockCountBefore <> lockCountAfter Then
            Debug.Print "警告: リソースリークの可能性があります"
            Debug.Print "ロック数 Before: " & lockCountBefore & _
                       ", After: " & lockCountAfter
        End If
    End Sub
    
    ' === パフォーマンスレポート ===
    Public Function GetPerformanceReport() As String
        If Not mPerformanceMonitor Is Nothing Then
            GetPerformanceReport = mPerformanceMonitor.GetAllMeasurements()
        Else
            GetPerformanceReport = "パフォーマンスモニターが初期化されていません。"
        End If
    End Function
    
    ' === 内部状態取得 ===
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
    
    Private Function GetActiveLockCount() As Long
        Dim result As Long
        result = 0
        
        If Not mLock Is Nothing Then
            If mLock.IsLocked Then
                result = result + 1
            End If
        End If
        
        GetActiveLockCount = result
    End Function
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