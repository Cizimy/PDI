VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SynchronizedErrorManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

Implements IError

' 定数定義
Private Const MAX_RECOVERY_CHAIN_LENGTH As Long = 5
Private Const MAX_RETRY_ATTEMPTS As Long = 3
Private Const RETRY_INTERVAL_MS As Long = 1000
Private Const LOG_BUFFER_SIZE As Long = 1000
Private Const ERROR_LOG_PATH As String = "error.log"

' プライベート変数
Private mErrorManager As clsErrorManager
Private mCriticalSection As clsCriticalSectionLock
Private mMessageResource As ErrorMessageResource
Private mLogger As Object ' FileSystemObject for logging

Private Sub Class_Initialize()
    Set mErrorManager = New clsErrorManager
    Set mCriticalSection = New clsCriticalSectionLock
    Set mMessageResource = New ErrorMessageResource
    Set mLogger = CreateObject("Scripting.FileSystemObject")
    InitializeErrorLog
End Sub

Private Sub InitializeErrorLog()
    On Error Resume Next
    If Not mLogger.FileExists(ERROR_LOG_PATH) Then
        Dim logFile As Object
        Set logFile = mLogger.CreateTextFile(ERROR_LOG_PATH, True)
        logFile.WriteLine "Timestamp,ErrorCode,Message,Source,Details"
        logFile.Close
    End If
End Sub

Private Sub Class_Terminate()
    Set mErrorManager = Nothing
    Set mCriticalSection = Nothing
    Set mMessageResource = Nothing
    Set mLogger = Nothing
    
    ' 最終のログエントリを記録
    On Error Resume Next
    LogError "Class_Terminate", "SynchronizedErrorManager terminated", "Cleanup"
End Sub

Private Function IError_HandleError(ByRef errorInfo As ErrorInfo, Optional ByVal options As ErrorHandlingOptions = ehoNone) As Boolean
    mCriticalSection.Enter
    
    On Error GoTo ErrorHandler
    
    ' エラー情報の検証
    If errorInfo Is Nothing Then
        LogError "HandleError", "ErrorInfo is Nothing", "Validation"
        mCriticalSection.Leave
        Exit Function
    End If
    
    ' エラーメッセージの国際化対応
    errorInfo.Message = mMessageResource.GetMessage(errorInfo.ErrorCode)
    
    ' エラー情報のログ記録
    LogError "HandleError", errorInfo.Message, "Processing", _
            "Code=" & errorInfo.ErrorCode & _
            ";Source=" & errorInfo.SourceModule & _
            ";Severity=" & errorInfo.Severity
    
    ' エラー処理の実行
    IError_HandleError = mErrorManager.HandleError(errorInfo, options)
    
    ' 処理結果のログ記録
    If IError_HandleError Then
        LogError "HandleError", "Error handled successfully", "Success", _
                "Code=" & errorInfo.ErrorCode
    Else
        LogError "HandleError", "Error handling failed", "Failure", _
                "Code=" & errorInfo.ErrorCode
    End If
    
CleanUp:
    mCriticalSection.Leave
    Exit Function
    
ErrorHandler:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description
    
    ' エラーのログ記録
    LogError "HandleError", errDesc, "Error", "ErrNumber=" & errNumber
    
    ' エラーの上位伝播
    Err.Raise errNumber, "SynchronizedErrorManager.HandleError", _
              "Error handling failed: " & errDesc
    Resume CleanUp
End Function

Private Sub LogError(ByVal source As String, ByVal message As String, _
                    ByVal category As String, Optional ByVal details As String = "")
    On Error Resume Next
    
    Dim logEntry As String
    logEntry = Now & "," & _
              source & "," & _
              Replace(message, ",", ";") & "," & _
              category & "," & _
              Replace(details, ",", ";")
    
    Dim logFile As Object
    Set logFile = mLogger.OpenTextFile(ERROR_LOG_PATH, 8, True) ' 8 = ForAppending
    logFile.WriteLine logEntry
    logFile.Close
End Sub

Private Sub IError_RegisterHandler(ByVal errorCode As ErrorCode, ByVal handler As IErrorHandler, Optional ByVal options As HandlerOptions = hoNone)
    mCriticalSection.Enter
    
    On Error GoTo ErrorHandler
    mErrorManager.RegisterHandler errorCode, handler, options
    
CleanUp:
    mCriticalSection.Leave
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in SynchronizedErrorManager.RegisterHandler: " & Err.Description
    Resume CleanUp
End Sub

Private Sub IError_UnregisterHandler(ByVal errorCode As ErrorCode, Optional ByVal options As UnregisterOptions = uoNone)
    mCriticalSection.Enter
    
    On Error GoTo ErrorHandler
    mErrorManager.UnregisterHandler errorCode, options
    
CleanUp:
    mCriticalSection.Leave
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in SynchronizedErrorManager.UnregisterHandler: " & Err.Description
    Resume CleanUp
End Sub

Private Sub IError_IncrementErrorCount(ByVal errorCode As ErrorCode, Optional ByVal options As CountOptions = coNone)
    mCriticalSection.Enter
    
    On Error GoTo ErrorHandler
    mErrorManager.IncrementErrorCount errorCode, options
    
CleanUp:
    mCriticalSection.Leave
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in SynchronizedErrorManager.IncrementErrorCount: " & Err.Description
    Resume CleanUp
End Sub

Private Function IError_CheckErrorThreshold(ByVal errorCode As ErrorCode, ByVal threshold As Long, Optional ByVal options As ThresholdOptions = toNone) As Boolean
    mCriticalSection.Enter
    
    On Error GoTo ErrorHandler
    IError_CheckErrorThreshold = mErrorManager.CheckErrorThreshold(errorCode, threshold, options)
    
CleanUp:
    mCriticalSection.Leave
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in SynchronizedErrorManager.CheckErrorThreshold: " & Err.Description
    Resume CleanUp
End Function

Private Function IError_GetErrorHistory(Optional ByVal options As HistoryOptions = hoNone) As Collection
    mCriticalSection.Enter
    
    On Error GoTo ErrorHandler
    Set IError_GetErrorHistory = mErrorManager.GetErrorHistory(options)
    
CleanUp:
    mCriticalSection.Leave
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in SynchronizedErrorManager.GetErrorHistory: " & Err.Description
    Set IError_GetErrorHistory = New Collection
    Resume CleanUp
End Function

Private Function IError_AnalyzeErrors(Optional ByVal options As AnalysisOptions = aoNone) As ErrorAnalysisResult
    mCriticalSection.Enter
    
    On Error GoTo ErrorHandler
    Set IError_AnalyzeErrors = mErrorManager.AnalyzeErrors(options)
    
CleanUp:
    mCriticalSection.Leave
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in SynchronizedErrorManager.AnalyzeErrors: " & Err.Description
    Set IError_AnalyzeErrors = New ErrorAnalysisResult
    Resume CleanUp
End Function

Private Function IError_CategorizeError(ByRef errorInfo As ErrorInfo, Optional ByVal options As CategoryOptions = coNone) As ErrorCategory
    mCriticalSection.Enter
    
    On Error GoTo ErrorHandler
    IError_CategorizeError = mErrorManager.CategorizeError(errorInfo, options)
    
CleanUp:
    mCriticalSection.Leave
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in SynchronizedErrorManager.CategorizeError: " & Err.Description
    IError_CategorizeError = ecUnknownCategory
    Resume CleanUp
End Function

Private Function IError_GetErrorStatistics(Optional ByVal options As StatisticsOptions = soNone) As ErrorStatistics
    mCriticalSection.Enter
    
    On Error GoTo ErrorHandler
    Set IError_GetErrorStatistics = mErrorManager.GetErrorStatistics(options)
    
CleanUp:
    mCriticalSection.Leave
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in SynchronizedErrorManager.GetErrorStatistics: " & Err.Description
    Set IError_GetErrorStatistics = New ErrorStatistics
    Resume CleanUp
End Function

Private Function IError_GetPerformanceMetrics() As Collection
    mCriticalSection.Enter
    
    On Error GoTo ErrorHandler
    Set IError_GetPerformanceMetrics = mErrorManager.GetPerformanceMetrics()
    
CleanUp:
    mCriticalSection.Leave
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in SynchronizedErrorManager.GetPerformanceMetrics: " & Err.Description
    Set IError_GetPerformanceMetrics = New Collection
    Resume CleanUp
End Function

Private Function IError_Validate() As Boolean
    mCriticalSection.Enter
    
    On Error GoTo ErrorHandler
    IError_Validate = mErrorManager.Validate()
    
CleanUp:
    mCriticalSection.Leave
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in SynchronizedErrorManager.Validate: " & Err.Description
    IError_Validate = False
    Resume CleanUp
End Function

Private Sub IError_Cleanup()
    mCriticalSection.Enter
    
    On Error GoTo ErrorHandler
    mErrorManager.Cleanup
    
CleanUp:
    mCriticalSection.Leave
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in SynchronizedErrorManager.Cleanup: " & Err.Description
    Resume CleanUp
End Sub

'@Description("現在のロケールを設定します")
Public Property Let CurrentLocale(ByVal locale As String)
    mMessageResource.CurrentLocale = locale
End Property

'@Description("現在のロケールを取得します")
Public Property Get CurrentLocale() As String
    CurrentLocale = mMessageResource.CurrentLocale
End Property