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

' プライベート変数
Private mErrorManager As clsErrorManager
Private mCriticalSection As clsCriticalSectionLock
Private mMessageResource As ErrorMessageResource

Private Sub Class_Initialize()
    Set mErrorManager = New clsErrorManager
    Set mCriticalSection = New clsCriticalSectionLock
    Set mMessageResource = New ErrorMessageResource
End Sub

Private Sub Class_Terminate()
    Set mErrorManager = Nothing
    Set mCriticalSection = Nothing
    Set mMessageResource = Nothing
End Sub

Private Function IError_HandleError(ByRef errorInfo As ErrorInfo, Optional ByVal options As ErrorHandlingOptions = ehoNone) As Boolean
    mCriticalSection.Enter
    
    On Error GoTo ErrorHandler
    ' エラーメッセージの国際化対応
    errorInfo.Message = mMessageResource.GetMessage(errorInfo.ErrorCode)
    IError_HandleError = mErrorManager.HandleError(errorInfo, options)
    
CleanUp:
    mCriticalSection.Leave
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in SynchronizedErrorManager.HandleError: " & Err.Description
    Resume CleanUp
End Function

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