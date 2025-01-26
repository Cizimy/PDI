VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

Implements IError

'==========================
' 定数
'==========================
Private Const MODULE_NAME As String = "ErrorImpl"
Private Const MAX_ERROR_RECURSION As Long = 3
Private Const ERROR_COUNT_DICT_SIZE As Long = 100

'==========================
' メンバ変数
'==========================
Private mHandlerManager As ErrorHandlerManager
Private mEmergencyLogger As IEmergencyLogger
Private mErrorCounts As Object ' Scripting.Dictionary
Private mLock As clsLock
Private mIsInitialized As Boolean
Private mErrorRecursionCount As Long
Private mPerformanceMonitor As clsPerformanceMonitor

'==========================
' 初期化・終了処理
'==========================
Private Sub Class_Initialize()
    Set mHandlerManager = New ErrorHandlerManager
    Set mEmergencyLogger = New EmergencyLogger
    Set mErrorCounts = CreateObject("Scripting.Dictionary")
    Set mLock = New clsLock
    Set mPerformanceMonitor = New clsPerformanceMonitor
    mErrorRecursionCount = 0
    mIsInitialized = True
End Sub

Private Sub Class_Terminate()
    Set mHandlerManager = Nothing
    Set mEmergencyLogger = Nothing
    Set mErrorCounts = Nothing
    Set mLock = Nothing
    Set mPerformanceMonitor = Nothing
    mIsInitialized = False
End Sub

'==========================
' IError 実装
'==========================
Private Function IError_HandleError(ByRef errorInfo As ErrorInfo) As Boolean
    If Not mIsInitialized Then Exit Function
    
    ' パフォーマンス計測開始
    mPerformanceMonitor.StartMeasurement "HandleError"
    
    ' エラーの再帰制御
    mErrorRecursionCount = mErrorRecursionCount + 1
    If mErrorRecursionCount > MAX_ERROR_RECURSION Then
        mEmergencyLogger.LogEmergencyError "エラー処理の再帰回数が上限を超えました。処理を中断します。", errorInfo
        IError_HandleError = False
        GoTo Cleanup
    End If
    
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    ' エラー情報の補完
    CompleteErrorInfo errorInfo
    
    ' エラー発生回数をカウント
    IError_IncrementErrorCount errorInfo.Code
    
    ' エラーハンドラーの取得と実行
    Dim handler As IErrorHandler
    Set handler = mHandlerManager.GetErrorHandler(errorInfo.Code)
    
    If Not handler Is Nothing Then
        IError_HandleError = handler.HandleError(errorInfo)
    End If
    
    mLock.ReleaseLock
    GoTo Cleanup
    
ErrorHandler:
    mLock.ReleaseLock
    mEmergencyLogger.LogEmergencyError "HandleError中にエラーが発生しました: " & Err.Description, errorInfo
    IError_HandleError = False
    
Cleanup:
    mErrorRecursionCount = mErrorRecursionCount - 1
    mPerformanceMonitor.EndMeasurement "HandleError"
End Function

Private Sub IError_RegisterHandler(ByVal errorCode As ErrorCode, ByVal handler As IErrorHandler)
    If Not mIsInitialized Then Exit Sub
    mHandlerManager.RegisterHandler errorCode, handler
End Sub

Private Sub IError_UnregisterHandler(ByVal errorCode As ErrorCode)
    If Not mIsInitialized Then Exit Sub
    mHandlerManager.UnregisterHandler errorCode
End Sub

Private Sub IError_IncrementErrorCount(ByVal errorCode As ErrorCode)
    mLock.AcquireLock
    
    Dim key As String
    key = CStr(errorCode)
    
    If Not mErrorCounts.Exists(key) Then
        mErrorCounts.Add key, 1
    Else
        mErrorCounts(key) = mErrorCounts(key) + 1
    End If
    
    ' エラーカウントの上限を超えた場合、古いエントリを削除
    If mErrorCounts.Count > ERROR_COUNT_DICT_SIZE Then
        RemoveOldestErrorCount
    End If
    
    mLock.ReleaseLock
End Sub

Private Function IError_CheckErrorThreshold(ByVal errorCode As ErrorCode, ByVal threshold As Long) As Boolean
    mLock.AcquireLock
    
    Dim key As String
    key = CStr(errorCode)
    
    IError_CheckErrorThreshold = False
    If mErrorCounts.Exists(key) Then
        IError_CheckErrorThreshold = (mErrorCounts(key) >= threshold)
    End If
    
    mLock.ReleaseLock
End Function

'==========================
' プライベートメソッド
'==========================

''' <summary>
''' エラー情報を補完します
''' </summary>
Private Sub CompleteErrorInfo(ByRef errorInfo As ErrorInfo)
    With errorInfo
        If .OccurredAt = #12:00:00 AM# Then .OccurredAt = Now
        If .Category = 0 Then .Category = modErrorCodes.GetErrorCategory(.Code)
        If Len(.StackTrace) = 0 Then .StackTrace = modStackTrace.GetStackTrace()
    End With
End Sub

''' <summary>
''' 最も古いエラーカウントを削除します
''' </summary>
Private Sub RemoveOldestErrorCount()
    If mErrorCounts.Count = 0 Then Exit Sub
    
    Dim oldestKey As String
    oldestKey = mErrorCounts.Keys()(0)
    mErrorCounts.Remove oldestKey
End Sub

'==========================
' テストサポート（開発環境用）
'==========================
#If DEBUG Then
    Public Function GetErrorCount(ByVal errorCode As ErrorCode) As Long
        mLock.AcquireLock
        
        Dim key As String
        key = CStr(errorCode)
        
        GetErrorCount = 0
        If mErrorCounts.Exists(key) Then
            GetErrorCount = mErrorCounts(key)
        End If
        
        mLock.ReleaseLock
    End Function
    
    Public Sub ClearErrorCounts()
        mLock.AcquireLock
        mErrorCounts.RemoveAll
        mLock.ReleaseLock
    End Sub
    
    Public Function GetPerformanceReport() As String
        GetPerformanceReport = mPerformanceMonitor.GetAllMeasurements()
    End Function
#End If