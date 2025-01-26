VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventLogLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILogger

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "EventLogLogger"

' ======================
' メンバ変数
' ======================
Private mEventLog As IEventLog
Private mLock As ILock
Private mPerformanceMonitor As IPerformanceMonitor
Private mEventSource As String
Private Const MAX_RETRY_COUNT As Long = 3
Private Const RETRY_DELAY_MS As Long = 1000

' イベントログの種類定数
Private Const EVENTLOG_SUCCESS As Long = 0
Private Const EVENTLOG_ERROR As Long = 1
Private Const EVENTLOG_WARNING As Long = 2
Private Const EVENTLOG_INFORMATION As Long = 4

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    ' 何もしない（依存性は外部から注入）
End Sub

Private Sub Class_Terminate()
    Set mEventLog = Nothing
    Set mLock = Nothing
    Set mPerformanceMonitor = Nothing
End Sub

' ======================
' 公開メソッド
' ======================
Public Sub Initialize(ByVal eventLog As IEventLog, _
                     ByVal lock As ILock, _
                     ByVal performanceMonitor As IPerformanceMonitor, _
                     ByVal eventSource As String)
    Set mEventLog = eventLog
    Set mLock = lock
    Set mPerformanceMonitor = performanceMonitor
    mEventSource = eventSource
End Sub

' ======================
' ILogger実装
' ======================
Private Sub ILogger_Log(ByVal logLevel As String, _
                       ByVal message As String, _
                       Optional ByVal stackTrace As String, _
                       Optional ByVal errorCode As Long)
    If mEventLog Is Nothing Then Exit Sub
    
    mPerformanceMonitor.StartMeasurement "EventLogLogger_Log"
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    Dim retryCount As Long
    Do While retryCount < MAX_RETRY_COUNT
        If WriteLogToEventLog(logLevel, message, stackTrace, errorCode) Then Exit Do
        retryCount = retryCount + 1
        If retryCount < MAX_RETRY_COUNT Then Sleep RETRY_DELAY_MS
    Loop
    
Cleanup:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "EventLogLogger_Log"
    Exit Sub
    
ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "イベントログ出力中にエラーが発生しました: " & Err.Description
        .Category = ECSystem
        .Source = MODULE_NAME
        .ProcedureName = "ILogger_Log"
        .StackTrace = stackTrace
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    Resume Cleanup
End Sub

' ======================
' プライベートメソッド
' ======================
Private Function WriteLogToEventLog(ByVal logLevel As String, _
                                  ByVal message As String, _
                                  ByVal stackTrace As String, _
                                  ByVal errorCode As Long) As Boolean
    On Error Resume Next
    
    ' イベントの種類を決定
    Dim eventType As Long
    Select Case UCase$(logLevel)
        Case "ERROR", "FATAL"
            eventType = EVENTLOG_ERROR
        Case "WARNING"
            eventType = EVENTLOG_WARNING
        Case "SUCCESS"
            eventType = EVENTLOG_SUCCESS
        Case Else
            eventType = EVENTLOG_INFORMATION
    End Select
    
    ' メッセージの組み立て
    Dim logMessage As String
    logMessage = Format$(Now, "yyyy/mm/dd hh:nn:ss") & vbCrLf & _
                 "Level: " & logLevel & vbCrLf & _
                 "Message: " & message
    
    If errorCode <> 0 Then
        logMessage = logMessage & vbCrLf & "Error Code: " & errorCode
    End If
    
    If Len(stackTrace) > 0 Then
        logMessage = logMessage & vbCrLf & "Stack Trace: " & stackTrace
    End If
    
    ' イベントログへの書き込み
    mEventLog.WriteEntry mEventSource, logMessage, eventType
    
    WriteLogToEventLog = (Err.Number = 0)
    On Error GoTo 0
End Function