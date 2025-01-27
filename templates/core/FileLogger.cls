VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILogger

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "FileLogger"

' ======================
' メンバ変数
' ======================
Private mFileOperations As IFileOperations
Private mLock As ILock
Private mPerformanceMonitor As IPerformanceMonitor
Private mLogFilePath As String
Private Const MAX_RETRY_COUNT As Long = 3
Private Const RETRY_DELAY_MS As Long = 1000

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    ' 何もしない（依存性は外部から注入）
End Sub

Private Sub Class_Terminate()
    Set mFileOperations = Nothing
    Set mLock = Nothing
    Set mPerformanceMonitor = Nothing
End Sub

' ======================
' 公開メソッド
' ======================
Public Sub Initialize(ByVal fileOperations As IFileOperations, _
                     ByVal lock As ILock, _
                     ByVal performanceMonitor As IPerformanceMonitor, _
                     ByVal logFilePath As String)
    Set mFileOperations = fileOperations
    Set mLock = lock
    Set mPerformanceMonitor = performanceMonitor
    mLogFilePath = logFilePath
End Sub

' ======================
' ILogger実装
' ======================
Private Sub ILogger_Log(ByVal logLevel As String, _
                       ByVal message As String, _
                       Optional ByVal stackTrace As String, _
                       Optional ByVal errorCode As Long)
    If mFileOperations Is Nothing Then Exit Sub
    
    mPerformanceMonitor.StartMeasurement "FileLogger_Log"
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    Dim logMessage As String
    logMessage = BuildLogMessage(logLevel, message, stackTrace, errorCode)
    
    Dim retryCount As Long
    Do While retryCount < MAX_RETRY_COUNT
        If WriteLogToFile(logMessage) Then Exit Do
        retryCount = retryCount + 1
        If retryCount < MAX_RETRY_COUNT Then Sleep RETRY_DELAY_MS
    Loop
    
Cleanup:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "FileLogger_Log"
    Exit Sub
    
ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "ファイルログ出力中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
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
Private Function BuildLogMessage(ByVal logLevel As String, _
                               ByVal message As String, _
                               ByVal stackTrace As String, _
                               ByVal errorCode As Long) As String
    BuildLogMessage = Format$(Now, "yyyy/mm/dd hh:nn:ss") & " " & _
                     "[" & logLevel & "] " & _
                     IIf(errorCode <> 0, "(Code: " & errorCode & ") ", "") & _
                     message & _
                     IIf(Len(stackTrace) > 0, vbCrLf & "Stack Trace: " & stackTrace, "")
End Function

Private Function WriteLogToFile(ByVal logMessage As String) As Boolean
    On Error Resume Next
    
    ' ファイルの存在確認と作成
    If Not mFileOperations.FileExists(mLogFilePath) Then
        mFileOperations.CreateFile mLogFilePath
    End If
    
    ' ログメッセージの書き込み
    mFileOperations.AppendText mLogFilePath, logMessage & vbCrLf
    
    WriteLogToFile = (Err.Number = 0)
    On Error GoTo 0
End Function