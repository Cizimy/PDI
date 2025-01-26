VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConsoleLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILogger

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "ConsoleLogger"

' ======================
' メンバ変数
' ======================
Private mPerformanceMonitor As IPerformanceMonitor
Private mUseColors As Boolean

' ANSIエスケープシーケンス用の色コード
Private Const ANSI_RESET As String = vbNullString
Private Const ANSI_RED As String = vbNullString
Private Const ANSI_YELLOW As String = vbNullString
Private Const ANSI_GREEN As String = vbNullString
Private Const ANSI_BLUE As String = vbNullString

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    ' 何もしない（依存性は外部から注入）
End Sub

Private Sub Class_Terminate()
    Set mPerformanceMonitor = Nothing
End Sub

' ======================
' 公開メソッド
' ======================
Public Sub Initialize(ByVal performanceMonitor As IPerformanceMonitor, _
                     Optional ByVal useColors As Boolean = False)
    Set mPerformanceMonitor = performanceMonitor
    mUseColors = useColors
End Sub

' ======================
' ILogger実装
' ======================
Private Sub ILogger_Log(ByVal logLevel As String, _
                       ByVal message As String, _
                       Optional ByVal stackTrace As String, _
                       Optional ByVal errorCode As Long)
    If mPerformanceMonitor Is Nothing Then Exit Sub
    
    mPerformanceMonitor.StartMeasurement "ConsoleLogger_Log"
    
    On Error GoTo ErrorHandler
    
    ' ログメッセージの組み立てと出力
    Dim logMessage As String
    logMessage = BuildLogMessage(logLevel, message, stackTrace, errorCode)
    WriteToConsole logMessage, logLevel
    
Cleanup:
    mPerformanceMonitor.EndMeasurement "ConsoleLogger_Log"
    Exit Sub
    
ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "コンソールログ出力中にエラーが発生しました: " & Err.Description
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

Private Sub WriteToConsole(ByVal logMessage As String, ByVal logLevel As String)
    If Not mUseColors Then
        Debug.Print logMessage
        Exit Sub
    End If
    
    ' ログレベルに応じた色付け
    Dim colorCode As String
    Select Case UCase$(logLevel)
        Case "ERROR", "FATAL"
            colorCode = ANSI_RED
        Case "WARNING"
            colorCode = ANSI_YELLOW
        Case "SUCCESS"
            colorCode = ANSI_GREEN
        Case "INFO"
            colorCode = ANSI_BLUE
        Case Else
            colorCode = ANSI_RESET
    End Select
    
    Debug.Print colorCode & logMessage & ANSI_RESET
End Sub