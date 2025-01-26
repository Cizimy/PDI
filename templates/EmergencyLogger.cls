VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmergencyLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

Implements IEmergencyLogger

'==========================
' 定数
'==========================
Private Const MODULE_NAME As String = "EmergencyLogger"
Private Const EMERGENCY_LOG_FILE As String = "\PDI_emergency.log"
Private Const EVENT_SOURCE As String = "PDI Error"

'==========================
' メンバ変数
'==========================
Private mLock As clsLock

'==========================
' 初期化・終了処理
'==========================
Private Sub Class_Initialize()
    Set mLock = New clsLock
End Sub

Private Sub Class_Terminate()
    Set mLock = Nothing
End Sub

'==========================
' IEmergencyLogger 実装
'==========================
Private Sub IEmergencyLogger_LogEmergencyError(ByVal message As String, Optional ByRef errorInfo As ErrorInfo)
    mLock.AcquireLock
    On Error Resume Next
    
    ' エラー情報をフォーマット
    Dim formattedMessage As String
    formattedMessage = FormatErrorMessage(message, errorInfo)
    
    ' イベントログへの出力を試みる
    WriteToEventLog formattedMessage
    
    ' ファイルへの出力を試みる
    WriteToEmergencyFile formattedMessage
    
    mLock.ReleaseLock
End Sub

'==========================
' プライベートメソッド
'==========================

''' <summary>
''' エラーメッセージをフォーマットします
''' </summary>
Private Function FormatErrorMessage(ByVal message As String, ByRef errorInfo As ErrorInfo) As String
    Dim builder As String
    builder = Now & " - " & message
    
    If Not IsMissing(errorInfo) Then
        With errorInfo
            builder = builder & vbCrLf & _
                     "コード: " & .Code & vbCrLf & _
                     "説明: " & .Description & vbCrLf & _
                     "ソース: " & .Source & vbCrLf & _
                     "プロシージャ: " & .ProcedureName & vbCrLf & _
                     "発生時刻: " & .OccurredAt & vbCrLf & _
                     "スタックトレース: " & vbCrLf & .StackTrace
        End With
    End If
    
    FormatErrorMessage = builder
End Function

''' <summary>
''' イベントログにエラーを記録します
''' </summary>
Private Sub WriteToEventLog(ByVal message As String)
    On Error Resume Next
    modWindowsAPI.WriteToEventLog EVENT_SOURCE, message, EVENTLOG_ERROR_TYPE
End Sub

''' <summary>
''' 緊急用ファイルにエラーを記録します
''' </summary>
Private Sub WriteToEmergencyFile(ByVal message As String)
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim logPath As String
    logPath = Environ$("TEMP") & EMERGENCY_LOG_FILE
    
    Dim logFile As Object
    Set logFile = fso.OpenTextFile(logPath, 8, True) ' 8: ForAppending, Create if doesn't exist
    logFile.WriteLine message
    logFile.Close
    
    Set logFile = Nothing
    Set fso = Nothing
End Sub

'==========================
' テストサポート（開発環境用）
'==========================
#If DEBUG Then
    Public Function GetEmergencyLogPath() As String
        GetEmergencyLogPath = Environ$("TEMP") & EMERGENCY_LOG_FILE
    End Function
    
    Public Sub ClearEmergencyLog()
        On Error Resume Next
        Kill GetEmergencyLogPath()
    End Sub
#End If