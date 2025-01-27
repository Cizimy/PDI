VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILogger

' ログエントリの構造体
Private Type LogEntry
    Level As String
    Module As String
    Procedure As String
    Message As String
    Timestamp As Date
End Type

' モック用の内部状態
Private mLogEntries() As LogEntry
Private mLogCount As Long
Private mMaxEntries As Long

' カウンター
Private mDebugLogCount As Long
Private mInfoLogCount As Long
Private mWarnLogCount As Long
Private mErrorLogCount As Long

'*******************************************************************************
' 初期化
'*******************************************************************************
Private Sub Class_Initialize()
    mMaxEntries = 1000
    ReDim mLogEntries(1 To mMaxEntries)
    mLogCount = 0
    ResetCounters
End Sub

'*******************************************************************************
' ILogger インターフェースの実装
'*******************************************************************************
Private Sub ILogger_LogDebug(ByVal Module As String, ByVal Procedure As String, ByVal Message As String)
    AddLogEntry "DEBUG", Module, Procedure, Message
    mDebugLogCount = mDebugLogCount + 1
End Sub

Private Sub ILogger_LogInfo(ByVal Module As String, ByVal Procedure As String, ByVal Message As String)
    AddLogEntry "INFO", Module, Procedure, Message
    mInfoLogCount = mInfoLogCount + 1
End Sub

Private Sub ILogger_LogWarning(ByVal Module As String, ByVal Procedure As String, ByVal Message As String)
    AddLogEntry "WARN", Module, Procedure, Message
    mWarnLogCount = mWarnLogCount + 1
End Sub

Private Sub ILogger_LogError(ByVal Module As String, ByVal Procedure As String, ByVal Message As String)
    AddLogEntry "ERROR", Module, Procedure, Message
    mErrorLogCount = mErrorLogCount + 1
End Sub

'*******************************************************************************
' プライベートメソッド
'*******************************************************************************
Private Sub AddLogEntry(ByVal Level As String, ByVal Module As String, ByVal Procedure As String, ByVal Message As String)
    If mLogCount >= mMaxEntries Then Exit Sub
    
    mLogCount = mLogCount + 1
    With mLogEntries(mLogCount)
        .Level = Level
        .Module = Module
        .Procedure = Procedure
        .Message = Message
        .Timestamp = Now
    End With
End Sub

'*******************************************************************************
' モック用のヘルパーメソッド
'*******************************************************************************
Public Property Get LogCount() As Long
    LogCount = mLogCount
End Property

Public Property Get DebugLogCount() As Long
    DebugLogCount = mDebugLogCount
End Property

Public Property Get InfoLogCount() As Long
    InfoLogCount = mInfoLogCount
End Property

Public Property Get WarnLogCount() As Long
    WarnLogCount = mWarnLogCount
End Property

Public Property Get ErrorLogCount() As Long
    ErrorLogCount = mErrorLogCount
End Property

Public Function GetLogEntry(ByVal index As Long) As String
    If index < 1 Or index > mLogCount Then
        GetLogEntry = ""
        Exit Function
    End If
    
    With mLogEntries(index)
        GetLogEntry = .Level & "|" & .Module & "|" & .Procedure & "|" & .Message
    End With
End Function

Public Function ContainsMessage(ByVal searchText As String) As Boolean
    Dim i As Long
    For i = 1 To mLogCount
        If InStr(1, mLogEntries(i).Message, searchText, vbTextCompare) > 0 Then
            ContainsMessage = True
            Exit Function
        End If
    Next i
    ContainsMessage = False
End Function

Public Sub ClearLogs()
    mLogCount = 0
    ResetCounters
End Sub

Private Sub ResetCounters()
    mDebugLogCount = 0
    mInfoLogCount = 0
    mWarnLogCount = 0
    mErrorLogCount = 0
End Sub

'*******************************************************************************
' テスト用のヘルパーメソッド
'*******************************************************************************
Public Function GetLastLogEntry() As String
    If mLogCount = 0 Then
        GetLastLogEntry = ""
    Else
        GetLastLogEntry = GetLogEntry(mLogCount)
    End If
End Function

Public Function GetLogEntriesByLevel(ByVal Level As String) As Long
    Dim count As Long
    Dim i As Long
    For i = 1 To mLogCount
        If mLogEntries(i).Level = Level Then count = count + 1
    Next i
    GetLogEntriesByLevel = count
End Function