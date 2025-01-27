VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DefaultLoggerSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILoggerSettings

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "DefaultLoggerSettings"

' ======================
' メンバ変数
' ======================
Private mAppConfig As IAppConfig
Private mLogLevel As LogLevel
Private mLogDestinations As Collection
Private mLogFilePath As String
Private mLogTableName As String
Private mLogEventSource As String
Private mTimerInterval As Long
Private mFileOperations As IFileOperations
Private mConnectionPool As IConnectionPool
Private mEventLog As IEventLog
Private mDatabaseConfig As IDatabaseConfig

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mLogDestinations = New Collection
    mTimerInterval = 1000 ' デフォルト値: 1秒
End Sub

Private Sub Class_Terminate()
    Set mAppConfig = Nothing
    Set mLogDestinations = Nothing
    Set mFileOperations = Nothing
    Set mConnectionPool = Nothing
    Set mEventLog = Nothing
    Set mDatabaseConfig = Nothing
End Sub

' ======================
' 公開メソッド
' ======================
Public Sub Initialize(ByVal appConfig As IAppConfig, _
                     ByVal fileOperations As IFileOperations, _
                     ByVal connectionPool As IConnectionPool, _
                     ByVal eventLog As IEventLog, _
                     ByVal databaseConfig As IDatabaseConfig)
    Set mAppConfig = appConfig
    Set mFileOperations = fileOperations
    Set mConnectionPool = connectionPool
    Set mEventLog = eventLog
    Set mDatabaseConfig = databaseConfig
    
    ' 設定の読み込み
    LoadSettings
End Sub

' ======================
' ILoggerSettings実装
' ======================
Private Property Get ILoggerSettings_LogLevel() As LogLevel
    ILoggerSettings_LogLevel = mLogLevel
End Property

Private Property Let ILoggerSettings_LogLevel(ByVal value As LogLevel)
    mLogLevel = value
End Property

Private Function ILoggerSettings_GetLogDestinations() As Variant
    Dim result() As String
    ReDim result(1 To mLogDestinations.Count)
    
    Dim i As Long
    For i = 1 To mLogDestinations.Count
        result(i) = mLogDestinations(i)
    Next i
    
    ILoggerSettings_GetLogDestinations = result
End Function

Private Property Get ILoggerSettings_LogFilePath() As String
    ILoggerSettings_LogFilePath = mLogFilePath
End Property

Private Property Get ILoggerSettings_LogTableName() As String
    ILoggerSettings_LogTableName = mLogTableName
End Property

Private Property Get ILoggerSettings_LogEventSource() As String
    ILoggerSettings_LogEventSource = mLogEventSource
End Property

Private Property Get ILoggerSettings_TimerInterval() As Long
    ILoggerSettings_TimerInterval = mTimerInterval
End Property

Private Function ILoggerSettings_GetFileOperations() As IFileOperations
    Set ILoggerSettings_GetFileOperations = mFileOperations
End Function

Private Function ILoggerSettings_GetConnectionPool() As IConnectionPool
    Set ILoggerSettings_GetConnectionPool = mConnectionPool
End Function

Private Function ILoggerSettings_GetEventLog() As IEventLog
    Set ILoggerSettings_GetEventLog = mEventLog
End Function

Private Function ILoggerSettings_GetDatabaseConfig() As IDatabaseConfig
    Set ILoggerSettings_GetDatabaseConfig = mDatabaseConfig
End Function

Private Function ILoggerSettings_ShouldLog(ByVal destination As String, ByVal level As LogLevel) As Boolean
    ' 指定された出力先が有効で、ログレベルが現在の設定以上の場合にログを出力
    Dim i As Long
    For i = 1 To mLogDestinations.Count
        If mLogDestinations(i) = destination Then
            ILoggerSettings_ShouldLog = (level >= mLogLevel)
            Exit Function
        End If
    Next i
    
    ILoggerSettings_ShouldLog = False
End Function

Private Function ILoggerSettings_FormatLogMessage(ByVal logLevel As String, _
                                                ByVal message As String, _
                                                ByVal stackTrace As String, _
                                                ByVal errorCode As Long) As String
    Dim result As String
    result = Format$(Now, "yyyy/mm/dd hh:nn:ss") & " " & _
             "[" & logLevel & "] " & _
             IIf(errorCode <> 0, "(Code: " & errorCode & ") ", "") & _
             message
             
    If Len(stackTrace) > 0 Then
        result = result & vbCrLf & "Stack Trace: " & stackTrace
    End If
    
    ILoggerSettings_FormatLogMessage = result
End Function

' ======================
' プライベートメソッド
' ======================
Private Sub LoadSettings()
    If mAppConfig Is Nothing Then Exit Sub
    
    ' ログレベルの読み込み
    mLogLevel = mAppConfig.GetValue("Logging", "Level", LOG_LEVEL_INFO)
    
    ' 出力先の読み込み
    Dim destinations As String
    destinations = mAppConfig.GetValue("Logging", "Destinations", "File")
    ParseDestinations destinations
    
    ' ファイル出力設定
    mLogFilePath = mAppConfig.GetValue("Logging", "FilePath", "log\application.log")
    
    ' データベース出力設定
    mLogTableName = mAppConfig.GetValue("Logging", "TableName", "ApplicationLog")
    
    ' イベントログ出力設定
    mLogEventSource = mAppConfig.GetValue("Logging", "EventSource", "PDI Logger")
    
    ' タイマー間隔設定
    mTimerInterval = CLng(mAppConfig.GetValue("Logging", "TimerInterval", "1000"))
End Sub

Private Sub ParseDestinations(ByVal destinations As String)
    Set mLogDestinations = New Collection
    
    Dim parts() As String
    parts = Split(destinations, ",")
    
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        mLogDestinations.Add Trim$(parts(i))
    Next i
End Sub