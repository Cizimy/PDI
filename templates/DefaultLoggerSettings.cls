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
' メンバ変数
' ======================
Private mLogLevel As LogLevelEnum
Private mLogDestination As ILogDestination
Private mLogFilePath As String
Private mLogTableName As String
Private mLogEventSource As String
Private mTimerInterval As Long
Private mConnectionString As String

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    ' デフォルト値の設定
    mLogLevel = LOG_LEVEL_INFO
    Set mLogDestination = New FileLogDestination
    mLogFilePath = App.Path & "\" & DEFAULT_LOG_FILE
    mLogTableName = "AppLog"
    mLogEventSource = APPLICATION_NAME
    mTimerInterval = 1000 ' 1秒
    mConnectionString = modConfig.Settings.DatabaseConnectionString
End Sub

Private Sub Class_Terminate()
    Set mLogDestination = Nothing
End Sub

' ======================
' ILoggerSettings実装
' ======================
Private Property Get ILoggerSettings_LogLevel() As LogLevelEnum
    ILoggerSettings_LogLevel = mLogLevel
End Property

Private Property Get ILoggerSettings_LogDestination() As ILogDestination
    Set ILoggerSettings_LogDestination = mLogDestination
End Property

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

Private Property Get ILoggerSettings_ConnectionString() As String
    ILoggerSettings_ConnectionString = mConnectionString
End Property

' ======================
' パブリックプロパティ
' ======================
Public Property Let LogLevel(ByVal Value As LogLevelEnum)
    mLogLevel = Value
End Property

Public Property Set LogDestination(ByVal Value As ILogDestination)
    Set mLogDestination = Value
End Property

Public Property Let LogFilePath(ByVal Value As String)
    mLogFilePath = Value
End Property

Public Property Let LogTableName(ByVal Value As String)
    mLogTableName = Value
End Property

Public Property Let LogEventSource(ByVal Value As String)
    mLogEventSource = Value
End Property

Public Property Let TimerInterval(ByVal Value As Long)
    mTimerInterval = Value
End Property

Public Property Let ConnectionString(ByVal Value As String)
    mConnectionString = Value
End Property

' ======================
' ファクトリメソッド
' ======================
''' <summary>
''' デバッグ用の設定を作成します
''' </summary>
Public Function CreateDebugSettings() As DefaultLoggerSettings
    Set CreateDebugSettings = New DefaultLoggerSettings
    With CreateDebugSettings
        .LogLevel = LOG_LEVEL_DEBUG
        Set .LogDestination = New ConsoleLogDestination
        .TimerInterval = 100 ' より頻繁な更新
    End With
End Function

''' <summary>
''' 本番用の設定を作成します
''' </summary>
Public Function CreateProductionSettings() As DefaultLoggerSettings
    Set CreateProductionSettings = New DefaultLoggerSettings
    With CreateProductionSettings
        .LogLevel = LOG_LEVEL_INFO
        Set .LogDestination = New FileLogDestination
        .TimerInterval = 5000 ' より低頻度な更新
    End With
End Function

''' <summary>
''' データベースログ用の設定を作成します
''' </summary>
Public Function CreateDatabaseSettings() As DefaultLoggerSettings
    Set CreateDatabaseSettings = New DefaultLoggerSettings
    With CreateDatabaseSettings
        .LogLevel = LOG_LEVEL_INFO
        Set .LogDestination = New DatabaseLogDestination
        .LogTableName = "SystemLog"
        .TimerInterval = 2000
    End With
End Function