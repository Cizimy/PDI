VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements IAppConfig
Implements IDatabaseConfig

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "ConfigImpl"

' ======================
' メンバ変数
' ======================
Private mIniFile As IIniFile
Private mSettings As Collection
Private mLock As clsLock
Private mConnectionStringBuilder As IConnectionStringBuilder

' データベース設定
Private mConnectionString As String
Private mConnectionTimeout As Long
Private mCommandTimeout As Long
Private mDatabaseType As String
Private mLogTableName As String

' アプリケーション設定
Private mLogFilePath As String
Private mLogLevel As String
Private mMaxRetryCount As Long
Private mMaxStackDepth As Long
Private mTimerInterval As Long
Private mLogEventSource As String
Private mPerformanceMonitoringEnabled As Boolean
Private mDiagnosticsEnabled As Boolean
Private mAutoSave As Boolean

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mSettings = New Collection
    Set mLock = New clsLock
    Set mIniFile = New IniFileImpl
    
    ' デフォルト値の設定
    mConnectionTimeout = 30
    mCommandTimeout = 30
    mDatabaseType = "ODBC"
    mLogTableName = "AppLog"
    mMaxRetryCount = 3
    mMaxStackDepth = 10
    mTimerInterval = 1000
    mPerformanceMonitoringEnabled = True
    mDiagnosticsEnabled = True
    mAutoSave = True
    
    ' データベースタイプに応じたConnectionStringBuilderの設定
    UpdateConnectionStringBuilder
End Sub

Private Sub Class_Terminate()
    Set mSettings = Nothing
    Set mLock = Nothing
    Set mIniFile = Nothing
    Set mConnectionStringBuilder = Nothing
End Sub

' ======================
' プライベートメソッド
' ======================
Private Sub UpdateConnectionStringBuilder()
    Select Case UCase$(mDatabaseType)
        Case "ODBC"
            Set mConnectionStringBuilder = New ODBCConnectionStringBuilder
        Case "OLEDB"
            Set mConnectionStringBuilder = New OLEDBConnectionStringBuilder
        Case Else
            ' デフォルトはODBCを使用
            Set mConnectionStringBuilder = New ODBCConnectionStringBuilder
    End Select
End Sub

' ======================
' IAppConfig実装
' ======================
Private Function IAppConfig_GetSetting(ByVal settingName As String) As Variant
    mLock.AcquireLock
    On Error Resume Next
    
    IAppConfig_GetSetting = mSettings(settingName)
    
    If Err.Number <> 0 Then
        ' 設定が見つからない場合はNullを返す
        IAppConfig_GetSetting = Null
    End If
    
    mLock.ReleaseLock
End Function

Private Sub IAppConfig_SetSetting(ByVal settingName As String, ByVal settingValue As Variant)
    mLock.AcquireLock
    On Error Resume Next
    
    mSettings.Remove settingName
    mSettings.Add settingValue, settingName
    
    If mAutoSave Then
        IAppConfig_SaveSettings App.Path & "\config.ini"
    End If
    
    mLock.ReleaseLock
End Sub

Private Sub IAppConfig_LoadSettings(ByVal filePath As String)
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    ' INIファイルから設定を読み込む
    mLogFilePath = mIniFile.GetPrivateProfileString("Logging", "LogFilePath", App.Path & "\app.log", filePath)
    mLogLevel = mIniFile.GetPrivateProfileString("Logging", "LogLevel", "INFO", filePath)
    mMaxRetryCount = CLng(mIniFile.GetPrivateProfileString("System", "MaxRetryCount", "3", filePath))
    mMaxStackDepth = CLng(mIniFile.GetPrivateProfileString("System", "MaxStackDepth", "10", filePath))
    mTimerInterval = CLng(mIniFile.GetPrivateProfileString("System", "TimerInterval", "1000", filePath))
    mLogEventSource = mIniFile.GetPrivateProfileString("Logging", "EventSource", "MyApp", filePath)
    mPerformanceMonitoringEnabled = CBool(mIniFile.GetPrivateProfileString("System", "PerformanceMonitoring", "True", filePath))
    mDiagnosticsEnabled = CBool(mIniFile.GetPrivateProfileString("System", "Diagnostics", "True", filePath))
    mAutoSave = CBool(mIniFile.GetPrivateProfileString("System", "AutoSave", "True", filePath))
    
    ' データベース設定の読み込み
    mDatabaseType = mIniFile.GetPrivateProfileString("Database", "Type", "ODBC", filePath)
    mConnectionTimeout = CLng(mIniFile.GetPrivateProfileString("Database", "ConnectionTimeout", "30", filePath))
    mCommandTimeout = CLng(mIniFile.GetPrivateProfileString("Database", "CommandTimeout", "30", filePath))
    mLogTableName = mIniFile.GetPrivateProfileString("Database", "LogTable", "AppLog", filePath)
    
    ' ConnectionStringBuilderの更新
    UpdateConnectionStringBuilder
    
    ' ConnectionStringBuilderのプロパティを設定
    With mConnectionStringBuilder
        .Server = mIniFile.GetPrivateProfileString("Database", "Server", "", filePath)
        .Database = mIniFile.GetPrivateProfileString("Database", "Database", "", filePath)
        .UserName = mIniFile.GetPrivateProfileString("Database", "UserName", "", filePath)
        .Password = mIniFile.GetPrivateProfileString("Database", "Password", "", filePath)
        .AdditionalParameters = mIniFile.GetPrivateProfileString("Database", "AdditionalParameters", "", filePath)
    End With
    
    ' 接続文字列の生成
    mConnectionString = mConnectionStringBuilder.BuildConnectionString()
    
    mLock.ReleaseLock
    Exit Sub
    
ErrorHandler:
    mLock.ReleaseLock
    Err.Raise Err.Number, MODULE_NAME, "設定の読み込み中にエラーが発生しました: " & Err.Description
End Sub

Private Sub IAppConfig_SaveSettings(ByVal filePath As String)
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    ' INIファイルに設定を保存
    mIniFile.WritePrivateProfileString "Logging", "LogFilePath", mLogFilePath, filePath
    mIniFile.WritePrivateProfileString "Logging", "LogLevel", mLogLevel, filePath
    mIniFile.WritePrivateProfileString "System", "MaxRetryCount", CStr(mMaxRetryCount), filePath
    mIniFile.WritePrivateProfileString "System", "MaxStackDepth", CStr(mMaxStackDepth), filePath
    mIniFile.WritePrivateProfileString "System", "TimerInterval", CStr(mTimerInterval), filePath
    mIniFile.WritePrivateProfileString "Logging", "EventSource", mLogEventSource, filePath
    mIniFile.WritePrivateProfileString "System", "PerformanceMonitoring", CStr(mPerformanceMonitoringEnabled), filePath
    mIniFile.WritePrivateProfileString "System", "Diagnostics", CStr(mDiagnosticsEnabled), filePath
    mIniFile.WritePrivateProfileString "System", "AutoSave", CStr(mAutoSave), filePath
    
    ' データベース設定の保存
    mIniFile.WritePrivateProfileString "Database", "Type", mDatabaseType, filePath
    mIniFile.WritePrivateProfileString "Database", "ConnectionTimeout", CStr(mConnectionTimeout), filePath
    mIniFile.WritePrivateProfileString "Database", "CommandTimeout", CStr(mCommandTimeout), filePath
    mIniFile.WritePrivateProfileString "Database", "LogTable", mLogTableName, filePath
    
    ' ConnectionStringBuilder設定の保存
    With mConnectionStringBuilder
        mIniFile.WritePrivateProfileString "Database", "Server", .Server, filePath
        mIniFile.WritePrivateProfileString "Database", "Database", .Database, filePath
        mIniFile.WritePrivateProfileString "Database", "UserName", .UserName, filePath
        mIniFile.WritePrivateProfileString "Database", "Password", .Password, filePath
        mIniFile.WritePrivateProfileString "Database", "AdditionalParameters", .AdditionalParameters, filePath
    End With
    
    mLock.ReleaseLock
    Exit Sub
    
ErrorHandler:
    mLock.ReleaseLock
    Err.Raise Err.Number, MODULE_NAME, "設定の保存中にエラーが発生しました: " & Err.Description
End Sub

' ======================
' IDatabaseConfig実装
' ======================
Private Function IDatabaseConfig_GetConnectionString() As String
    IDatabaseConfig_GetConnectionString = mConnectionString
End Function

Private Function IDatabaseConfig_GetDatabaseSetting(ByVal settingName As String) As Variant
    mLock.AcquireLock
    On Error Resume Next
    
    Select Case UCase$(settingName)
        Case "CONNECTIONTIMEOUT"
            IDatabaseConfig_GetDatabaseSetting = mConnectionTimeout
        Case "COMMANDTIMEOUT"
            IDatabaseConfig_GetDatabaseSetting = mCommandTimeout
        Case "DATABASETYPE"
            IDatabaseConfig_GetDatabaseSetting = mDatabaseType
        Case "LOGTABLENAME"
            IDatabaseConfig_GetDatabaseSetting = mLogTableName
        Case Else
            ' 設定が見つからない場合はNullを返す
            IDatabaseConfig_GetDatabaseSetting = Null
    End Select
    
    mLock.ReleaseLock
End Function

' ======================
' IAppConfigプロパティ実装
' ======================
Private Property Get IAppConfig_LogFilePath() As String
    IAppConfig_LogFilePath = mLogFilePath
End Property

Private Property Let IAppConfig_LogFilePath(ByVal value As String)
    mLogFilePath = value
End Property

Private Property Get IAppConfig_LogLevel() As String
    IAppConfig_LogLevel = mLogLevel
End Property

Private Property Let IAppConfig_LogLevel(ByVal value As String)
    mLogLevel = value
End Property

Private Property Get IAppConfig_MaxRetryCount() As Long
    IAppConfig_MaxRetryCount = mMaxRetryCount
End Property

Private Property Let IAppConfig_MaxRetryCount(ByVal value As Long)
    mMaxRetryCount = value
End Property

Private Property Get IAppConfig_MaxStackDepth() As Long
    IAppConfig_MaxStackDepth = mMaxStackDepth
End Property

Private Property Let IAppConfig_MaxStackDepth(ByVal value As Long)
    mMaxStackDepth = value
End Property

Private Property Get IAppConfig_TimerInterval() As Long
    IAppConfig_TimerInterval = mTimerInterval
End Property

Private Property Let IAppConfig_TimerInterval(ByVal value As Long)
    mTimerInterval = value
End Property

Private Property Get IAppConfig_LogEventSource() As String
    IAppConfig_LogEventSource = mLogEventSource
End Property

Private Property Let IAppConfig_LogEventSource(ByVal value As String)
    mLogEventSource = value
End Property

Private Property Get IAppConfig_PerformanceMonitoringEnabled() As Boolean
    IAppConfig_PerformanceMonitoringEnabled = mPerformanceMonitoringEnabled
End Property

Private Property Let IAppConfig_PerformanceMonitoringEnabled(ByVal value As Boolean)
    mPerformanceMonitoringEnabled = value
End Property

Private Property Get IAppConfig_DiagnosticsEnabled() As Boolean
    IAppConfig_DiagnosticsEnabled = mDiagnosticsEnabled
End Property

Private Property Let IAppConfig_DiagnosticsEnabled(ByVal value As Boolean)
    mDiagnosticsEnabled = value
End Property

Private Property Get IAppConfig_AutoSave() As Boolean
    IAppConfig_AutoSave = mAutoSave
End Property

Private Property Let IAppConfig_AutoSave(ByVal value As Boolean)
    mAutoSave = value
End Property

' ======================
' IDatabaseConfigプロパティ実装
' ======================
Private Property Get IDatabaseConfig_ConnectionString() As String
    IDatabaseConfig_ConnectionString = mConnectionString
End Property

Private Property Let IDatabaseConfig_ConnectionString(ByVal value As String)
    mConnectionString = value
End Property

Private Property Get IDatabaseConfig_ConnectionTimeout() As Long
    IDatabaseConfig_ConnectionTimeout = mConnectionTimeout
End Property

Private Property Let IDatabaseConfig_ConnectionTimeout(ByVal value As Long)
    mConnectionTimeout = value
End Property

Private Property Get IDatabaseConfig_CommandTimeout() As Long
    IDatabaseConfig_CommandTimeout = mCommandTimeout
End Property

Private Property Let IDatabaseConfig_CommandTimeout(ByVal value As Long)
    mCommandTimeout = value
End Property

Private Property Get IDatabaseConfig_DatabaseType() As String
    IDatabaseConfig_DatabaseType = mDatabaseType
End Property

Private Property Let IDatabaseConfig_DatabaseType(ByVal value As String)
    mDatabaseType = value
    UpdateConnectionStringBuilder
End Property

Private Property Get IDatabaseConfig_LogTableName() As String
    IDatabaseConfig_LogTableName = mLogTableName
End Property

Private Property Let IDatabaseConfig_LogTableName(ByVal value As String)
    mLogTableName = value
End Property