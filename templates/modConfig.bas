Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "modConfig"

' ======================
' 定数定義
' ======================
Private Const CONFIG_FILE_PATH As String = "config.ini"
Private Const MAX_BUFFER_SIZE As Long = 1024
Private Const DEFAULT_SECTION As String = "Settings"

' ======================
' 型定義
' ======================
Private Type ConfigurationSettings
    LogLevel As LogLevel
    LogDestination As LogDestination
    LogFilePath As String
    DatabaseConnectionString As String
    SecurityLevel As SecurityLevel
    PerformanceMonitoringEnabled As Boolean
    DiagnosticsEnabled As Boolean
    EncryptionKey As String
End Type

' ======================
' プライベート変数
' ======================
Private mSettings As ConfigurationSettings
Private mSettingsLock As clsLock
Private mPerformanceMonitor As clsPerformanceMonitor
Private mIsInitialized As Boolean

' ======================
' 初期化・終了処理
' ======================
Public Sub InitializeModule()
    If mIsInitialized Then Exit Sub
    
    Set mSettingsLock = New clsLock
    Set mPerformanceMonitor = New clsPerformanceMonitor
    
    LoadDefaultSettings
    LoadConfigurationFromFile
    
    mIsInitialized = True
    
    ' パフォーマンスモニタリング開始
    If mSettings.PerformanceMonitoringEnabled Then
        mPerformanceMonitor.StartMeasurement "ConfigInitialization"
    End If
End Sub

Public Sub TerminateModule()
    If Not mIsInitialized Then Exit Sub
    
    ' パフォーマンスモニタリング終了
    If mSettings.PerformanceMonitoringEnabled Then
        mPerformanceMonitor.EndMeasurement "ConfigInitialization"
    End If
    
    Set mSettingsLock = Nothing
    Set mPerformanceMonitor = Nothing
    mIsInitialized = False
End Sub

' ======================
' 公開プロパティ
' ======================
Public Property Get Settings() As ConfigurationSettings
    If Not mIsInitialized Then InitializeModule
    
    mSettingsLock.AcquireLock
    Settings = mSettings
    mSettingsLock.ReleaseLock
End Property

Public Property Let Settings(ByVal Value As ConfigurationSettings)
    If Not mIsInitialized Then InitializeModule
    
    mSettingsLock.AcquireLock
    mSettings = Value
    mSettingsLock.ReleaseLock
    
    ' 設定の永続化
    SaveConfigurationToFile
End Property

' ======================
' 公開メソッド
' ======================
Public Function GetConfigValue(ByVal section As String, ByVal key As String, _
                             Optional ByVal defaultValue As String = "") As String
    If Not mIsInitialized Then InitializeModule
    
    Dim buffer As String
    Dim result As Long
    
    buffer = String$(MAX_BUFFER_SIZE, 0)
    result = modWindowsAPI.GetPrivateProfileString(section, key, defaultValue, buffer, Len(buffer), GetConfigFilePath())
    
    If result > 0 Then
        GetConfigValue = Left$(buffer, result)
    Else
        GetConfigValue = defaultValue
    End If
End Function

Public Function SetConfigValue(ByVal section As String, ByVal key As String, _
                             ByVal Value As String) As Boolean
    If Not mIsInitialized Then InitializeModule
    
    SetConfigValue = (modWindowsAPI.WritePrivateProfileString(section, key, Value, GetConfigFilePath()) <> 0)
End Function

' ======================
' プライベートメソッド
' ======================
Private Sub LoadDefaultSettings()
    With mSettings
        .LogLevel = LevelInfo
        .LogDestination = DestFile
        .LogFilePath = DEFAULT_LOG_FILE
        .SecurityLevel = LevelMedium
        .PerformanceMonitoringEnabled = True
        .DiagnosticsEnabled = True
    End With
End Sub

Private Sub LoadConfigurationFromFile()
    On Error GoTo ErrorHandler
    
    With mSettings
        ' ログ設定
        .LogLevel = CInt(GetConfigValue(DEFAULT_SECTION, "LogLevel", CStr(LevelInfo)))
        .LogDestination = CInt(GetConfigValue(DEFAULT_SECTION, "LogDestination", CStr(DestFile)))
        .LogFilePath = GetConfigValue(DEFAULT_SECTION, "LogFilePath", DEFAULT_LOG_FILE)
        
        ' データベース設定
        .DatabaseConnectionString = GetConfigValue("Database", "ConnectionString", "")
        
        ' セキュリティ設定
        .SecurityLevel = CInt(GetConfigValue("Security", "Level", CStr(LevelMedium)))
        .EncryptionKey = GetConfigValue("Security", "EncryptionKey", "")
        
        ' 診断設定
        .PerformanceMonitoringEnabled = CBool(GetConfigValue("Diagnostics", "PerformanceMonitoring", "True"))
        .DiagnosticsEnabled = CBool(GetConfigValue("Diagnostics", "Enabled", "True"))
    End With
    
    Exit Sub

ErrorHandler:
    Dim errDetail As typErrorIDetail
    With errDetail
        .ErrorCode = ERR_FILEIO_INVALID_FORMAT
        .Description = "設定ファイルの読み込み中にエラーが発生しました: " & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "LoadConfigurationFromFile"
        .StackTrace = GetCurrentCallStack
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    Resume Next
End Sub

Private Sub SaveConfigurationToFile()
    On Error GoTo ErrorHandler
    
    With mSettings
        ' ログ設定
        SetConfigValue DEFAULT_SECTION, "LogLevel", CStr(.LogLevel)
        SetConfigValue DEFAULT_SECTION, "LogDestination", CStr(.LogDestination)
        SetConfigValue DEFAULT_SECTION, "LogFilePath", .LogFilePath
        
        ' データベース設定
        SetConfigValue "Database", "ConnectionString", .DatabaseConnectionString
        
        ' セキュリティ設定
        SetConfigValue "Security", "Level", CStr(.SecurityLevel)
        SetConfigValue "Security", "EncryptionKey", .EncryptionKey
        
        ' 診断設定
        SetConfigValue "Diagnostics", "PerformanceMonitoring", CStr(.PerformanceMonitoringEnabled)
        SetConfigValue "Diagnostics", "Enabled", CStr(.DiagnosticsEnabled)
    End With
    
    Exit Sub

ErrorHandler:
    Dim errDetail As typErrorIDetail
    With errDetail
        .ErrorCode = ERR_FILEIO_ACCESS_DENIED
        .Description = "設定ファイルの保存中にエラーが発生しました: " & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "SaveConfigurationToFile"
        .StackTrace = GetCurrentCallStack
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    Resume Next
End Sub

Private Function GetConfigFilePath() As String
    GetConfigFilePath = App.Path & "\" & CONFIG_FILE_PATH
End Function

' ======================
' ヘルパー関数
' ======================
Private Function GetCurrentCallStack() As String
    Dim callStack As New clsCallStack
    
    ' 現在のプロシージャ情報をスタックに追加
    callStack.Push MODULE_NAME, "GetCurrentCallStack"
    
    ' スタックトレースを取得
    GetCurrentCallStack = callStack.StackTrace
End Function

' ======================
' テストサポート機能
' ======================
#If DEBUG Then
    Public Sub ResetModule()
        TerminateModule
        InitializeModule
    End Sub
    
    Public Function ValidateSettings() As Boolean
        With mSettings
            ValidateSettings = _
                .LogLevel >= LevelDebug And .LogLevel <= LevelFatal And _
                .LogDestination >= DestNone And .LogDestination <= DestEmail And _
                .SecurityLevel >= LevelLow And .SecurityLevel <= LevelExtreme
        End With
    End Function
#End If