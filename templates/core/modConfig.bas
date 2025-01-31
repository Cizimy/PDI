Option Explicit
Implements IDatabaseConfig

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
    AutoSave As Boolean
End Type

' ======================
' プライベート変数
' ======================
Private settings As ConfigurationSettings
Private settingsLock As clsLock
Private performanceMonitor As clsPerformanceMonitor
Private isInitialized As Boolean
Private isDirty As Boolean

' ======================
' 初期化・終了処理
' ======================
Public Sub InitializeModule()
    If isInitialized Then Exit Sub
    
    Set settingsLock = New clsLock
    Set performanceMonitor = New clsPerformanceMonitor
    
    LoadDefaultSettings
    LoadConfigurationFromFile
    
    isInitialized = True
    
    ' パフォーマンスモニタリング開始
    If settings.PerformanceMonitoringEnabled Then
        performanceMonitor.StartMeasurement "ConfigInitialization"
    End If
End Sub

Public Sub TerminateModule()
    If Not isInitialized Then Exit Sub
    
    ' パフォーマンスモニタリング終了
    If settings.PerformanceMonitoringEnabled Then
        performanceMonitor.EndMeasurement "ConfigInitialization"
    End If
    
    ' 変更された設定を保存
    If isDirty And settings.AutoSave Then
        SaveConfigurationToFile
    End If
    Set settingsLock = Nothing
    Set performanceMonitor = Nothing
    isInitialized = False
End Sub

' ======================
' 公開プロパティ
' ======================
Public Property Get Settings() As ConfigurationSettings
    If Not isInitialized Then InitializeModule
    
    settingsLock.AcquireLock
    Settings = settings
    settingsLock.ReleaseLock
End Property

Public Property Let Settings(ByVal Value As ConfigurationSettings)
    If Not isInitialized Then InitializeModule
    
    settingsLock.AcquireLock
    settings = Value
    settingsLock.ReleaseLock
    
    isDirty = True
    If settings.AutoSave Then
        SaveConfigurationToFile
    End If
End Property

' ======================
' 公開メソッド
' ======================
Public Function GetConfigValue(ByVal section As String, ByVal key As String, _
                             Optional ByVal defaultValue As String = "") As String
    If Not isInitialized Then InitializeModule
    
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
    If Not isInitialized Then InitializeModule
    
    Dim result As Boolean
    result = (modWindowsAPI.WritePrivateProfileString(section, key, Value, GetConfigFilePath()) <> 0)
    
    If result Then
        isDirty = True
    End If
    SetConfigValue = result
End Function

' ======================
' プライベートメソッド
' ======================
Private Sub LoadDefaultSettings()
    With settings
        .LogLevel = LevelInfo
        .LogDestination = DestFile
        .LogFilePath = DEFAULT_LOG_FILE
        .SecurityLevel = LevelMedium
        .PerformanceMonitoringEnabled = True
        .DiagnosticsEnabled = True
        .AutoSave = True
    End With
End Sub

Private Sub LoadConfigurationFromFile()
    On Error GoTo ErrorHandler
    
    With settings
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
        
        ' 自動保存設定
        .AutoSave = CBool(GetConfigValue(DEFAULT_SECTION, "AutoSave", "True"))
    End With
    
    Exit Sub

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrFileInvalidFormat
        .Description = "設定ファイルの読み込み中にエラーが発生しました: " & Err.Description
        .Category = ECFileIO
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
    
    With settings
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
        
        ' 自動保存設定
        SetConfigValue DEFAULT_SECTION, "AutoSave", CStr(.AutoSave)
    End With
    
    isDirty = False
    Exit Sub

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrFileAccessDenied
        .Description = "設定ファイルの保存中にエラーが発生しました: " & Err.Description
        .Category = ECFileIO
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
' 設定管理
' ======================
Public Sub SaveChanges()
    If Not isInitialized Then InitializeModule
    
    If isDirty Then
        SaveConfigurationToFile
    End If
End Sub

Public Property Get HasUnsavedChanges() As Boolean
    HasUnsavedChanges = isDirty
End Property

Public Property Let AutoSave(ByVal Value As Boolean)
    settings.AutoSave = Value
End Property

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
' IDatabaseConfig インターフェースの実装
' ======================
Private Function IDatabaseConfig_GetConnectionString() As String
    If Not isInitialized Then InitializeModule
    
    IDatabaseConfig_GetConnectionString = Me.Settings.DatabaseConnectionString
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
        With settings
            ValidateSettings = _
                .LogLevel >= LevelDebug And .LogLevel <= LevelFatal And _
                .LogDestination >= DestNone And .LogDestination <= DestEmail And _
                .SecurityLevel >= LevelLow And .SecurityLevel <= LevelExtreme
        End With
    End Function
#End If