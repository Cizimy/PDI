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
' 依存インターフェース
' ======================
Private mIniFile As IIniFile
Private mCrypto As ICryptography
Private mPerformanceMonitor As IPerformanceMonitor
Private mFileOperations As IFileOperations

' ======================
' メンバ変数
' ======================
Private mLock As clsLock
Private mSettings As Collection
Private mSettingHistory As Collection
Private mConnectionStringBuilder As IConnectionStringBuilder
Private mValidationRules As Collection
Private mPerformanceMetrics As Collection
Private mEncryptedSettings As Collection

' データベース設定
Private mConnectionString As String
Private mConnectionTimeout As Long
Private mCommandTimeout As Long
Private mDatabaseType As String
Private mLogTableName As String

' アプリケーション設定
Private mLogFilePath As String
Private mLogLevel As LogLevel
Private mMaxRetryCount As Long
Private mMaxStackDepth As Long
Private mTimerInterval As Long
Private mLogEventSource As String
Private mPerformanceMonitoringEnabled As Boolean
Private mDiagnosticsEnabled As Boolean
Private mAutoSave As Boolean

' 暗号化設定
Private mEncryptionKey As String
Private mEncryptionKeyVersion As String
Private mCurrentEnvironment As String

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    ' コレクションの初期化
    Set mSettings = New Collection
    Set mSettingHistory = New Collection
    Set mValidationRules = New Collection
    Set mPerformanceMetrics = New Collection
    Set mEncryptedSettings = New Collection
    
    ' 依存オブジェクトの初期化
    Set mLock = New clsLock
    Set mIniFile = New IniFileImpl
    Set mCrypto = New CryptographyImpl
    Set mPerformanceMonitor = New PerformanceMonitorImpl
    Set mFileOperations = New FileSystemOperations
    
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
    mLogLevel = llInfo
    mCurrentEnvironment = "development"
    
    ' データベースタイプに応じたConnectionStringBuilderの設定
    UpdateConnectionStringBuilder
End Sub

Private Sub Class_Terminate()
    Set mSettings = Nothing
    Set mSettingHistory = Nothing
    Set mValidationRules = Nothing
    Set mPerformanceMetrics = Nothing
    Set mEncryptedSettings = Nothing
    Set mLock = Nothing
    Set mIniFile = Nothing
    Set mCrypto = Nothing
    Set mPerformanceMonitor = Nothing
    Set mFileOperations = Nothing
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
            Set mConnectionStringBuilder = New ODBCConnectionStringBuilder
    End Select
End Sub

Private Sub AddSettingHistory(ByVal settingName As String, ByVal oldValue As Variant, ByVal newValue As Variant)
    Dim historyEntry As New Dictionary
    historyEntry.Add "SettingName", settingName
    historyEntry.Add "OldValue", oldValue
    historyEntry.Add "NewValue", newValue
    historyEntry.Add "Timestamp", Now
    historyEntry.Add "Environment", mCurrentEnvironment
    mSettingHistory.Add historyEntry
End Sub

Private Function IsEncrypted(ByVal settingName As String) As Boolean
    On Error Resume Next
    mEncryptedSettings.Item settingName
    IsEncrypted = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function EncryptValue(ByVal value As String) As String
    If mEncryptionKey = "" Then
        Err.Raise vbObjectError + 1, MODULE_NAME, "暗号化キーが設定されていません。"
    End If
    EncryptValue = mCrypto.Encrypt(value, mEncryptionKey)
End Function

Private Function DecryptValue(ByVal value As String) As String
    If mEncryptionKey = "" Then
        Err.Raise vbObjectError + 1, MODULE_NAME, "暗号化キーが設定されていません。"
    End If
    DecryptValue = mCrypto.Decrypt(value, mEncryptionKey)
End Function

' ======================
' IAppConfig実装
' ======================
Private Function IAppConfig_GetSetting(ByVal settingName As String, Optional ByVal options As SettingOptions = soNone) As Variant
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    Dim result As Variant
    result = mSettings(settingName)
    
    ' 暗号化された設定値の復号化
    If Not IsNull(result) And IsEncrypted(settingName) And (options And soEncrypt) = soEncrypt Then
        result = DecryptValue(CStr(result))
    End If
    
    ' パフォーマンスメトリクスの記録
    If mPerformanceMonitoringEnabled Then
        Dim metric As New Dictionary
        metric.Add "Type", mtResponseTime
        metric.Add "Operation", "GetSetting"
        metric.Add "SettingName", settingName
        metric.Add "Timestamp", Now
        mPerformanceMetrics.Add metric
    End If
    
    mLock.ReleaseLock
    IAppConfig_GetSetting = result
    Exit Function
    
ErrorHandler:
    mLock.ReleaseLock
    If Err.Number = 5 Then ' Item not found in collection
        IAppConfig_GetSetting = Null
    Else
        Err.Raise Err.Number, MODULE_NAME, "設定の取得中にエラーが発生しました: " & Err.Description
    End If
End Function

Private Sub IAppConfig_SetSetting(ByVal settingName As String, ByVal settingValue As Variant, Optional ByVal options As SettingOptions = soNone)
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    ' 現在の値を履歴用に保存
    Dim oldValue As Variant
    oldValue = IAppConfig_GetSetting(settingName)
    
    ' 値の暗号化
    Dim valueToStore As Variant
    valueToStore = settingValue
    If (options And soEncrypt) = soEncrypt Then
        valueToStore = EncryptValue(CStr(settingValue))
        mEncryptedSettings.Add settingName, settingName
    End If
    
    ' 値の設定
    mSettings.Remove settingName
    mSettings.Add valueToStore, settingName
    
    ' 履歴の記録
    If (options And soLogHistory) = soLogHistory Then
        AddSettingHistory settingName, oldValue, settingValue
    End If
    
    ' 自動保存
    If mAutoSave Then
        IAppConfig_SaveSettings App.Path & "\config.ini"
    End If
    
    mLock.ReleaseLock
    Exit Sub
    
ErrorHandler:
    mLock.ReleaseLock
    Err.Raise Err.Number, MODULE_NAME, "設定の更新中にエラーが発生しました: " & Err.Description
End Sub

Private Sub IAppConfig_LoadSettings(ByVal filePath As String, Optional ByVal options As LoadOptions = loNone)
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    ' 基本設定の読み込み
    mLogFilePath = mIniFile.GetPrivateProfileString("Logging", "LogFilePath", App.Path & "\app.log", filePath)
    mLogLevel = CLng(mIniFile.GetPrivateProfileString("Logging", "LogLevel", CStr(llInfo), filePath))
    mMaxRetryCount = CLng(mIniFile.GetPrivateProfileString("System", "MaxRetryCount", "3", filePath))
    mMaxStackDepth = CLng(mIniFile.GetPrivateProfileString("System", "MaxStackDepth", "10", filePath))
    mTimerInterval = CLng(mIniFile.GetPrivateProfileString("System", "TimerInterval", "1000", filePath))
    mLogEventSource = mIniFile.GetPrivateProfileString("Logging", "EventSource", "MyApp", filePath)
    mPerformanceMonitoringEnabled = CBool(mIniFile.GetPrivateProfileString("System", "PerformanceMonitoring", "True", filePath))
    mDiagnosticsEnabled = CBool(mIniFile.GetPrivateProfileString("System", "Diagnostics", "True", filePath))
    mAutoSave = CBool(mIniFile.GetPrivateProfileString("System", "AutoSave", "True", filePath))
    
    ' 環境固有の設定の適用
    If (options And loApplyEnvironment) = loApplyEnvironment Then
        Dim envSection As String
        envSection = "Environment_" & mCurrentEnvironment
        
        ' 環境固有の値で上書き
        If mIniFile.GetPrivateProfileString(envSection, "LogFilePath", "", filePath) <> "" Then
            mLogFilePath = mIniFile.GetPrivateProfileString(envSection, "LogFilePath", mLogFilePath, filePath)
        End If
        ' ... 他の環境固有設定も同様に処理
    End If
    
    ' データベース設定の読み込み
    mDatabaseType = mIniFile.GetPrivateProfileString("Database", "Type", "ODBC", filePath)
    mConnectionTimeout = CLng(mIniFile.GetPrivateProfileString("Database", "ConnectionTimeout", "30", filePath))
    mCommandTimeout = CLng(mIniFile.GetPrivateProfileString("Database", "CommandTimeout", "30", filePath))
    mLogTableName = mIniFile.GetPrivateProfileString("Database", "LogTable", "AppLog", filePath)
    
    ' ConnectionStringBuilderの更新と設定
    UpdateConnectionStringBuilder
    With mConnectionStringBuilder
        .Server = mIniFile.GetPrivateProfileString("Database", "Server", "", filePath)
        .Database = mIniFile.GetPrivateProfileString("Database", "Database", "", filePath)
        .UserName = mIniFile.GetPrivateProfileString("Database", "UserName", "", filePath)
        .Password = mIniFile.GetPrivateProfileString("Database", "Password", "", filePath)
        .AdditionalParameters = mIniFile.GetPrivateProfileString("Database", "AdditionalParameters", "", filePath)
    End With
    
    ' 接続文字列の生成
    mConnectionString = mConnectionStringBuilder.BuildConnectionString()
    
    ' 設定の検証
    If (options And loValidate) = loValidate Then
        Dim validationResult As ValidationResult
        Set validationResult = IAppConfig_ValidateSettings(voFull)
        If Not validationResult.IsValid Then
            Err.Raise vbObjectError + 2, MODULE_NAME, "設定の検証に失敗しました: " & validationResult.GetSummary()
        End If
    End If
    
    mLock.ReleaseLock
    Exit Sub
    
ErrorHandler:
    mLock.ReleaseLock
    Err.Raise Err.Number, MODULE_NAME, "設定の読み込み中にエラーが発生しました: " & Err.Description
End Sub

Private Sub IAppConfig_SaveSettings(ByVal filePath As String, Optional ByVal options As SaveOptions = soNone)
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    ' バックアップの作成
    If (options And soBackup) = soBackup Then
        Dim backupPath As String
        backupPath = filePath & ".bak"
        mFileOperations.CopyFile filePath, backupPath
    End If
    
    ' 基本設定の保存
    mIniFile.WritePrivateProfileString "Logging", "LogFilePath", mLogFilePath, filePath
    mIniFile.WritePrivateProfileString "Logging", "LogLevel", CStr(mLogLevel), filePath
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

Private Sub IAppConfig_SetEncryptionKey(ByVal encryptionKey As String, Optional ByVal options As EncryptionOptions = eoNone)
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    ' 新しいキーの設定
    Dim oldKey As String
    oldKey = mEncryptionKey
    mEncryptionKey = encryptionKey
    
    ' キーバージョンの設定
    If (options And eoSetKeyVersion) = eoSetKeyVersion Then
        mEncryptionKeyVersion = Format$(Now, "yyyymmddhhnnss")
    End If
    
    ' 既存の暗号化された設定の再暗号化
    If (options And eoReEncrypt) = eoReEncrypt And oldKey <> "" Then
        Dim setting As Variant
        For Each setting In mEncryptedSettings
            Dim value As String
            value = mSettings(setting)
            value = mCrypto.Decrypt(value, oldKey)
            value = mCrypto.Encrypt(value, mEncryptionKey)
            mSettings.Remove setting
            mSettings.Add value, setting
        Next setting
    End If
    
    mLock.ReleaseLock
    Exit Sub
    
ErrorHandler:
    mLock.ReleaseLock
    Err.Raise Err.Number, MODULE_NAME, "暗号化キーの設定中にエラーが発生しました: " & Err.Description
End Sub

Private Sub IAppConfig_SetEnvironment(ByVal environment As String, Optional ByVal options As EnvironmentOptions = enoNone)
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    mCurrentEnvironment = environment
    
    ' 環境変更時の設定再読み込み
    If (options And enoReload) = enoReload Then
        IAppConfig_LoadSettings App.Path & "\config.ini", loApplyEnvironment
    End If
    
    ' 環境変更時の設定検証
    If (options And enoValidate) = enoValidate Then
        Dim validationResult As ValidationResult
        Set validationResult = IAppConfig_ValidateSettings(voEnvironment)
        If Not validationResult.IsValid Then
            Err.Raise vbObjectError + 3, MODULE_NAME, "環境設定の検証に失敗しました: " & validationResult.GetSummary()
        End If
    End If
    
    mLock.ReleaseLock
    Exit Sub
    
ErrorHandler:
    mLock.ReleaseLock
    Err.Raise Err.Number, MODULE_NAME, "環境の設定中にエラーが発生しました: " & Err.Description
End Sub

Private Function IAppConfig_ValidateSettings(Optional ByVal options As ValidationOptions = voNone) As ValidationResult
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    Dim result As New ValidationResult
    result.StartTime = Now
    result.State = vsRunning
    
    ' 基本的なデータ型の検証
    If (options And voDataType) = voDataType Then
        ' LogLevel の検証
        If mLogLevel < llTrace Or mLogLevel > llFatal Then
            result.AddError "LogLevel は無効な値です: " & mLogLevel
        End If
        
        ' 数値パラメータの検証
        If mMaxRetryCount < 0 Then result.AddError "MaxRetryCount は 0 以上である必要があります"
        If mMaxStackDepth < 1 Then result.AddError "MaxStackDepth は 1 以上である必要があります"
        If mTimerInterval < 100 Then result.AddError "TimerInterval は 100 ミリ秒以上である必要があります"
    End If
    
    ' 値の範囲検証
    If (options And voRange) = voRange Then
        If mMaxRetryCount > 10 Then result.AddWarning "MaxRetryCount が 10 を超えています"
        If mMaxStackDepth > 100 Then result.AddWarning "MaxStackDepth が 100 を超えています"
        If mTimerInterval > 60000 Then result.AddWarning "TimerInterval が 60 秒を超えています"
    End If
    
    ' 環境固有の設定の検証
    If (options And voEnvironment) = voEnvironment Then
        If mCurrentEnvironment = "" Then
            result.AddError "現在の環境が設定されていません"
        End If
    End If
    
    result.EndTime = Now
    result.State = IIf(result.IsValid, vsCompleted, vsFailed)
    
    mLock.ReleaseLock
    Set IAppConfig_ValidateSettings = result
    Exit Function
    
ErrorHandler:
    mLock.ReleaseLock
    Err.Raise Err.Number, MODULE_NAME, "設定の検証中にエラーが発生しました: " & Err.Description
End Function

Private Function IAppConfig_GetSettingHistory(Optional ByVal settingName As String, Optional ByVal options As HistoryOptions = hoNone) As Collection
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    Dim result As New Collection
    
    ' 指定された設定の履歴のみを取得
    If settingName <> "" Then
        Dim entry As Dictionary
        For Each entry In mSettingHistory
            If entry("SettingName") = settingName Then
                result.Add entry
            End If
        Next entry
    Else
        Set result = mSettingHistory
    End If
    
    ' ソート（新しい順）
    If (options And hoSort) = hoSort Then
        ' ソート処理の実装（別途ソートヘルパー関数が必要）
    End If
    
    ' 件数制限
    If (options And hoLimit) = hoLimit Then
        While result.Count > 100 ' デフォルトの制限値
            result.Remove result.Count
        Wend
    End If
    
    mLock.ReleaseLock
    Set IAppConfig_GetSettingHistory = result
    Exit Function
    
ErrorHandler:
    mLock.ReleaseLock
    Err.Raise Err.Number, MODULE_NAME, "設定履歴の取得中にエラーが発生しました: " & Err.Description
End Function

Private Function IAppConfig_CreateBackup(ByVal backupPath As String) As Boolean
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    ' 現在の設定を保存
    IAppConfig_SaveSettings backupPath, soEncrypt
    IAppConfig_CreateBackup = True
    
    mLock.ReleaseLock
    Exit Function
    
ErrorHandler:
    mLock.ReleaseLock
    IAppConfig_CreateBackup = False
End Function

Private Function IAppConfig_RestoreFromBackup(ByVal backupPath As String, Optional ByVal options As RestoreOptions = roNone) As Boolean
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    ' バックアップから設定を読み込む
    IAppConfig_LoadSettings backupPath, loDecrypt Or loValidate
    IAppConfig_RestoreFromBackup = True
    
    mLock.ReleaseLock
    Exit Function
    
ErrorHandler:
    mLock.ReleaseLock
    IAppConfig_RestoreFromBackup = False
End Function

Private Function IAppConfig_GetPerformanceMetrics() As Collection
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    Set IAppConfig_GetPerformanceMetrics = mPerformanceMetrics
    
    mLock.ReleaseLock
    Exit Function
    
ErrorHandler:
    mLock.ReleaseLock
    Err.Raise Err.Number, MODULE_NAME, "パフォーマンスメトリクスの取得中にエラーが発生しました: " & Err.Description
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

Private Property Get IAppConfig_LogLevel() As LogLevel
    IAppConfig_LogLevel = mLogLevel
End Property

Private Property Let IAppConfig_LogLevel(ByVal value As LogLevel)
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

Private Property Get IAppConfig_CurrentEnvironment() As String
    IAppConfig_CurrentEnvironment = mCurrentEnvironment
End Property

Private Property Get IAppConfig_EncryptionKeyVersion() As String
    IAppConfig_EncryptionKeyVersion = mEncryptionKeyVersion
End Property

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
            IDatabaseConfig_GetDatabaseSetting = Null
    End Select
    
    mLock.ReleaseLock
End Function

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