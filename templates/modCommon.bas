Option Explicit

' ======================
' 定数定義
' ======================
Private Const MODULE_NAME As String = "modCommon"

' --- アプリケーション全体で使用する定数 ---
Public Const APPLICATION_NAME As String = "MyApp"
Public Const APPLICATION_VERSION As String = "1.0.0"
Public Const DEFAULT_LOG_FILE As String = "app.log"
Public Const DEFAULT_DATE_FORMAT As String = "yyyy/MM/dd"
Public Const MAX_RETRY_COUNT As Integer = 3

' ======================
' 列挙型定義
' ======================

' --- エラーコードカテゴリ ---
Public Enum ErrorCodeCategory
    ECGeneral = 1000    ' 一般エラー
    ECFileIO = 2000     ' ファイル操作エラー
    ECDatabase = 3000   ' データベースエラー
    ECNetwork = 4000    ' ネットワークエラー
    ECSystem = 5000     ' システムエラー
    ECSecurity = 6000   ' セキュリティエラー
End Enum

' --- エラーコード ---
Public Enum ErrorCode
    ' 一般エラー (1000-1999)
    ErrUnexpected = ErrorCodeCategory.ECGeneral
    ErrInvalidInput = ErrorCodeCategory.ECGeneral + 1
    
    ' ファイル操作エラー (2000-2999)
    ErrFileNotFound = ErrorCodeCategory.ECFileIO
    ErrFileInvalidFormat = ErrorCodeCategory.ECFileIO + 1
    ErrFileAccessDenied = ErrorCodeCategory.ECFileIO + 2
    
    ' データベースエラー (3000-3999)
    ErrDbConnectionFailed = ErrorCodeCategory.ECDatabase
    ErrDbQueryFailed = ErrorCodeCategory.ECDatabase + 1
    
    ' ネットワークエラー (4000-4999)
    ErrNetworkError = ErrorCodeCategory.ECNetwork
    ErrNetworkTimeout = ErrorCodeCategory.ECNetwork + 1
    
    ' システムエラー (5000-5999)
    ErrSystemOutOfMemory = ErrorCodeCategory.ECSystem
    ErrSystemResourceUnavailable = ErrorCodeCategory.ECSystem + 1
    
    ' セキュリティエラー (6000-6999)
    ErrSecurityAccessDenied = ErrorCodeCategory.ECSecurity
    ErrSecurityInvalidCredentials = ErrorCodeCategory.ECSecurity + 1
End Enum

' --- ログレベル ---
Public Enum LogLevel
    LevelDebug
    LevelInfo
    LevelWarning
    LevelError
    LevelFatal
End Enum

' --- ログ出力先 ---
Public Enum LogDestination
    DestNone
    DestFile
    DestDatabase
    DestEventLog
    DestConsole
    DestEmail
End Enum

' --- ファイルアクセスモード ---
Public Enum FileAccessMode
    ModeReadOnly = 1
    ModeReadWrite = 2
    ModeAppend = 3
    ModeExclusive = 4
End Enum

' ======================
' 型定義
' ======================

' --- ファイル情報 ---
Public Type FileInfo
    Name As String
    Path As String
    Size As Long
    Created As Date
    LastModified As Date
    FileType As String
    Attributes As Long
End Type

' --- エラー情報 ---
Public Type ErrorInfo
    Code As ErrorCode
    Category As ErrorCodeCategory
    Description As String
    Source As String
    ProcedureName As String
    StackTrace As String
    OccurredAt As Date
    AdditionalInfo As String
End Type

' --- アプリケーション設定 ---
Public Type AppSettings
    LogLevel As LogLevel
    LogDestination As LogDestination
    LogFilePath As String
    LogRetentionDays As Long
    MaxLogFileSize As Long
    DatabaseConnectionString As String
    SecurityLevel As SecurityLevel
    EncryptionKey As String
    PerformanceMonitoringEnabled As Boolean
    DiagnosticsEnabled As Boolean
End Type

' --- セキュリティレベル ---
Public Enum SecurityLevel
    LevelLow = 1
    LevelMedium = 2
    LevelHigh = 3
    LevelExtreme = 4
End Enum

' ======================
' グローバル変数
' ======================
''' <summary>
''' アプリケーション全体で共有する設定情報
''' グローバル変数として定義する理由：
''' - アプリケーションの起動時に一度だけ読み込み、以降は変更されないため
''' - 各モジュールから頻繁にアクセスされるため、パフォーマンスを考慮
''' - スレッドセーフな実装により、複数のプロセスからの同時アクセスに対応
''' </summary>
Private mAppSettings As AppSettings
Private mSettingsLock As New clsLock ' スレッドセーフ用のロックオブジェクト

' ======================
' プロパティ
' ======================
Public Property Get AppSettings() As AppSettings
    mSettingsLock.AcquireLock
    AppSettings = mAppSettings
    mSettingsLock.ReleaseLock
End Property

Public Property Let AppSettings(ByVal Value As AppSettings)
    mSettingsLock.AcquireLock
    mAppSettings = Value
    mSettingsLock.ReleaseLock
End Property

' ======================
' 初期化・終了処理
' ======================
Public Sub InitializeModule()
    ' モジュールの初期化処理
    Set mSettingsLock = New clsLock
    LoadDefaultSettings
End Sub

Public Sub TerminateModule()
    ' モジュールの終了処理
    Set mSettingsLock = Nothing
End Sub

' ======================
' 内部関数
' ======================
Private Sub LoadDefaultSettings()
    With mAppSettings
        .LogLevel = LevelInfo
        .LogDestination = DestFile
        .LogFilePath = DEFAULT_LOG_FILE
        .LogRetentionDays = 30
        .MaxLogFileSize = 10485760 ' 10MB
        .SecurityLevel = LevelMedium
        .PerformanceMonitoringEnabled = True
        .DiagnosticsEnabled = True
    End With
End Sub

' ======================
' 診断・モニタリング機能
' ======================
Public Function GetModuleStatus() As String
    Dim status As String
    status = "ModCommon Status Report" & vbCrLf & _
            "Time: " & Now & vbCrLf & _
            "Security Level: " & mAppSettings.SecurityLevel & vbCrLf & _
            "Performance Monitoring: " & mAppSettings.PerformanceMonitoringEnabled & vbCrLf & _
            "Diagnostics: " & mAppSettings.DiagnosticsEnabled
    GetModuleStatus = status
End Function

' ======================
' テストサポート機能
' ======================
#If DEBUG Then
    Public Sub ResetModule()
        Set mSettingsLock = Nothing
        LoadDefaultSettings
        Set mSettingsLock = New clsLock
    End Sub
    
    Public Function ValidateSettings() As Boolean
        ' 設定値の妥当性検証
        With mAppSettings
            ValidateSettings = _
                .LogLevel >= LevelDebug And .LogLevel <= LevelFatal And _
                .LogDestination >= DestNone And .LogDestination <= DestEmail And _
                .LogRetentionDays > 0 And _
                .MaxLogFileSize > 0 And _
                .SecurityLevel >= LevelLow And .SecurityLevel <= LevelExtreme
        End With
    End Function
#End If