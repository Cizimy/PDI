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
''' <summary>
''' ロガー設定を初期化します。
''' </summary>
''' <param name="connectionString">データベース接続文字列（オプション）。
''' 指定しない場合、ファイルログのみが有効になります。</param>
''' <param name="logLevel">ログレベル（オプション、デフォルト：INFO）</param>
''' <param name="logFilePath">ログファイルパス（オプション、デフォルト：アプリケーションパス）</param>
''' <param name="logTableName">ログテーブル名（オプション、デフォルト：AppLog）</param>
''' <param name="logEventSource">イベントソース名（オプション、デフォルト：アプリケーション名）</param>
''' <param name="timerInterval">タイマー間隔（オプション、デフォルト：1000ms）</param>
Public Sub Initialize(ByVal connectionString As String, _
                     Optional ByVal logLevel As LogLevelEnum = LOG_LEVEL_INFO, _
                     Optional ByVal logFilePath As String = "", _
                     Optional ByVal logTableName As String = "AppLog", _
                     Optional ByVal logEventSource As String = "", _
                     Optional ByVal timerInterval As Long = 1000)
    On Error GoTo ErrorHandler
    
    mLogLevel = logLevel
    Set mLogDestination = New FileLogDestination
    mLogFilePath = IIf(logFilePath = "", App.Path & "\" & DEFAULT_LOG_FILE, logFilePath)
    mLogTableName = logTableName
    mLogEventSource = IIf(logEventSource = "", APPLICATION_NAME, logEventSource)
    mTimerInterval = timerInterval
    mConnectionString = connectionString
    
    Exit Sub

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "ロガー設定の初期化中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = "DefaultLoggerSettings"
        .ProcedureName = "Class_Initialize"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
End Sub

Private Sub Class_Initialize()
    ' デフォルト設定で初期化
    Initialize "", LOG_LEVEL_INFO, "", "AppLog", "", 1000
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
' テストサポート機能（開発環境専用）
' 警告: これらのメソッドは開発時のテスト目的でのみ使用し、
' 本番環境では使用しないでください。
' ======================
#If DEBUG Then
''' <summary>
''' デバッグ用の設定を作成（テスト用）
''' 詳細なログ出力と高頻度な更新を行う設定を生成します
''' </summary>
Private Function CreateDebugSettings() As DefaultLoggerSettings
    On Error GoTo ErrorHandler
    
    Set CreateDebugSettings = New DefaultLoggerSettings
    With CreateDebugSettings
        .LogLevel = LOG_LEVEL_DEBUG
        Set .LogDestination = New ConsoleLogDestination
        .TimerInterval = 100 ' より頻繁な更新
    End With
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "デバッグ用ロガー設定の作成中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = "DefaultLoggerSettings"
        .ProcedureName = "CreateDebugSettings"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    Set CreateDebugSettings = Nothing
End Function

''' <summary>
''' 本番用の設定を作成（テスト用）
''' 標準的なログレベルと低頻度な更新を行う設定を生成します
''' </summary>
Private Function CreateProductionSettings() As DefaultLoggerSettings
    On Error GoTo ErrorHandler
    
    Set CreateProductionSettings = New DefaultLoggerSettings
    With CreateProductionSettings
        .LogLevel = LOG_LEVEL_INFO
        Set .LogDestination = New FileLogDestination
        .TimerInterval = 5000 ' より低頻度な更新
    End With
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "本番用ロガー設定の作成中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = "DefaultLoggerSettings"
        .ProcedureName = "CreateProductionSettings"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    Set CreateProductionSettings = Nothing
End Function

''' <summary>
''' データベースログ用の設定を作成（テスト用）
''' データベースへのログ出力に特化した設定を生成します
''' </summary>
Private Function CreateDatabaseSettings() As DefaultLoggerSettings
    On Error GoTo ErrorHandler
    
    Set CreateDatabaseSettings = New DefaultLoggerSettings
    With CreateDatabaseSettings
        .LogLevel = LOG_LEVEL_INFO
        Set .LogDestination = New DatabaseLogDestination
        .LogTableName = "SystemLog"
        .TimerInterval = 2000
    End With
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "データベース用ロガー設定の作成中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = "DefaultLoggerSettings"
        .ProcedureName = "CreateDatabaseSettings"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    Set CreateDatabaseSettings = Nothing
End Function
#End If