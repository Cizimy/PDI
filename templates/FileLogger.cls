VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "FileLogger"

' ======================
' インターフェース実装
' ======================
Implements ILogger

' ======================
' プライベート変数
' ======================
Private mLogLevel As LogLevel
Private mLogDestination As LogDestination
Private mLogFilePath As String
Private mLock As clsLock
Private mPerformanceMonitor As clsPerformanceMonitor

' ======================
' 初期化・終了処理
' ======================
Public Sub Initialize(ByVal settings As ILoggerSettings)
    mLogLevel = LevelInfo
    mLogDestination = DestFile
    mLogFilePath = settings.LogFilePath
    Set mLock = New clsLock
    Set mPerformanceMonitor = New clsPerformanceMonitor
End Sub

Private Sub Class_Initialize()
    ' 初期化は Initialize メソッドで行う
End Sub

Private Sub Class_Terminate()
    Set mLock = Nothing
    Set mPerformanceMonitor = Nothing
End Sub

' ======================
' ILogger実装
' ======================
Private Sub ILogger_Log(ByVal moduleName As String, _
                       ByVal message As String, _
                       Optional ByVal errorCode As ErrorCode)
    ' ログレベルチェック
    If Not ShouldLog(errorCode) Then Exit Sub
    
    ' ログ出力の排他制御
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    Dim logMessage As String
    logMessage = BuildLogMessage(moduleName, message, errorCode)
    
    mPerformanceMonitor.StartMeasurement "File Write Operation"
    ' modFileIOを使用してファイルに書き込み
    modFileIO.WriteTextFile mLogFilePath, logMessage & vbCrLf, True
    
CleanUp:
    mLock.ReleaseLock
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "File Write Operation"
    End If
    Exit Sub
    
ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrFileAccessDenied
        .Description = "ログファイルへの書き込み中にエラーが発生しました: " & Err.Description
        .Category = ECFileIO
        .Source = MODULE_NAME
        .ProcedureName = "ILogger_Log"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "File Write Operation"
    End If
    Resume CleanUp
End Sub

Private Property Let ILogger_LogLevel(ByVal level As LogLevel)
    mLogLevel = level
End Property

Private Property Get ILogger_LogLevel() As LogLevel
    ILogger_LogLevel = mLogLevel
End Property

Private Property Let ILogger_LogDestination(ByVal destination As LogDestination)
    mLogDestination = destination
End Property

Private Property Get ILogger_LogDestination() As LogDestination
    ILogger_LogDestination = mLogDestination
End Property

' ======================
' プライベートメソッド
' ======================
Private Function GetLogLevelForErrorCode(ByVal errorCode As ErrorCode) As LogLevel
    Select Case errorCode
        Case ErrUnexpected, ErrSystemOutOfMemory
            GetLogLevelForErrorCode = LevelFatal
        Case ErrDbConnectionFailed, ErrNetworkError
            GetLogLevelForErrorCode = LevelError
        Case ErrFileNotFound, ErrInvalidInput
            GetLogLevelForErrorCode = LevelWarning
        Case Else
            GetLogLevelForErrorCode = LevelInfo
    End Select
End Function

Private Function ShouldLog(ByVal errorCode As ErrorCode) As Boolean
    ' エラーコードからログレベルを判断
    Dim level As LogLevel    
    level = GetLogLevelForErrorCode(errorCode)    
    ShouldLog = (level >= mLogLevel)
End Function

Private Function BuildLogMessage(ByVal moduleName As String, _
                               ByVal message As String, _
                               ByVal errorCode As ErrorCode) As String
    Dim builder As String
    
    ' タイムスタンプ
    builder = Format$(Now, DEFAULT_DATETIME_FORMAT) & " "
    
    ' ログレベル
    builder = builder & "[" & GetLogLevelText(errorCode) & "] "
    
    ' モジュール名
    builder = builder & "[" & moduleName & "] "
    
    ' エラーコード（存在する場合）
    If errorCode <> 0 Then
        builder = builder & "(Code: " & errorCode & ") "
    End If
    
    ' メッセージ本体
    builder = builder & message
    
    BuildLogMessage = builder
End Function

Private Function GetLogLevelText(ByVal errorCode As ErrorCode) As String
    ' エラーコードに対応するログレベルを取得し、テキストに変換
    Dim level As LogLevel
    level = GetLogLevelForErrorCode(errorCode
)
    
    Select Case level
        Case LevelFatal
            GetLogLevelText = "FATAL"
        Case LevelError
            GetLogLevelText = "ERROR"
        Case LevelWarning
            GetLogLevelText = "WARNING"
        Case LevelInfo
            GetLogLevelText = "INFO"
    End Select
End Function