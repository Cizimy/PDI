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

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    mLogLevel = LevelInfo
    mLogDestination = DestFile
    mLogFilePath = modConfig.Settings.LogFilePath
    Set mLock = New clsLock
End Sub

Private Sub Class_Terminate()
    Set mLock = Nothing
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
    
    ' modFileIOを使用してファイルに書き込み
    modFileIO.WriteTextFile mLogFilePath, logMessage & vbCrLf, True
    
CleanUp:
    mLock.ReleaseLock
    Exit Sub
    
ErrorHandler:
    ' ログ出力自体のエラーは上位に伝播させない
    Debug.Print "ログ出力エラー: " & Err.Description
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
Private Function ShouldLog(ByVal errorCode As ErrorCode) As Boolean
    ' エラーコードからログレベルを判断
    Dim level As LogLevel
    
    Select Case errorCode
        Case ErrUnexpected, ErrSystemOutOfMemory
            level = LevelFatal
        Case ErrDbConnectionFailed, ErrNetworkError
            level = LevelError
        Case ErrFileNotFound, ErrInvalidInput
            level = LevelWarning
        Case Else
            level = LevelInfo
    End Select
    
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
    Select Case errorCode
        Case ErrUnexpected, ErrSystemOutOfMemory
            GetLogLevelText = "FATAL"
        Case ErrDbConnectionFailed, ErrNetworkError
            GetLogLevelText = "ERROR"
        Case ErrFileNotFound, ErrInvalidInput
            GetLogLevelText = "WARNING"
        Case Else
            GetLogLevelText = "INFO"
    End Select
End Function