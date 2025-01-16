Attribute VB_Name = "modLogger"
Option Explicit

'*******************************************************************************
' モジュール: modLogger
' 目的：     システム全体のログ管理
' 作成日：   2025/01/17
'*******************************************************************************

' ログレベルの定義
Public Enum LogLevel
    llDebug = 1
    llInfo = 2
    llWarning = 3
    llError = 4
    llCritical = 5
End Enum

' ログ設定
Private Type LogConfig
    LogPath As String
    MinLogLevel As LogLevel
    MaxFileSize As Long
    RotateCount As Integer
    EnableConsole As Boolean
End Type

' デフォルト設定
Private Const DEFAULT_LOG_PATH As String = "app.log"
Private Const DEFAULT_MAX_FILE_SIZE As Long = 5242880 ' 5MB
Private Const DEFAULT_ROTATE_COUNT As Integer = 5
Private Const DEFAULT_MIN_LOG_LEVEL As LogLevel = llInfo

' 現在の設定
Private mConfig As LogConfig

'*******************************************************************************
' 目的：    モジュールの初期化
' 引数：    なし
' 戻り値：  なし
'*******************************************************************************
Public Sub Initialize()
    ' デフォルト設定の適用
    With mConfig
        .LogPath = DEFAULT_LOG_PATH
        .MinLogLevel = DEFAULT_MIN_LOG_LEVEL
        .MaxFileSize = DEFAULT_MAX_FILE_SIZE
        .RotateCount = DEFAULT_ROTATE_COUNT
        .EnableConsole = True
    End With
End Sub

'*******************************************************************************
' 目的：    ログ設定の更新
' 引数：    logPath - ログファイルのパス
'           minLevel - 最小ログレベル
'           maxSize - 最大ファイルサイズ
'           rotateCount - ローテーション数
'           enableConsole - コンソール出力の有効化
' 戻り値：  なし
'*******************************************************************************
Public Sub Configure(Optional ByVal logPath As String = "", _
                    Optional ByVal minLevel As LogLevel = llInfo, _
                    Optional ByVal maxSize As Long = -1, _
                    Optional ByVal rotateCount As Integer = -1, _
                    Optional ByVal enableConsole As Boolean = True)
                    
    With mConfig
        If logPath <> "" Then .LogPath = logPath
        .MinLogLevel = minLevel
        If maxSize > 0 Then .MaxFileSize = maxSize
        If rotateCount > 0 Then .RotateCount = rotateCount
        .EnableConsole = enableConsole
    End With
End Sub

'*******************************************************************************
' 目的：    ログメッセージの記録
' 引数：    message - ログメッセージ
'           level - ログレベル
'           source - 発生源
' 戻り値：  なし
'*******************************************************************************
Public Sub LogMessage(ByVal message As String, _
                     Optional ByVal level As LogLevel = llInfo, _
                     Optional ByVal source As String = "")
                     
    If level < mConfig.MinLogLevel Then Exit Sub
    
    Dim logEntry As String
    Dim timestamp As String
    
    ' タイムスタンプの生成
    timestamp = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    ' ログエントリの作成
    logEntry = timestamp & vbTab & _
              GetLogLevelName(level) & vbTab & _
              IIf(source <> "", source & vbTab, "") & _
              message
              
    ' ファイルへの書き込み
    WriteToLogFile logEntry
    
    ' コンソール出力
    If mConfig.EnableConsole Then
        Debug.Print logEntry
    End If
End Sub

'*******************************************************************************
' 目的：    デバッグログの記録
' 引数：    message - ログメッセージ
'           source - 発生源
' 戻り値：  なし
'*******************************************************************************
Public Sub Debug(ByVal message As String, Optional ByVal source As String = "")
    LogMessage message, llDebug, source
End Sub

'*******************************************************************************
' 目的：    情報ログの記録
' 引数：    message - ログメッセージ
'           source - 発生源
' 戻り値：  なし
'*******************************************************************************
Public Sub Info(ByVal message As String, Optional ByVal source As String = "")
    LogMessage message, llInfo, source
End Sub

'*******************************************************************************
' 目的：    警告ログの記録
' 引数：    message - ログメッセージ
'           source - 発生源
' 戻り値：  なし
'*******************************************************************************
Public Sub Warning(ByVal message As String, Optional ByVal source As String = "")
    LogMessage message, llWarning, source
End Sub

'*******************************************************************************
' 目的：    エラーログの記録
' 引数：    message - ログメッセージ
'           source - 発生源
' 戻り値：  なし
'*******************************************************************************
Public Sub Error(ByVal message As String, Optional ByVal source As String = "")
    LogMessage message, llError, source
End Sub

'*******************************************************************************
' 目的：    重大エラーログの記録
' 引数：    message - ログメッセージ
'           source - 発生源
' 戻り値：  なし
'*******************************************************************************
Public Sub Critical(ByVal message As String, Optional ByVal source As String = "")
    LogMessage message, llCritical, source
End Sub

'*******************************************************************************
' 目的：    ログレベル名の取得
' 引数：    level - ログレベル
' 戻り値：  ログレベルの文字列表現
'*******************************************************************************
Private Function GetLogLevelName(ByVal level As LogLevel) As String
    Select Case level
        Case llDebug
            GetLogLevelName = "DEBUG"
        Case llInfo
            GetLogLevelName = "INFO"
        Case llWarning
            GetLogLevelName = "WARNING"
        Case llError
            GetLogLevelName = "ERROR"
        Case llCritical
            GetLogLevelName = "CRITICAL"
        Case Else
            GetLogLevelName = "UNKNOWN"
    End Select
End Function

'*******************************************************************************
' 目的：    ログファイルへの書き込み
' 引数：    logEntry - ログエントリ
' 戻り値：  なし
'*******************************************************************************
Private Sub WriteToLogFile(ByVal logEntry As String)
    On Error Resume Next
    
    Dim fileNum As Integer
    
    ' ファイルサイズのチェックとローテーション
    CheckFileSize
    
    ' ログの書き込み
    fileNum = FreeFile
    Open mConfig.LogPath For Append As fileNum
    Print #fileNum, logEntry
    Close fileNum
End Sub

'*******************************************************************************
' 目的：    ログファイルのサイズチェックとローテーション
' 引数：    なし
' 戻り値：  なし
'*******************************************************************************
Private Sub CheckFileSize()
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' ファイルが存在し、サイズ制限を超えている場合
    If fso.FileExists(mConfig.LogPath) Then
        If fso.GetFile(mConfig.LogPath).Size >= mConfig.MaxFileSize Then
            RotateLogFiles
        End If
    End If
    
    Set fso = Nothing
End Sub

'*******************************************************************************
' 目的：    ログファイルのローテーション
' 引数：    なし
' 戻り値：  なし
'*******************************************************************************
Private Sub RotateLogFiles()
    On Error Resume Next
    
    Dim fso As Object
    Dim i As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 最も古いログファイルの削除
    If fso.FileExists(mConfig.LogPath & "." & mConfig.RotateCount) Then
        fso.DeleteFile mConfig.LogPath & "." & mConfig.RotateCount
    End If
    
    ' ファイルの移動
    For i = mConfig.RotateCount - 1 To 1 Step -1
        If fso.FileExists(mConfig.LogPath & "." & i) Then
            fso.MoveFile mConfig.LogPath & "." & i, _
                        mConfig.LogPath & "." & (i + 1)
        End If
    Next i
    
    ' 現在のログファイルの移動
    If fso.FileExists(mConfig.LogPath) Then
        fso.MoveFile mConfig.LogPath, mConfig.LogPath & ".1"
    End If
    
    Set fso = Nothing
End Sub
