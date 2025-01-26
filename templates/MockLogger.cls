VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' 警告: このクラスは開発時のテスト目的でのみ使用し、
' 本番環境では使用しないでください。
' ======================

Implements ILogger

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "MockLogger"

' ======================
' メンバ変数
' ======================
Private Type LogEntry
    ModuleName As String
    Message As String
    ErrorCode As ErrorCode
    Timestamp As Date
End Type

' エラー情報を格納する型
Private Type LogOperationResult
    Success As Boolean
    Value As Variant
    ErrorDetail As typErrorDetail
End Type

Private mLogEntries As Collection
Private mLogLevel As LogLevelEnum
Private mLogDestination As LogDestination

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mLogEntries = New Collection
    mLogLevel = LOG_LEVEL_INFO
    mLogDestination = DestNone
End Sub

Private Sub Class_Terminate()
    Set mLogEntries = Nothing
End Sub

' ======================
' ILogger実装
' ======================
Private Sub ILogger_Log(ByVal moduleName As String, ByVal message As String, Optional ByVal errorCode As ErrorCode)
    On Error GoTo ErrorHandler
    
    Dim entry As LogEntry
    With entry
        .ModuleName = moduleName
        .Message = message
        .ErrorCode = errorCode
        .Timestamp = Now
    End With
    
    mLogEntries.Add entry
    Exit Sub

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "モックロガーへのログ追加中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "ILogger_Log"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
End Sub

Private Property Let ILogger_LogLevel(ByVal level As LogLevelEnum)
    mLogLevel = level
End Property

Private Property Get ILogger_LogLevel() As LogLevelEnum
    ILogger_LogLevel = mLogLevel
End Property

Private Property Let ILogger_LogDestination(ByVal destination As LogDestination)
    mLogDestination = destination
End Property

Private Property Get ILogger_LogDestination() As LogDestination
    ILogger_LogDestination = mLogDestination
End Property

' ======================
' テストサポートメソッド（開発環境専用）
' ======================

''' <summary>
''' ログエントリの総数を取得（テスト用）
''' </summary>
Private Property Get LogCount() As Long
    LogCount = mLogEntries.Count
End Property

''' <summary>
''' 指定されたインデックスのログエントリを取得（テスト用）
''' </summary>
''' <param name="index">取得するログエントリのインデックス</param>
''' <returns>フォーマット済みのログエントリ文字列、無効なインデックスの場合は空文字列</returns>
Private Function GetLogEntry(ByVal index As Long) As LogOperationResult
    On Error GoTo ErrorHandler
    
    Dim result As LogOperationResult
    result.Success = True
    
    If index < 1 Or index > mLogEntries.Count Then
        result.Success = False
        With result.ErrorDetail
            .ErrorCode = ERR_INVALID_PARAMETER
            .Description = "無効なインデックスが指定されました: " & index
            .Category = ECValidation
            .Source = MODULE_NAME
            .ProcedureName = "GetLogEntry"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        GetLogEntry = result
        Exit Function
    End If
    
    Dim entry As LogEntry
    entry = mLogEntries(index)
    
    result.Value = Format$(entry.Timestamp, "yyyy/mm/dd hh:nn:ss") & " " & _
                  "[" & entry.ModuleName & "] " & _
                  IIf(entry.ErrorCode <> 0, "(Code: " & entry.ErrorCode & ") ", "") & _
                  entry.Message
    GetLogEntry = result
    Exit Function

ErrorHandler:
    result.Success = False
    With result.ErrorDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "ログエントリの取得中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "GetLogEntry"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    GetLogEntry = result
End Function

''' <summary>
''' すべてのログエントリをクリア（テスト用）
''' </summary>
Private Sub ClearLogs()
    Set mLogEntries = New Collection
End Sub

''' <summary>
''' 特定のモジュールからのログエントリ数を取得（テスト用）
''' </summary>
''' <param name="moduleName">対象のモジュール名</param>
''' <returns>指定されたモジュールのログエントリ数</returns>
Private Function GetLogCountForModule(ByVal moduleName As String) As LogOperationResult
    On Error GoTo ErrorHandler
    
    Dim result As LogOperationResult
    result.Success = True
    
    Dim count As Long
    Dim i As Long
    Dim entry As LogEntry
    
    For i = 1 To mLogEntries.Count
        entry = mLogEntries(i)
        If entry.ModuleName = moduleName Then
            count = count + 1
        End If
    Next i
    
    result.Value = count
    GetLogCountForModule = result
    Exit Function

ErrorHandler:
    result.Success = False
    With result.ErrorDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "モジュール別ログ数の取得中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "GetLogCountForModule"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    GetLogCountForModule = result
End Function

''' <summary>
''' 特定のエラーコードのログエントリ数を取得（テスト用）
''' </summary>
''' <param name="errorCode">対象のエラーコード</param>
''' <returns>指定されたエラーコードのログエントリ数</returns>
Private Function GetLogCountForErrorCode(ByVal errorCode As ErrorCode) As LogOperationResult
    On Error GoTo ErrorHandler
    
    Dim result As LogOperationResult
    result.Success = True
    
    Dim count As Long
    Dim i As Long
    Dim entry As LogEntry
    
    For i = 1 To mLogEntries.Count
        entry = mLogEntries(i)
        If entry.ErrorCode = errorCode Then
            count = count + 1
        End If
    Next i
    
    result.Value = count
    GetLogCountForErrorCode = result
    Exit Function

ErrorHandler:
    result.Success = False
    With result.ErrorDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "エラーコード別ログ数の取得中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "GetLogCountForErrorCode"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    GetLogCountForErrorCode = result
End Function

''' <summary>
''' 最後のログエントリを取得（テスト用）
''' </summary>
''' <returns>最後のログエントリ、ログが空の場合は空文字列</returns>
Private Function GetLastLogEntry() As LogOperationResult
    On Error GoTo ErrorHandler
    
    Dim result As LogOperationResult
    result.Success = True
    
    If mLogEntries.Count = 0 Then
        result.Success = False
        With result.ErrorDetail
            .ErrorCode = ERR_NO_DATA
            .Description = "ログエントリが存在しません"
            .Category = ECValidation
            .Source = MODULE_NAME
            .ProcedureName = "GetLastLogEntry"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        GetLastLogEntry = result
        Exit Function
    End If
    
    GetLastLogEntry = GetLogEntry(mLogEntries.Count)
    Exit Function

ErrorHandler:
    result.Success = False
    With result.ErrorDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "最後のログエントリの取得中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "GetLastLogEntry"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    GetLastLogEntry = result
End Function