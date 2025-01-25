' ======================
' 1.1 modCommon (共通定義)
' ======================
Option Explicit

' --- アプリケーション全体で使用する定数 ---
Public Const APPLICATION_NAME As String = "MyApp"
Public Const APPLICATION_VERSION As String = "1.0.0"
Public Const DEFAULT_LOG_FILE As String = "app.log"
Public Const DEFAULT_DATE_FORMAT As String = "yyyy/MM/dd"
Public Const MAX_RETRY_COUNT As Integer = 3 ' リトライ回数

' --- アプリケーション全体で使用する列挙型 ---
' --- エラーコード ---
' modError から移動・統合
Public Enum ErrorCode
    ERR_UNEXPECTED = vbObjectError + 1000             ' 予期せぬエラー
    ERR_FILEIO_NOT_FOUND = vbObjectError + 1001       ' ファイルが見つからない
    ERR_FILEIO_INVALID_FORMAT = vbObjectError + 1002   ' ファイル形式エラー
    ERR_FILEIO_ACCESS_DENIED = vbObjectError + 1003   ' アクセス拒否
    ERR_DATABASE_CONNECTION_FAILED = vbObjectError + 1004 ' データベース接続エラー
    ERR_DATABASE_QUERY_FAILED = vbObjectError + 1005   ' データベースクエリエラー
    ERR_INPUT_INVALID = vbObjectError + 1006          ' 無効な入力
    ERR_NETWORK_ERROR = vbObjectError + 1007          ' ネットワークエラー
    ERR_NETWORK_TIMEOUT = vbObjectError + 1008        ' タイムアウト
    ERR_SYSTEM_OUT_OF_MEMORY = vbObjectError + 1009   ' メモリ不足
    ' ... (その他のエラーコード)
End Enum

' --- ログレベル ---
Public Enum LogLevelEnum
    LOG_LEVEL_DEBUG = 0
    LOG_LEVEL_INFO = 1
    LOG_LEVEL_WARNING = 2
    LOG_LEVEL_ERROR = 3
    LOG_LEVEL_FATAL = 4
End Enum

' --- ログ出力先 ---
Public Enum LogDestinationEnum
    LOG_DESTINATION_NONE = 0
    LOG_DESTINATION_FILE = 1
    LOG_DESTINATION_DATABASE = 2
    LOG_DESTINATION_EVENTLOG = 3
    LOG_DESTINATION_CONSOLE = 4 ' デバッグ用
    LOG_DESTINATION_EMAIL = 5   ' メール通知 (将来対応)
End Enum

' --- ファイルアクセスモード ---
' 改善案：他のモジュールで定義されている famReadOnly, famReadWrite との統合を検討 -> 列挙型名とプレフィックスを変更
Public Enum FileAccessModeEnum
    FAM_READ_ONLY = 1
    FAM_READ_WRITE = 2
End Enum

' --- アプリケーション全体で使用するユーザー定義型 ---
' --- ファイル情報 ---
Public Type typFileInfo
    Name As String
    Path As String
    Size As Long
    Created As Date
    LastModified As Date
    Type As String ' e.g., "Text", "Image", "Document"
    Attributes As Long ' e.g., Read-only, Hidden, System (vbFileAttribute 列挙型を使用)
End Type

' --- エラー情報 ---
' modError から移動・統合
Public Type typErrorDetail
    ErrorCode As ErrorCode    ' エラーコード
    Description As String     ' エラー説明
    Source As String          ' エラー発生元 (任意でモジュール名等)
    ProcedureName As String   ' どのプロシージャで起きたか
    StackTrace As String      ' スタックトレース
    OccurredAt As Date       ' エラー発生時刻
End Type

' --- アプリケーション設定 ---
' modConfig から移動
Public Type typAppSettings
    LogLevel As LogLevelEnum
    LogDestination As LogDestinationEnum
    DatabaseConnectionString As String ' データベース接続文字列
    ' ... (その他の設定項目)
End Type

' --- アプリケーション設定を保持するグローバル変数 ---
' modConfig から移動
''' <summary>
''' アプリケーション全体で共有する設定情報
''' グローバル変数として定義する理由：
''' - アプリケーションの起動時に一度だけ読み込み、以降は変更されないため
''' - 各モジュールから頻繁にアクセスされるため、パフォーマンスを考慮
''' </summary>
Public gAppSettings As typAppSettings


' ======================
' 1.2 modConfig (設定管理)
' ======================
Option Explicit

' --- API関数の宣言 (INI読み込み用) ---
' modCommonに移動
' Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
'     ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
'     ByVal lpDefault As String, ByVal lpReturnedString As String, _
'     ByVal nSize As Long, ByVal lpFileName As String) As Long

' --- 設定ファイルのパス ---
Private Const CONFIG_FILE_PATH As String = "config.ini" ' INIファイルのパス

' --- 設定ファイルの読み込み ---
Public Sub LoadConfig()
    Dim buffer As String
    Dim result As Long
    Dim configFilePath As String

    ' 設定ファイルのフルパスを取得（実行ファイルと同じディレクトリ）
    configFilePath = App.Path & "\" & CONFIG_FILE_PATH

    ' LogLevel の読み込み
    buffer = String$(255, vbNullChar)
    result = GetPrivateProfileString("Settings", "LogLevel", "INFO", buffer, Len(buffer), configFilePath)
    If result > 0 Then
        Select Case UCase(Left$(buffer, result))
            Case "DEBUG"
                gAppSettings.LogLevel = LOG_LEVEL_DEBUG
            Case "INFO"
                gAppSettings.LogLevel = LOG_LEVEL_INFO
            Case "WARNING"
                gAppSettings.LogLevel = LOG_LEVEL_WARNING
            Case "ERROR"
                gAppSettings.LogLevel = LOG_LEVEL_ERROR
            Case "FATAL"
                gAppSettings.LogLevel = LOG_LEVEL_FATAL
            Case Else
                gAppSettings.LogLevel = LOG_LEVEL_INFO ' デフォルト値
        End Select
    Else
        gAppSettings.LogLevel = LOG_LEVEL_INFO ' デフォルト値
    End If

    ' LogDestination の読み込み
    buffer = String$(255, vbNullChar)
    result = GetPrivateProfileString("Settings", "LogDestination", "FILE", buffer, Len(buffer), configFilePath)
    If result > 0 Then
        Select Case UCase(Left$(buffer, result))
            Case "NONE"
                gAppSettings.LogDestination = LOG_DESTINATION_NONE
            Case "FILE"
                gAppSettings.LogDestination = LOG_DESTINATION_FILE
            Case "DATABASE"
                gAppSettings.LogDestination = LOG_DESTINATION_DATABASE
            Case "EVENTLOG"
                gAppSettings.LogDestination = LOG_DESTINATION_EVENTLOG
            Case "CONSOLE"
                gAppSettings.LogDestination = LOG_DESTINATION_CONSOLE
            Case "EMAIL"
                gAppSettings.LogDestination = LOG_DESTINATION_EMAIL
            Case Else
                gAppSettings.LogDestination = LOG_DESTINATION_FILE ' デフォルト値
        End Select
    Else
        gAppSettings.LogDestination = LOG_DESTINATION_FILE ' デフォルト値
    End If

    ' DatabaseConnectionString の読み込み
    buffer = String$(255, vbNullChar)
    result = GetPrivateProfileString("Database", "ConnectionString", "", buffer, Len(buffer), configFilePath)
    If result > 0 Then
        gAppSettings.DatabaseConnectionString = Left$(buffer, result)
    Else
        gAppSettings.DatabaseConnectionString = "" ' デフォルト値
    End If

    ' ... (その他の設定項目の読み込み)
End Sub


' ======================
' 1.3 modWindowsAPI (Windows API 宣言)
' ======================
Option Explicit

'--- API関数の宣言 (INI読み込み用) ---
' modConfigより移動
Public Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpDefault As String, ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long

'--- その他のAPI宣言 ---
' 例：ファイル属性を取得するAPI
Public Declare PtrSafe Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
' 例：ファイル属性を設定するAPI
Public Declare PtrSafe Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

'--- タイマー関連 ---
Public Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Public Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long

' --- APIで使用する定数 ---
' 例：ファイル属性
Public Const FILE_ATTRIBUTE_READONLY As Long = &H1
Public Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
Public Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Public Const INVALID_FILE_ATTRIBUTES As Long = -1

' ... その他のAPI宣言と定数 ...


' ======================
' 1.4 clsCallStack (呼び出し履歴管理クラス)
' ======================
Option Explicit

Private Const MAX_STACK_TRACE_DEPTH As Long = 10 ' スタックトレースの最大深さ

Private stack As Collection

Private Sub Class_Initialize()
    Set stack = New Collection
End Sub

Private Sub Class_Terminate()
    Set stack = Nothing
End Sub

' 呼び出し履歴のプッシュ
Public Sub Push(ModuleName As String, ProcedureName As String)
    If stack.Count < MAX_STACK_TRACE_DEPTH Then
        stack.Add ModuleName & "." & ProcedureName
    End If
End Sub

' 呼び出し履歴のポップ
Public Function Pop() As String
    If stack.Count > 0 Then
        Pop = stack(stack.Count)
        stack.Remove stack.Count
    End If
End Function

' スタックトレースの取得
Public Property Get StackTrace() As String
    Dim i As Long
    For i = stack.Count To 1 Step -1
        StackTrace = StackTrace & "  " & stack(i) & vbCrLf
    Next i
End Property

' 呼び出し回数の取得
Public Property Get Count() As Long
    Count = stack.Count
End Property


' ======================
' 1.5 modError (エラーハンドリング)
' ======================
Option Explicit

'---------------------------------
' 定数
'---------------------------------
Private Const MODULE_NAME As String = "modError"

'---------------------------------
' インターフェースの宣言
'---------------------------------
Private logger As ILogger ' ILogger インターフェース
Private notifier As IUserNotifier ' IUserNotifier インターフェース

'---------------------------------
' プロパティの設定
'---------------------------------
' ロガーの設定
Public Property Set Logger(ByVal obj As ILogger)
    Set logger = obj
End Property

' 通知方法の設定
Public Property Set Notifier(ByVal obj As IUserNotifier)
    Set notifier = obj
End Property


'---------------------------------
' 汎用エラーメッセージ通知
'---------------------------------
Public Sub NotifyUser(ByVal errorDetail As typErrorDetail, _
                        Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
                        Optional ByVal title As String = "エラー")

    If notifier Is Nothing Then
        ' デフォルトの通知方法 (MsgBox)
        ' アイコンをエラーコードに応じて変更
        Dim icon As VbMsgBoxStyle
        Select Case errorDetail.ErrorCode
            Case ERR_DATABASE_CONNECTION_FAILED, ERR_NETWORK_TIMEOUT, ERR_SYSTEM_OUT_OF_MEMORY ' 重大
                icon = vbCritical
            Case ERR_FILEIO_NOT_FOUND, ERR_INPUT_INVALID, ERR_FILEIO_INVALID_FORMAT, ERR_FILEIO_ACCESS_DENIED, ERR_NETWORK_ERROR
                icon = vbExclamation
            Case Else
                icon = vbInformation
        End Select

        MsgBox GetErrorMessage(errorDetail.ErrorCode) & vbCrLf & errorDetail.Description, buttons Or icon, title
    Else
        ' 設定された通知方法を使用
        notifier.Notify errorDetail, buttons, title
    End If

End Sub

'---------------------------------
' エラー情報をログに出力
'---------------------------------
Public Sub LogError(ByVal errorDetail As typErrorDetail)

    If logger Is Nothing Then Exit Sub ' ロガーが設定されていない場合は何もしない

    ' スタックトレースは errorDetail.StackTrace に格納済みを想定（typErrorDetailの定義より）
    Dim fullMessage As String
    fullMessage = GetErrorMessage(errorDetail.ErrorCode) & " (Source: " & errorDetail.Source & ")" & " (Procedure: " & errorDetail.ProcedureName & ")"
    If errorDetail.Description <> "" Then
        fullMessage = fullMessage & " - " & errorDetail.Description
    End If
    fullMessage = fullMessage & vbCrLf & "StackTrace:" & vbCrLf & errorDetail.StackTrace

    logger.Log MODULE_NAME, fullMessage, errorDetail.ErrorCode

End Sub

'---------------------------------
' エラーコード→メッセージ変換
'---------------------------------
Friend Function GetErrorMessage(ByVal code As ErrorCode) As String
    ' エラーメッセージを外部ファイル（リソースファイルやデータベースなど）から取得する
    ' ここでは簡略化のために、モジュール内の関数 GetErrorMessageFromResource を使用
    ' 将来的には、modConfig などから設定を読み込み、エラーメッセージの取得方法を切り替えられるようにする
    GetErrorMessage = GetErrorMessageFromResource(code)
End Function

'---------------------------------
' エラーメッセージをリソースから取得 (ダミー実装)
'---------------------------------
Private Function GetErrorMessageFromResource(ByVal code As ErrorCode) As String
    ' 本来は外部ファイルやデータベースからエラーメッセージを取得するロジックを実装
    ' ここでは簡略化のため、Select Case 文でメッセージを返す
    Select Case code
        Case ERR_FILEIO_NOT_FOUND
            GetErrorMessageFromResource = "ファイルが見つかりません。"
        Case ERR_INPUT_INVALID
            GetErrorMessageFromResource = "無効な入力です。"
        Case ERR_DATABASE_CONNECTION_FAILED
            GetErrorMessageFromResource = "データベースに接続できません。"
        Case ERR_NETWORK_TIMEOUT
            GetErrorMessageFromResource = "タイムアウトが発生しました。"
        Case ERR_SYSTEM_OUT_OF_MEMORY
            GetErrorMessageFromResource = "メモリ不足です。"
        Case ERR_FILEIO_INVALID_FORMAT
            GetErrorMessageFromResource = "ファイル形式が正しくありません。"
        Case ERR_FILEIO_ACCESS_DENIED
            GetErrorMessageFromResource = "アクセスが拒否されました。"
        Case ERR_NETWORK_ERROR
            GetErrorMessageFromResource = "ネットワークエラーが発生しました。"
        Case ERR_DATABASE_QUERY_FAILED
            GetErrorMessageFromResource = "データベースクエリの実行に失敗しました。"
        Case ERR_UNEXPECTED
            GetErrorMessageFromResource = "予期しないエラーが発生しました。"
        Case Else
            GetErrorMessageFromResource = "不明なエラーです。(Code:" & code & ")"
    End Select
End Function

'---------------------------------
' Strategyパターン実装: 各コード毎にハンドラを返す
'---------------------------------
Public Function GetErrorHandler(ByVal code As ErrorCode) As IErrorHandler

    ' 各エラーハンドラのインスタンスを保持する Dictionary
    Static errorHandlers As Object ' Scripting.Dictionary

    If errorHandlers Is Nothing Then
        Set errorHandlers = CreateObject("Scripting.Dictionary")
    End If

    ' エラーコードに対応するエラーハンドラが既に存在するか確認
    If Not errorHandlers.Exists(code) Then
        ' 存在しない場合は、新しいエラーハンドラを生成して Dictionary に追加
        Select Case code
            Case ERR_FILEIO_NOT_FOUND
                Set errorHandlers(code) = FileNotFoundErrorHandlerFactory.Create
            Case ERR_INPUT_INVALID
                Set errorHandlers(code) = InvalidInputErrorHandlerFactory.Create
            Case ERR_DATABASE_CONNECTION_FAILED
                Set errorHandlers(code) = DatabaseConnectionErrorHandlerFactory.Create
            Case Else
                Set errorHandlers(code) = DefaultErrorHandlerFactory.Create
        End Select
    End If

    ' エラーコードに対応するエラーハンドラを返す
    Set GetErrorHandler = errorHandlers(code)

End Function

'---------------------------------
' メインのエラーハンドル呼び出し
'---------------------------------
Public Function HandleError(ByVal errorDetail As typErrorDetail) As Boolean

    ' --- スタックトレース情報は呼び出し元で設定 ---
    ' エラー発生箇所により近い位置で取得する方が正確な情報を得られるため

    Dim handler As IErrorHandler
    Set handler = GetErrorHandler(errorDetail.ErrorCode)

    ' エラーハンドラによる処理の結果を取得
    Dim continueProcessing As Boolean
    continueProcessing = handler.HandleError(errorDetail)

    ' エラーハンドラの結果に基づいて、後続の処理を決定
    If continueProcessing Then
        ' 処理を継続
        HandleError = True
    Else
        ' 処理を中断 (上位層にエラーを伝播)
        HandleError = False
        ' 必要に応じて、ここで上位層にエラーを再スローすることも検討
        ' 例: RaiseError errorDetail (エラーオブジェクトを上位に伝播させるカスタム関数)
    End If

End Function

'---------------------------------
' エラー処理クラスで共通利用するHelper (カスタムエラーオブジェクト再発生用)
'---------------------------------
Public Sub RaiseError(ByVal errorDetail As typErrorDetail)
    ' カスタムエラーオブジェクトを上位のレイヤーに伝播させる
    ' この例では、エラーコードとエラーメッセージを結合して上位にスロー
    Err.Raise errorDetail.ErrorCode, MODULE_NAME, errorDetail.Description
End Sub


' ======================
' 1.5.x エラー処理用インターフェースクラス(IErrorHandler.cls)
' ======================
' これは「クラスモジュール」としてプロジェクトに追加
' 名称は「IErrorHandler」などにする。VBAでは純粋インターフェースを直接表現しにくいため、
' メソッドのシグネチャのみ定義したクラスを用意し、実装側がこれを「Implements」する形。
Option Explicit

Public Function HandleError(ByVal errorDetail As typErrorDetail) As Boolean
    ' インターフェース用の宣言のみ。実装は各クラスへ委譲。
End Function


' ======================
' 1.5.x 個別エラーハンドラの例 (FileNotFoundErrorHandler.cls)
' ======================
' ファイルが見つからない場合のハンドリング
Option Explicit

Implements IErrorHandler

'---------------------------------
' シングルトンインスタンス
'---------------------------------
' エラーハンドラはステートレスな処理を行うため、Singleton パターンを採用
Private Shared instance As FileNotFoundErrorHandler

'---------------------------------
' ファクトリ
'---------------------------------
Public Function Create() As FileNotFoundErrorHandler
    If instance Is Nothing Then
        Set instance = New FileNotFoundErrorHandler
    End If
    Set Create = instance
End Function

'---------------------------------
' IErrorHandler インターフェースの実装
'---------------------------------
Private Function IErrorHandler_HandleError(ByVal errorDetail As typErrorDetail) As Boolean
    ' ファイルが見つからない場合、ここで必要な処理を行う
    ' 例: デフォルトファイルパスに切り替える、メッセージ表示で終了、など
    Dim proceed As Boolean
    proceed = True ' ファイルがなくても代替処理を進めるなら True

    ' ユーザー通知 (modError の Notifier を使用)
    ' エラーメッセージは modError の GetErrorMessage で取得
    modError.Notifier.Notify errorDetail, vbExclamation

    ' ログに出力 (modError の Logger を使用)
    modError.LogError errorDetail

    ' 必要なら再発生
    ' modError.RaiseError errorDetail ' カスタムエラーオブジェクトを上位に伝播させる

    IErrorHandler_HandleError = proceed
End Function


' ======================
' 1.5.x 個別エラーハンドラの例 (InvalidInputErrorHandler.cls)
' ======================
' 無効な入力エラー時のハンドリング
Option Explicit

Implements IErrorHandler

'---------------------------------
' シングルトンインスタンス
'---------------------------------
' エラーハンドラはステートレスな処理を行うため、Singleton パターンを採用
Private Shared instance As InvalidInputErrorHandler

'---------------------------------
' ファクトリ
'---------------------------------
Public Function Create() As InvalidInputErrorHandler
    If instance Is Nothing Then
        Set instance = New InvalidInputErrorHandler
    End If
    Set Create = instance
End Function

'---------------------------------
' IErrorHandler インターフェースの実装
'---------------------------------
Private Function IErrorHandler_HandleError(ByVal errorDetail As typErrorDetail) As Boolean
    ' 無効な入力の場合の処理
    ' 例: ユーザーに再入力を促す、処理を中断する etc.
    Dim proceed As Boolean
    proceed = False ' 入力エラーは処理失敗とする

    ' ユーザー通知 (modError の Notifier を使用)
    ' エラーメッセージは modError の GetErrorMessage で取得
    modError.Notifier.Notify errorDetail, vbExclamation

    ' ログ出力 (modError の Logger を使用)
    modError.LogError errorDetail

    IErrorHandler_HandleError = proceed
End Function


' ======================
' 1.5.x 個別エラーハンドラの例 (DatabaseConnectionErrorHandler.cls)
' ======================
Option Explicit

Implements IErrorHandler

' modCommon から最大リトライ回数を取得
Private Const MAX_RETRY_COUNT As Long = MAX_RETRY_COUNT ' 最大リトライ回数

'---------------------------------
' シングルトンインスタンス
'---------------------------------
' エラーハンドラはステートレスな処理を行うため、Singleton パターンを採用
Private Shared instance As DatabaseConnectionErrorHandler

'---------------------------------
' ファクトリ
'---------------------------------
Public Function Create() As DatabaseConnectionErrorHandler
    If instance Is Nothing Then
        Set instance = New DatabaseConnectionErrorHandler
    End If
    Set Create = instance
End Function

'---------------------------------
' IErrorHandler インターフェースの実装
'---------------------------------
Private Function IErrorHandler_HandleError(ByVal errorDetail As typErrorDetail) As Boolean
    ' データベース接続エラー時の処理
    Dim retryCount As Integer
    Dim success As Boolean

    For retryCount = 1 To MAX_RETRY_COUNT
        If TryDatabaseConnection() Then
            success = True
            Exit For
        End If
        ' リトライ間隔を設ける場合は、ここに Wait 処理などを追加
        ' 例: Application.Wait (Now + TimeValue("0:00:01")) ' 1秒待機
    Next retryCount

    If Not success Then
        ' ユーザー通知 (modError の Notifier を使用)
        ' エラーメッセージは modError の GetErrorMessage で取得
        modError.Notifier.Notify errorDetail, vbCritical

        ' ログ出力 (modError の Logger を使用)
        modError.LogError errorDetail
    End If

    IErrorHandler_HandleError = success
End Function

'---------------------------------
' データベース接続試行
'---------------------------------
Private Function TryDatabaseConnection() As Boolean
    On Error Resume Next
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    ' 接続文字列は環境に合わせて変更、または modConfig から取得
    ' conn.Open "Provider=SQLOLEDB;Data Source=myServerAddress;Initial Catalog=myDataBase;User Id=myUsername;Password=myPassword;"
    conn.Open gAppSettings.DatabaseConnectionString
    If Err.Number = 0 Then
        conn.Close
        TryDatabaseConnection = True
    End If
    Set conn = Nothing
    On Error GoTo 0
End Function


' ======================
' 1.5.x 個別エラーハンドラの例 (DefaultErrorHandler.cls)
' ======================
' 上記以外のエラー汎用ハンドリング
Option Explicit

Implements IErrorHandler

'---------------------------------
' シングルトンインスタンス
'---------------------------------
' エラーハンドラはステートレスな処理を行うため、Singleton パターンを採用
Private Shared instance As DefaultErrorHandler

'---------------------------------
' ファクトリ
'---------------------------------
Public Function Create() As DefaultErrorHandler
    If instance Is Nothing Then
        Set instance = New DefaultErrorHandler
    End If
    Set Create = instance
End Function

'---------------------------------
' IErrorHandler インターフェースの実装
'---------------------------------
Private Function IErrorHandler_HandleError(ByVal errorDetail As typErrorDetail) As Boolean
    ' 汎用エラー時の処理
    ' ユーザー通知 & ログを出力して、エラーを上位に伝播させる例

    ' ユーザー通知 (modError の Notifier を使用)
    ' エラーメッセージは modError の GetErrorMessage で取得
    modError.Notifier.Notify errorDetail, vbCritical

    ' ログ出力 (modError の Logger を使用)
    modError.LogError errorDetail

    ' エラーを上位に伝播させる
    modError.RaiseError errorDetail

    IErrorHandler_HandleError = False
End Function


' ======================
' IUserNotifier インターフェース (IUserNotifier.cls)
' ======================
Option Explicit

'---------------------------------
' エラー情報をユーザーに通知する
'---------------------------------
Public Sub Notify(ByVal errorDetail As typErrorDetail, _
                    Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
                    Optional ByVal title As String = "エラー")
End Sub


' ======================
' MsgBox による通知クラス (MsgBoxNotifier.cls)
' ======================
Option Explicit

Implements IUserNotifier

'---------------------------------
' IUserNotifier インターフェースの実装
'---------------------------------
Private Sub IUserNotifier_Notify(ByVal errorDetail As typErrorDetail, _
                                    Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
                                    Optional ByVal title As String = "エラー")

    Dim msg As String

    msg = modError.GetErrorMessage(errorDetail.ErrorCode)

    If errorDetail.Description <> "" Then
        msg = msg & vbCrLf & errorDetail.Description
    End If

    MsgBox msg, buttons, title

End Sub


' ======================
' ステータスバーによる通知クラス (StatusBarNotifier.cls)
' ======================
Option Explicit

Implements IUserNotifier

'---------------------------------
' IUserNotifier インターフェースの実装
'---------------------------------
Private Sub IUserNotifier_Notify(ByVal errorDetail As typErrorDetail, _
                                    Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
                                    Optional ByVal title As String = "エラー")

    Application.StatusBar = "エラー: " & modError.GetErrorMessage(errorDetail.ErrorCode) & _
                            " (" & errorDetail.Description & ")"

End Sub


' ======================
' ILogger インターフェース (ILogger.cls)
' ======================
Option Explicit

'---------------------------------
' ログを記録する
'---------------------------------
' 引数を変更: エラーコード、発生時刻を追加
Public Sub Log(ByVal moduleName As String, ByVal message As String, ByVal errorCode As ErrorCode)
End Sub


' ======================
' ファイルへのログ出力クラス (FileLogger.cls)
' ======================
Option Explicit

Implements ILogger

' modCommon からログファイルのパスを取得
Private Property Get LOG_FILE_PATH() As String
    LOG_FILE_PATH = App.Path & "\" & DEFAULT_LOG_FILE
End Property

'---------------------------------
' ILogger インターフェースの実装
'---------------------------------
Private Sub ILogger_Log(ByVal moduleName As String, ByVal message As String, ByVal errorCode As ErrorCode)

    Dim fileNum As Integer
    fileNum = FreeFile

    On Error Resume Next ' ログファイルが存在しない場合など
    Open LOG_FILE_PATH For Append As #fileNum
    If Err.Number <> 0 Then
        ' ログファイルに書き込めない場合は、エラーを上位に伝播させるか、代替手段を検討
        ' 例: EventLog に書き込む、別の通知方法でユーザーに通知するなど
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Print #fileNum, Format(Now, "yyyy/MM/dd HH:mm:ss") & " [" & moduleName & "] [Code: " & errorCode & "] " & message
    Close #fileNum

End Sub


' ======================
' データベースへのログ出力クラス (DatabaseLogger.cls)
' ======================
Option Explicit

Implements ILogger

'---------------------------------
' ILogger インターフェースの実装
'---------------------------------
Private Sub ILogger_Log(ByVal moduleName As String, ByVal message As String, ByVal errorCode As ErrorCode)

    On Error Resume Next

    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")

    ' 接続文字列は環境に合わせて変更、または modConfig から取得
    conn.Open gAppSettings.DatabaseConnectionString

    If Err.Number <> 0 Then
        ' データベースに接続できない場合は、エラーを上位に伝播させるか、代替手段を検討
        On Error GoTo 0
        Exit Sub
    End If

    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn

    cmd.CommandText = "INSERT INTO ErrorLog (Timestamp, ModuleName, ErrorCode, Message) VALUES (?, ?, ?, ?)"
    cmd.Parameters.Append cmd.CreateParameter("@Timestamp", 7, 1, , Now) ' adDBTimeStamp
    cmd.Parameters.Append cmd.CreateParameter("@ModuleName", 200, 1, 255, moduleName) ' adVarChar
    cmd.Parameters.Append cmd.CreateParameter("@ErrorCode", 3, 1, , errorCode) ' adInteger
    cmd.Parameters.Append cmd.CreateParameter("@Message", 201, 1, -1, message) ' adLongVarChar (-1 は最大サイズ)

    cmd.Execute

    conn.Close
    Set conn = Nothing

    On Error GoTo 0

End Sub


' ======================
' 1.6 modFileIO (ファイル入出力)
' ======================
Option Explicit

' --- 列挙型 ---
' --- ファイルエンコーディング ---
Public Enum FileEncoding
    FE_UTF8 = 0
    FE_SHIFT_JIS = 1
    FE_UTF16_LE = 2 ' UTF-16 Little Endian
    FE_UTF16_BE = 3 ' UTF-16 Big Endian
    ' ... 他のエンコーディングを追加 ...
End Enum

' --- 関数 ---
' --- テキストファイル読み込み ---
'   パラメータ:
'       filePath: ファイルパス
'       encoding: エンコーディング (オプション、デフォルトは UTF-8)
'       streamMode: ストリームモード (オプション、デフォルトは False)
'   戻り値:
'       ファイルの内容 (文字列)
Public Function FileReadText(ByVal filePath As String, _
                            Optional ByVal encoding As FileEncoding = FE_UTF8, _
                            Optional ByVal streamMode As Boolean = False) As String
    Dim fileNum As Integer
    On Error GoTo ErrorHandler
    
    If Not FileExists(filePath) Then
        Dim errDetail As typErrorDetail
        errDetail.ErrorCode = ERR_FILEIO_NOT_FOUND
        errDetail.Description = "File not found: " & filePath
        errDetail.Source = "modFileIO"
        errDetail.ProcedureName = "FileReadText"
        errDetail.StackTrace = "" ' 必要なら独自スタックトレースを取得
        errDetail.OccurredAt = Now
        Err.Clear
        HandleError errDetail
        Exit Function ' エラー時は空文字列を返す
    End If
        
    If streamMode Then
        ' ストリーミング処理
        FileReadText = FileReadTextStream(filePath, encoding)
    Else
        ' 従来の一括読み込み
        fileNum = FreeFile
        Open filePath For Input As #fileNum Encoding GetEncodingString(encoding)
            FileReadText = Input$(LOF(fileNum), fileNum)
        Close #fileNum
    End If
    
    Exit Function
    
Cleanup:
    ' リソース解放 (万が一Open済ならClose)
    If fileNum <> 0 Then
        Close #fileNum
    End If
    Exit Function
    
ErrorHandler:
    Dim errDetail As typErrorDetail
    errDetail.ErrorCode = GetFileIOErrorCode(Err.Number)
    errDetail.Description = Err.Description
    errDetail.Source = "modFileIO"
    errDetail.ProcedureName = "FileReadText"
    errDetail.StackTrace = ""
    errDetail.OccurredAt = Now
    
    Err.Clear
    HandleError errDetail
    GoTo Cleanup
End Function

' --- テキストファイルストリーム読み込み ---
'   パラメータ:
'       filePath: ファイルパス
'       encoding: エンコーディング (オプション、デフォルトは UTF-8)
'   戻り値:
'       ファイルの内容 (文字列)
Private Function FileReadTextStream(ByVal filePath As String, _
                                    Optional ByVal encoding As FileEncoding = FE_UTF8) As String
    On Error GoTo ErrorHandler
    
    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    With objStream
        .Type = 2 ' adTypeText
        .Charset = GetEncodingString(encoding)
        .Open
        .LoadFromFile filePath
        FileReadTextStream = .ReadText()
        .Close
    End With
    Set objStream = Nothing
    Exit Function
    
Cleanup:
    ' クローズ処理 (ADODB.StreamはWithでCloseしている想定)
    If Not objStream Is Nothing Then
        On Error Resume Next
        objStream.Close
        Set objStream = Nothing
    End If
    Exit Function
    
ErrorHandler:
    Dim errDetail As typErrorDetail
    errDetail.ErrorCode = GetFileIOErrorCode(Err.Number)
    errDetail.Description = Err.Description
    errDetail.Source = "modFileIO"
    errDetail.ProcedureName = "FileReadTextStream"
    errDetail.StackTrace = ""
    errDetail.OccurredAt = Now
    
    Err.Clear
    HandleError errDetail
    GoTo Cleanup
End Function

' --- テキストファイル書き込み ---
'   パラメータ:
'       filePath: ファイルパス
'       content: 書き込む内容
'       encoding: エンコーディング (オプション、デフォルトは UTF-8)
'       append: 追記モード (オプション、デフォルトは True)
'   戻り値:
'       書き込み成功 (True/False)
Public Function FileWriteText(ByVal filePath As String, ByVal content As String, _
                              Optional ByVal encoding As FileEncoding = FE_UTF8, _
                              Optional ByVal append As Boolean = True) As Boolean
    Dim fileNum As Integer
    On Error GoTo ErrorHandler
    
    fileNum = FreeFile
    If append Then
        Open filePath For Append As #fileNum Encoding GetEncodingString(encoding)
    Else
        Open filePath For Output As #fileNum Encoding GetEncodingString(encoding)
    End If
    
    Print #fileNum, content
    Close #fileNum
    FileWriteText = True
    
    Exit Function
    
Cleanup:
    If fileNum <> 0 Then Close #fileNum
    Exit Function
    
ErrorHandler:
    FileWriteText = False
    
    Dim errDetail As typErrorDetail
    errDetail.ErrorCode = GetFileIOErrorCode(Err.Number)
    errDetail.Description = Err.Description
    errDetail.Source = "modFileIO"
    errDetail.ProcedureName = "FileWriteText"
    errDetail.StackTrace = ""
    errDetail.OccurredAt = Now
    
    Err.Clear
    HandleError errDetail
    GoTo Cleanup
End Function

' --- バイナリファイル読み込み ---
'   パラメータ:
'       filePath: ファイルパス
'   戻り値:
'       ファイルの内容 (バイト配列)
Public Function FileReadBinary(ByVal filePath As String) As Byte()
    Dim fileNum As Integer
    On Error GoTo ErrorHandler
    
    If Not FileExists(filePath) Then
        Dim errDetail As typErrorDetail
        errDetail.ErrorCode = ERR_FILEIO_NOT_FOUND
        errDetail.Description = "File not found: " & filePath
        errDetail.Source = "modFileIO"
        errDetail.ProcedureName = "FileReadBinary"
        errDetail.StackTrace = ""
        errDetail.OccurredAt = Now
        Err.Clear
        HandleError errDetail
        Exit Function ' エラー時は空の配列を返す
    End If

    fileNum = FreeFile
    Open filePath For Binary Access Read As #fileNum
        Dim fileData() As Byte
        ReDim fileData(LOF(fileNum) - 1)
        Get #fileNum, , fileData
    Close #fileNum
    FileReadBinary = fileData
    Exit Function
    
Cleanup:
    If fileNum <> 0 Then Close #fileNum
    Exit Function
    
ErrorHandler:
    Dim errDetail As typErrorDetail
    errDetail.ErrorCode = GetFileIOErrorCode(Err.Number)
    errDetail.Description = Err.Description
    errDetail.Source = "modFileIO"
    errDetail.ProcedureName = "FileReadBinary"
    errDetail.StackTrace = ""
    errDetail.OccurredAt = Now
    
    Err.Clear
    HandleError errDetail
    GoTo Cleanup
End Function

' --- バイナリファイル書き込み ---
'   パラメータ:
'       filePath: ファイルパス
'       data: 書き込むデータ (バイト配列)
'   戻り値:
'       書き込み成功 (True/False)
Public Function FileWriteBinary(ByVal filePath As String, ByVal data() As Byte) As Boolean
    Dim fileNum As Integer
    On Error GoTo ErrorHandler
    
    fileNum = FreeFile
    Open filePath For Binary Access Write As #fileNum
        Put #fileNum, , data
    Close #fileNum
    FileWriteBinary = True
    Exit Function
    
Cleanup:
    If fileNum <> 0 Then Close #fileNum
    Exit Function
    
ErrorHandler:
    FileWriteBinary = False
    
    Dim errDetail As typErrorDetail
    errDetail.ErrorCode = GetFileIOErrorCode(Err.Number)
    errDetail.Description = Err.Description
    errDetail.Source = "modFileIO"
    errDetail.ProcedureName = "FileWriteBinary"
    errDetail.StackTrace = ""
    errDetail.OccurredAt = Now
    
    Err.Clear
    HandleError errDetail
    GoTo Cleanup
End Function

' --- ファイル存在確認 ---
'   パラメータ:
'       filePath: ファイルパス
'   戻り値:
'       ファイルが存在する (True/False)
'       エラーが発生した場合、エラーコードを返す
Public Function FileExists(ByVal filePath As String) As Variant
    On Error GoTo ErrorHandler
    FileExists = (Dir(filePath) <> "")
    Exit Function
    
ErrorHandler:
    FileExists = GetFileIOErrorCode(Err.Number)
    Exit Function
End Function

' --- フォルダ存在確認 ---
'   パラメータ:
'       folderPath: フォルダパス
'   戻り値:
'       フォルダが存在する (True/False)
'       エラーが発生した場合、エラーコードを返す
Public Function FolderExists(ByVal folderPath As String) As Variant
    On Error GoTo ErrorHandler
    FolderExists = (Dir(folderPath, vbDirectory) <> "")
    Exit Function
    
ErrorHandler:
    FolderExists = GetFileIOErrorCode(Err.Number)
    Exit Function
End Function

' --- 絶対パス取得 ---
'   パラメータ:
'       relativePath: 相対パス
'       basePath: ベースパス (オプション)
'   戻り値:
'       絶対パス
Public Function GetAbsolutePath(ByVal relativePath As String, Optional ByVal basePath As String) As Variant
    On Error GoTo ErrorHandler
    If IsMissing(basePath) Then
        basePath = CurDir
    End If
    GetAbsolutePath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(basePath & "\" & relativePath)
    Exit Function
    
ErrorHandler:
    GetAbsolutePath = GetFileIOErrorCode(Err.Number)
    Exit Function
End Function

' --- ファイル名取得 ---
'   パラメータ:
'       filePath: ファイルパス
'   戻り値:
'       ファイル名
Public Function GetFileName(ByVal filePath As String) As String
    GetFileName = CreateObject("Scripting.FileSystemObject").GetFileName(filePath)
End Function

' --- ファイル拡張子取得 ---
'   パラメータ:
'       filePath: ファイルパス
'   戻り値:
'       ファイル拡張子
Public Function GetFileExtension(ByVal filePath As String) As String
    GetFileExtension = CreateObject("Scripting.FileSystemObject").GetExtensionName(filePath)
End Function

' --- フォルダ作成 ---
'   パラメータ:
'       folderPath: フォルダパス
'   戻り値:
'       作成成功 (True/False)
Public Function CreateFolder(ByVal folderPath As String) As Boolean
    On Error Resume Next
    MkDir folderPath
    CreateFolder = (Err.Number = 0)
    On Error GoTo 0
End Function

' --- ファイル削除 ---
'   パラメータ:
'       filePath: ファイルパス
'   戻り値:
'       削除成功 (True/False)
Public Function DeleteFile(ByVal filePath As String) As Boolean
    On Error Resume Next
    Kill filePath
    DeleteFile = (Err.Number = 0)
    On Error GoTo 0
End Function

' --- フォルダ削除 ---
'   パラメータ:
'       folderPath: フォルダパス
'   戻り値:
'       削除成功 (True/False)
Public Function DeleteFolder(ByVal folderPath As String) As Boolean
    On Error Resume Next
    RmDir folderPath
    DeleteFolder = (Err.Number = 0)
    On Error GoTo 0
End Function

' --- エンコーディング文字列取得 ---
'   パラメータ:
'       encoding: エンコーディング
'   戻り値:
'       エンコーディング文字列
Private Function GetEncodingString(ByVal encoding As FileEncoding) As String
    Select Case encoding
        Case FE_UTF8
            GetEncodingString = "UTF-8"
        Case FE_SHIFT_JIS
            GetEncodingString = "Shift_JIS"
        Case FE_UTF16_LE
            GetEncodingString = "UTF-16LE"
        Case FE_UTF16_BE
            GetEncodingString = "UTF-16BE"
        Case Else
            GetEncodingString = "UTF-8" ' デフォルトは UTF-8
    End Select
End Function

' --- ファイルIOエラーコード取得 ---
'   パラメータ:
'       errNumber: Err.Number
'   戻り値:
'       ファイルIO関連のエラーコード
Private Function GetFileIOErrorCode(ByVal errNumber As Long) As ErrorCode
    Select Case errNumber
        Case 53 ' File not found
            GetFileIOErrorCode = ERR_FILEIO_NOT_FOUND
        Case 70 ' Permission denied
            GetFileIOErrorCode = ERR_FILEIO_ACCESS_DENIED
        Case 75, 76 ' Path/File access error, Path not found
            GetFileIOErrorCode = ERR_FILEIO_ACCESS_DENIED
        Case 55 ' File already open
            GetFileIOErrorCode = ERR_FILEIO_ACCESS_DENIED
        Case Else
            GetFileIOErrorCode = ERR_UNEXPECTED
    End Select
End Function

' --- ファイルシステム操作抽象化クラス (clsFileSystem) ---
' クラスモジュールとして以下を新規作成し、modFileIOから関連する関数を移動

' ======================
' clsFileSystem (ファイルシステム操作)
' ======================
' Option Explicit
'
' ' --- ファイル存在確認 ---
' Public Function FileExists(ByVal filePath As String) As Variant
'     On Error GoTo ErrorHandler
'     FileExists = (Dir(filePath) <> "")
'     Exit Function
'     
' ErrorHandler:
'     FileExists = GetFileIOErrorCode(Err.Number)
'     Exit Function
' End Function
'
' ' --- フォルダ存在確認 ---
' Public Function FolderExists(ByVal folderPath As String) As Variant
'     On Error GoTo ErrorHandler
'     FolderExists = (Dir(folderPath, vbDirectory) <> "")
'     Exit Function
'     
' ErrorHandler:
'     FolderExists = GetFileIOErrorCode(Err.Number)
'     Exit Function
' End Function
'
' ' --- 絶対パス取得 ---
' Public Function GetAbsolutePath(ByVal relativePath As String, Optional ByVal basePath As String) As Variant
'     On Error GoTo ErrorHandler
'     If IsMissing(basePath) Then
'         basePath = CurDir
'     End If
'     GetAbsolutePath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(basePath & "\" & relativePath)
'     Exit Function
'     
' ErrorHandler:
'     GetAbsolutePath = GetFileIOErrorCode(Err.Number)
'     Exit Function
' End Function
'
' ' --- ファイル名取得 ---
' Public Function GetFileName(ByVal filePath As String) As String
'     GetFileName = CreateObject("Scripting.FileSystemObject").GetFileName(filePath)
' End Function
'
' ' --- ファイル拡張子取得 ---
' Public Function GetFileExtension(ByVal filePath As String) As String
'     GetFileExtension = CreateObject("Scripting.FileSystemObject").GetExtensionName(filePath)
' End Function
'
' ' --- フォルダ作成 ---
' Public Function CreateFolder(ByVal folderPath As String) As Boolean
'     On Error Resume Next
'     MkDir folderPath
'     CreateFolder = (Err.Number = 0)
'     On Error GoTo 0
' End Function
'
' ' --- ファイル削除 ---
' Public Function DeleteFile(ByVal filePath As String) As Boolean
'     On Error Resume Next
'     Kill filePath
'     DeleteFile = (Err.Number = 0)
'     On Error GoTo 0
' End Function
'
' ' --- フォルダ削除 ---
' Public Function DeleteFolder(ByVal folderPath As String) As Boolean
'     On Error Resume Next
'     RmDir folderPath
'     DeleteFolder = (Err.Number = 0)
'     On Error GoTo 0
' End Function
'
' ' --- ファイルIOエラーコード取得 ---
' Private Function GetFileIOErrorCode(ByVal errNumber As Long) As ErrorCode
'     Select Case errNumber
'         Case 53 ' File not found
'             GetFileIOErrorCode = ERR_FILEIO_NOT_FOUND
'         Case 70 ' Permission denied
'             GetFileIOErrorCode = ERR_FILEIO_ACCESS_DENIED
'         Case 75, 76 ' Path/File access error, Path not found
'             GetFileIOErrorCode = ERR_FILEIO_ACCESS_DENIED
'         Case 55 ' File already open
'             GetFileIOErrorCode = ERR_FILEIO_ACCESS_DENIED
'         Case Else
'             GetFileIOErrorCode = ERR_UNEXPECTED
'     End Select
' End Function


' ======================
' 1.7 modUtility (その他の汎用関数)
' ======================
Option Explicit

' 安全な除算関数
Public Function SafeDivide(ByVal numerator As Double, ByVal denominator As Double, ByVal defaultValue As Variant) As Variant
    On Error GoTo ErrorHandler
    If denominator = 0 Then
        SafeDivide = defaultValue
    Else
        SafeDivide = numerator / denominator
    End If
    Exit Function

ErrorHandler:
    SafeDivide = CVErr(xlErrDiv0)
End Function

' 日付の妥当性チェック関数
Public Function IsValidDate(ByVal testDate As Variant) As Boolean
    IsValidDate = VBA.IsDate(testDate)
End Function

' メールアドレスの妥当性チェック関数
Public Function IsValidEmail(ByVal email As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$"
        .IgnoreCase = True
    End With
    IsValidEmail = regex.Test(email)
    Set regex = Nothing
End Function

' カレントディレクトリ取得関数
Public Function GetWorkingDir() As String
    GetWorkingDir = CurDir
End Function

' 文字列の左パディング関数
Public Function PadLeft(ByVal baseStr As String, ByVal totalWidth As Integer, Optional ByVal padChar As String = " ") As String
    If Len(baseStr) >= totalWidth Then
        PadLeft = baseStr
    Else
        PadLeft = String(totalWidth - Len(baseStr), padChar) & baseStr
    End If
End Function

' 文字列のトリム関数
Public Function TrimString(ByVal str As String) As String
    TrimString = Trim(str)
End Function

' 文字列の分割関数
Public Function SplitString(ByVal str As String, ByVal delimiter As String) As Variant
    SplitString = Split(str, delimiter)
End Function

' 日付加算関数
Public Function DateAdd(ByVal interval As String, ByVal number As Double, ByVal dateValue As Date) As Date
    DateAdd = VBA.DateAdd(interval, number, dateValue)
End Function

' 日付差分関数
Public Function DateDiff(ByVal interval As String, ByVal date1 As Date, ByVal date2 As Date, _
                            Optional ByVal firstDayOfWeek As VbDayOfWeek = vbSunday, _
                            Optional ByVal firstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1) As Long
    DateDiff = VBA.DateDiff(interval, date1, date2, firstDayOfWeek, firstWeekOfYear)
End Function

' 日付フォーマット関数
Public Function FormatDate(ByVal dateValue As Date, Optional ByVal format As String = "yyyy/MM/dd") As String
    FormatDate = VBA.Format$(dateValue, format)
End Function

' 型変換関数は、エラー処理を含めて呼び出し元で対応する方針に変更
' INIファイルから任意の設定を読み込む (引数からファイルパスを受け取るように変更)
Public Function ReadConfig(ByVal filePath As String, ByVal section As String, ByVal key As String, _
                           Optional ByVal defaultValue As String = "") As String
    Dim buffer As String
    Dim retVal As Long
    
    buffer = String$(255, 0)
    
    retVal = GetPrivateProfileString(section, key, defaultValue, buffer, Len(buffer), filePath)
    If retVal > 0 Then
        ReadConfig = Left$(buffer, retVal)
    Else
        ReadConfig = defaultValue
    End If
End Function

' DB接続文字列を動的に取得 (エラー時は呼び出し元に例外を投げる)
' 引数から設定ファイルのパスを受け取るように変更
Public Function GetDbConnectionString() As String
    Dim connectionString As String

    ' config.ini から接続文字列全体を読み込む
    ' connectionString = ReadConfig(configFilePath, "Database", "ConnectionString", "") ' ReadConfig関数はmodUtilityにあるが、引数でファイルパスを受け取るよう変更されている
    connectionString = gAppSettings.DatabaseConnectionString

    If connectionString = "" Then
        ' 接続文字列が設定されていない場合は例外を投げる
        Dim errDetail As typErrorDetail
        errDetail.ErrorCode = ERR_DATABASE_CONNECTION_FAILED
        errDetail.Description = "config.ini の [Database] セクションに ConnectionString が設定されていません。"
        errDetail.Source = "modUtility"
        errDetail.ProcedureName = "GetDbConnectionString"
        errDetail.StackTrace = ""
        errDetail.OccurredAt = Now

        Err.Raise errDetail.ErrorCode, errDetail.Source & ":" & errDetail.ProcedureName, errDetail.Description
    End If

    GetDbConnectionString = connectionString
End Function


' ======================
' 1.7 clsLogger (ログ出力 - クラスモジュール)
' ======================
Option Explicit

Public LogLevel As LogLevelEnum
Public LogDestination As LogDestinationEnum
Public LogFilePath As String
Public LogTableName As String
Public LogEventSource As String

Private m_adoConnection As Object
Private m_initialized As Boolean

Public Event Logged(ByVal logMessage As String, ByVal logLevel As LogLevelEnum)
' Loggedイベントは同期実行。非同期化は呼び出し側で対応可能 (コメントで明記)

'---------------------------------
' デバッグログ
'---------------------------------
Public Sub LogDebug(ByVal message As String)
    Call WriteLog(message, LOG_LEVEL_DEBUG)
End Sub

Public Sub LogInfo(ByVal message As String)
    Call WriteLog(message, LOG_LEVEL_INFO)
End Sub

Public Sub LogWarning(ByVal message As String)
    Call WriteLog(message, LOG_LEVEL_WARNING)
End Sub

Public Sub LogError(ByVal message As String)
    Call WriteLog(message, LOG_LEVEL_ERROR)
End Sub

Public Sub LogMessage(ByVal message As String, ByVal logLevel As LogLevelEnum)
    Call WriteLog(message, logLevel)
End Sub

'---------------------------------
' ログ設定
'---------------------------------
Public Sub Configure(ByVal logLevel As LogLevelEnum, ByVal logDestination As LogDestinationEnum, _
                     Optional ByVal filePath As String = "", _
                     Optional ByVal tableName As String = "", _
                     Optional ByVal eventSource As String = "")

    Me.LogLevel = logLevel
    Me.LogDestination = logDestination
    Me.LogFilePath = filePath
    Me.LogTableName = tableName
    Me.LogEventSource = eventSource
    m_initialized = True
End Sub

'---------------------------------
' ログ出力メイン
'---------------------------------
Private Sub WriteLog(ByVal message As String, ByVal logLevel As LogLevelEnum)
    If Not m_initialized Then
        ' 未設定ならデフォルトを適用
        Me.Configure LOG_LEVEL_INFO, LOG_DESTINATION_FILE, DEFAULT_LOG_FILE
    End If

    If logLevel >= Me.LogLevel Then
        Select Case Me.LogDestination
            Case LOG_DESTINATION_FILE
                WriteToFile message, logLevel
            Case LOG_DESTINATION_DATABASE
                WriteToDatabase message, logLevel
            Case LOG_DESTINATION_EVENTLOG
                WriteToEventLog message, logLevel
            Case LOG_DESTINATION_CONSOLE
                Debug.Print Format$(Now, DEFAULT_DATE_FORMAT) & " [" & GetLogLevelName(logLevel) & "] " & message
            Case LOG_DESTINATION_EMAIL
                ' 将来的にメール通知を実装
            Case Else
                ' 何もしない
        End Select

        RaiseEvent Logged(message, logLevel)
    End If
End Sub

Private Function GetLogLevelName(ByVal logLevel As LogLevelEnum) As String
    Select Case logLevel
        Case LOG_LEVEL_DEBUG:   GetLogLevelName = "DEBUG"
        Case LOG_LEVEL_INFO:    GetLogLevelName = "INFO"
        Case LOG_LEVEL_WARNING: GetLogLevelName = "WARNING"
        Case LOG_LEVEL_ERROR:   GetLogLevelName = "ERROR"
        Case LOG_LEVEL_FATAL:   GetLogLevelName = "FATAL"
        Case Else:              GetLogLevelName = "UNKNOWN"
    End Select
End Function

Private Sub WriteToFile(ByVal message As String, ByVal logLevel As LogLevelEnum)
    On Error GoTo ErrorHandler

    Dim fileNum As Integer
    fileNum = FreeFile
    Open Me.LogFilePath For Append As #fileNum
    Print #fileNum, Format$(Now, DEFAULT_DATE_FORMAT) & " [" & GetLogLevelName(logLevel) & "] " & message
    Close #fileNum
    Exit Sub

ErrorHandler:
    ' ファイルエラーは無視するか、あるいは別途通知する
    Debug.Print "Error writing to log file: " & Err.Description
End Sub

Private Sub WriteToDatabase(ByVal message As String, ByVal logLevel As LogLevelEnum)
    On Error GoTo ErrorHandler

    If m_adoConnection Is Nothing Then
        Set m_adoConnection = ADOConnection
    End If

    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = m_adoConnection
        .CommandText = "INSERT INTO " & Me.LogTableName & " (LogTime, LogLevel, LogMessage) VALUES (?, ?, ?)"
        .Parameters.Append .CreateParameter("LogTime", 7, 1, , Now)          ' adDate=7, adParamInput=1
        .Parameters.Append .CreateParameter("LogLevel", 3, 1, , logLevel)   ' adInteger=3
        .Parameters.Append .CreateParameter("LogMessage", 200, 1, 4000, message) ' adVarChar=200
        .Execute
    End With
    Exit Sub

ErrorHandler:
    Debug.Print "Error writing to database: " & Err.Description
End Sub

Private Sub WriteToEventLog(ByVal message As String, ByVal logLevel As LogLevelEnum)
    On Error Resume Next

    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    Dim eventType As Integer

    Select Case logLevel
        Case LOG_LEVEL_ERROR, LOG_LEVEL_FATAL
            eventType = 1 ' ERROR
        Case LOG_LEVEL_WARNING
            eventType = 2 ' WARNING
        Case Else
            eventType = 4 ' INFORMATION
    End Select

    objShell.LogEvent eventType, message, Me.LogEventSource
    Set objShell = Nothing
    On Error GoTo 0
End Sub

'---------------------------------
' ADOConnectionプロパティ
'---------------------------------
Public Property Get ADOConnection() As Object
    If m_adoConnection Is Nothing Then
        Set m_adoConnection = CreateObject("ADODB.Connection")
        m_adoConnection.ConnectionString = GetDbConnectionString()
        On Error Resume Next
        m_adoConnection.Open
        If Err.Number <> 0 Then
            Debug.Print "Error opening DB connection: " & Err.Description
            ' 必要に応じてエラー通知
            Dim errDetail As typErrorDetail
            errDetail.ErrorCode = ERR_DATABASE_CONNECTION
            errDetail.Description = "DB接続に失敗: " & Err.Description
            errDetail.Source = "clsLogger"
            errDetail.ProcedureName = "ADOConnection"
            errDetail.StackTrace = ""
            errDetail.OccurredAt = Now

            Err.Clear
            HandleError errDetail
        End If
        On Error GoTo 0
    End If
    Set ADOConnection = m_adoConnection
End Property

Public Property Let ADOConnection(ByVal adoConnection As Object)
    Set m_adoConnection = adoConnection
End Property

Private Sub Class_Initialize()
    Me.LogLevel = LOG_LEVEL_INFO
    Me.LogDestination = LOG_DESTINATION_FILE
    Me.LogFilePath = DEFAULT_LOG_FILE
    Me.LogTableName = "AppLog"
    Me.LogEventSource = APPLICATION_NAME
    m_initialized = True
End Sub

Private Sub Class_Terminate()
    If Not m_adoConnection Is Nothing Then
        If m_adoConnection.State = 1 Then ' adStateOpen=1
            m_adoConnection.Close
        End If
        Set m_adoConnection = Nothing
    End If
End Sub


' ======================
' 1.8 clsLogger (ログ出力 - クラスモジュール)
' ======================
Option Explicit

' ログレベル
' modCommon の LogLevelEnum を使用

' ログ出力先インターフェース
Private logDestination As ILogDestination

' ログ出力先
' modCommon の LogDestinationEnum は使用しない

' ログファイルパス
Private logFilePath As String

' ログテーブル名
Private logTableName As String

' ログイベントソース
Private logEventSource As String

' ADOコネクションオブジェクト
Private adoConnection As Object

' 初期化フラグ
Private initialized As Boolean

' ログメッセージキュー
Private logQueue As Queue

' タイマー間隔（ミリ秒）
Private timerInterval As Long

' タイマーID
Private timerID As LongPtr

' ログ設定
Private Type typLoggerSettings
    LogLevel As LogLevelEnum
    LogDestination As ILogDestination ' インターフェース型に変更
    LogFilePath As String
    LogTableName As String
    LogEventSource As String
    TimerInterval As Long ' 非同期処理のタイマー間隔
End Type

Private loggerSettings As typLoggerSettings

' Loggedイベント (非同期イベント)
Public Event Logged(ByVal logMessage As String, ByVal logLevel As LogLevelEnum)

'---------------------------------
' デバッグログ
'---------------------------------
Public Sub LogDebug(ByVal message As String)
    Call WriteLog(message, LOG_LEVEL_DEBUG)
End Sub

Public Sub LogInfo(ByVal message As String)
    Call WriteLog(message, LOG_LEVEL_INFO)
End Sub

Public Sub LogWarning(ByVal message As String)
    Call WriteLog(message, LOG_LEVEL_WARNING)
End Sub

Public Sub LogError(ByVal message As String)
    Call WriteLog(message, LOG_LEVEL_ERROR)
End Sub

Public Sub LogMessage(ByVal message As String, ByVal logLevel As LogLevelEnum)
    Call WriteLog(message, logLevel)
End Sub

'---------------------------------
' ログ設定
'---------------------------------
' 設定を外部から注入できるように変更
Public Sub Configure(settings As typLoggerSettings)

    With settings
        loggerSettings.LogLevel = .LogLevel
        Set loggerSettings.LogDestination = .LogDestination ' インターフェース型に変更
        loggerSettings.LogFilePath = .LogFilePath
        loggerSettings.LogTableName = .LogTableName
        loggerSettings.LogEventSource = .LogEventSource
        loggerSettings.TimerInterval = .TimerInterval
    End With

    initialized = True

    ' キューの初期化
    Set logQueue = New Queue

    ' タイマーの設定
    timerInterval = loggerSettings.TimerInterval
    timerID = SetTimer(0, 0, timerInterval, AddressOf TimerProc)
End Sub

'---------------------------------
' ログ出力メイン
'---------------------------------
Private Sub WriteLog(ByVal message As String, ByVal logLevel As LogLevelEnum)
    If Not initialized Then
        ' 未設定ならデフォルトを適用
        Dim defaultSettings As typLoggerSettings
        With defaultSettings
            .LogLevel = LOG_LEVEL_INFO
            Set .LogDestination = New FileLogDestination ' デフォルトはファイル出力
            .LogFilePath = DEFAULT_LOG_FILE
            .TimerInterval = 1000 ' デフォルトのタイマー間隔は1秒
        End With

        Me.Configure defaultSettings
    End If

    If logLevel >= loggerSettings.LogLevel Then
        ' ログメッセージをキューに追加
        logQueue.Enqueue Array(message, logLevel)

        ' 同期実行の場合は、ここでイベントを発行
        ' RaiseEvent Logged(message, logLevel)
    End If
End Sub

'---------------------------------
' タイマープロシージャ
'---------------------------------
Private Sub TimerProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
    ' キューからログメッセージを取り出して処理
    While logQueue.Count > 0
        Dim logItem As Variant
        logItem = logQueue.Dequeue

        Dim message As String
        Dim logLevel As LogLevelEnum
        message = logItem(0)
        logLevel = logItem(1)

        ' ログの出力先に合わせて処理を分岐
        loggerSettings.LogDestination.WriteLog message, logLevel, loggerSettings

        ' 非同期イベントを発行
        RaiseEvent Logged(message, logLevel)
    Wend
End Sub

'---------------------------------
' ログレベル名取得（国際化対応）
'---------------------------------
' リソースファイルからログレベル名を取得するように変更
Private Function GetLogLevelName(ByVal logLevel As LogLevelEnum) As String
    Select Case logLevel
        Case LOG_LEVEL_DEBUG: GetLogLevelName = GetResourceString("LogLevelDebug") ' "DEBUG"
        Case LOG_LEVEL_INFO:    GetLogLevelName = GetResourceString("LogLevelInfo") ' "INFO"
        Case LOG_LEVEL_WARNING: GetLogLevelName = GetResourceString("LogLevelWarning") ' "WARNING"
        Case LOG_LEVEL_ERROR:   GetLogLevelName = GetResourceString("LogLevelError") ' "ERROR"
        Case LOG_LEVEL_FATAL:   GetLogLevelName = GetResourceString("LogLevelFatal") ' "FATAL"
        Case Else:              GetLogLevelName = GetResourceString("LogLevelUnknown") ' "UNKNOWN"
    End Select
End Function

'---------------------------------
' リソース文字列取得（ダミー実装）
'---------------------------------
' 本来は外部ファイルやデータベースからリソース文字列を取得する
Private Function GetResourceString(ByVal resourceKey As String) As String
    Select Case resourceKey
        Case "LogLevelDebug": GetResourceString = "DEBUG"
        Case "LogLevelInfo": GetResourceString = "INFO"
        Case "LogLevelWarning": GetResourceString = "WARNING"
        Case "LogLevelError": GetResourceString = "ERROR"
        Case "LogLevelFatal": GetResourceString = "FATAL"
        Case "LogLevelUnknown": GetResourceString = "UNKNOWN"
        Case Else: GetResourceString = ""
    End Select
End Function

'---------------------------------
' ADOConnectionプロパティ
'---------------------------------
' データベースへの接続を遅延させる
Public Property Get ADOConnection() As Object
    If adoConnection Is Nothing Then
        Set adoConnection = CreateObject("ADODB.Connection")
        adoConnection.ConnectionString = GetDbConnectionString()
    End If
    Set ADOConnection = adoConnection
End Property

Public Property Let ADOConnection(ByVal adoConnection As Object)
    Set Me.adoConnection = adoConnection
End Property

'---------------------------------
' データベース接続を閉じる
'---------------------------------
Public Sub CloseConnection()
    If Not adoConnection Is Nothing Then
        On Error Resume Next
        If adoConnection.State = 1 Then ' adStateOpen=1
            adoConnection.Close
        End If
        Set adoConnection = Nothing
        On Error GoTo 0
    End If
End Sub

'---------------------------------
' クラス初期化/終了処理
'---------------------------------
Private Sub Class_Initialize()
    ' デフォルトの設定は不要になったため削除
    ' キューの初期化
    Set logQueue = New Queue
End Sub

Private Sub Class_Terminate()
    ' タイマーの破棄
    If timerID <> 0 Then
        KillTimer 0, timerID
        timerID = 0
    End If

    ' データベース接続を閉じる
    CloseConnection

    ' キューの破棄
    Set logQueue = Nothing
End Sub


' ======================
' 1.8.x ILogDestination インターフェース (クラスモジュール)
' ======================
Option Explicit

' ログ出力メソッド
Public Sub WriteLog(ByVal message As String, ByVal logLevel As LogLevelEnum, settings As Variant)
End Sub


' ======================
' 1.8.x FileLogDestination クラス (クラスモジュール)
' ======================
Option Explicit

Implements ILogDestination

'---------------------------------
' ログ出力 (ファイル)
'---------------------------------
Private Sub ILogDestination_WriteLog(ByVal message As String, ByVal logLevel As LogLevelEnum, settings As Variant)
    On Error GoTo ErrorHandler

    Dim fileNum As Integer
    fileNum = FreeFile

    ' settings からファイルパスを取得
    Dim filePath As String
    filePath = settings.LogFilePath

    Open filePath For Append As #fileNum
    Print #fileNum, Format$(Now, DEFAULT_DATE_FORMAT) & " [" & GetLogLevelName(logLevel) & "] " & message
    Close #fileNum
    Exit Sub

ErrorHandler:
    ' ファイルエラーは無視するか、あるいは別途通知する
    ' エラー発生時にイベントを発行して呼び出し元に通知することも検討
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_FILEIO_ACCESS_DENIED ' 仮のエラーコード
        .Description = "Error writing to log file: " & Err.Description
        .Source = "FileLogDestination"
        .ProcedureName = "ILogDestination_WriteLog"
        .StackTrace = ""
        .OccurredAt = Now
    End With

    ' エラーハンドリングモジュール呼び出し (modErrorの使用を想定)
    HandleError errDetail

    ' デバッグ出力 (一時的な対応)
    Debug.Print "Error writing to log file: " & Err.Description
End Sub

'---------------------------------
' ログレベル名取得（国際化対応）
'---------------------------------
' リソースファイルからログレベル名を取得するように変更
Private Function GetLogLevelName(ByVal logLevel As LogLevelEnum) As String
    Select Case logLevel
        Case LOG_LEVEL_DEBUG: GetLogLevelName = GetResourceString("LogLevelDebug") ' "DEBUG"
        Case LOG_LEVEL_INFO:    GetLogLevelName = GetResourceString("LogLevelInfo") ' "INFO"
        Case LOG_LEVEL_WARNING: GetLogLevelName = GetResourceString("LogLevelWarning") ' "WARNING"
        Case LOG_LEVEL_ERROR:   GetLogLevelName = GetResourceString("LogLevelError") ' "ERROR"
        Case LOG_LEVEL_FATAL:   GetLogLevelName = GetResourceString("LogLevelFatal") ' "FATAL"
        Case Else:              GetLogLevelName = GetResourceString("LogLevelUnknown") ' "UNKNOWN"
    End Select
End Function

'---------------------------------
' リソース文字列取得（ダミー実装）
'---------------------------------
' 本来は外部ファイルやデータベースからリソース文字列を取得する
Private Function GetResourceString(ByVal resourceKey As String) As String
    Select Case resourceKey
        Case "LogLevelDebug": GetResourceString = "DEBUG"
        Case "LogLevelInfo": GetResourceString = "INFO"
        Case "LogLevelWarning": GetResourceString = "WARNING"
        Case "LogLevelError": GetResourceString = "ERROR"
        Case "LogLevelFatal": GetResourceString = "FATAL"
        Case "LogLevelUnknown": GetResourceString = "UNKNOWN"
        Case Else: GetResourceString = ""
    End Select
End Function


' ======================
' 1.8.x DatabaseLogDestination クラス (クラスモジュール)
' ======================
Option Explicit

Implements ILogDestination

'---------------------------------
' ログ出力 (データベース)
'---------------------------------
Private Sub ILogDestination_WriteLog(ByVal message As String, ByVal logLevel As LogLevelEnum, settings As Variant)
    On Error GoTo ErrorHandler

    ' settings からテーブル名とADOコネクションを取得
    Dim tableName As String
    tableName = settings.LogTableName

    Dim conn As Object
    Set conn = GetADOConnection(settings) ' ADOコネクション取得メソッド

    If conn Is Nothing Then
        ' データベース接続エラー処理
        Dim errDetail As typErrorDetail
        With errDetail
            .ErrorCode = ERR_DATABASE_CONNECTION_FAILED
            .Description = "Failed to get ADO connection."
            .Source = "DatabaseLogDestination"
            .ProcedureName = "ILogDestination_WriteLog"
            .StackTrace = ""
            .OccurredAt = Now
        End With

        ' エラーハンドリングモジュール呼び出し (modErrorの使用を想定)
        HandleError errDetail

        ' デバッグ出力 (一時的な対応)
        Debug.Print "Failed to get ADO connection."

        Exit Sub
    End If

    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "INSERT INTO " & tableName & " (LogTime, LogLevel, LogMessage) VALUES (?, ?, ?)"
        .Parameters.Append .CreateParameter("LogTime", 7, 1, , Now)          ' adDate=7, adParamInput=1
        .Parameters.Append .CreateParameter("LogLevel", 3, 1, , logLevel)   ' adInteger=3
        .Parameters.Append .CreateParameter("LogMessage", 200, 1, 4000, message) ' adVarChar=200
        .Execute
    End With

    Exit Sub

ErrorHandler:
    ' データベースエラー処理
    ' エラー発生時にイベントを発行して呼び出し元に通知することも検討
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_DATABASE_QUERY_FAILED ' 仮のエラーコード
        .Description = "Error writing to database: " & Err.Description
        .Source = "DatabaseLogDestination"
        .ProcedureName = "ILogDestination_WriteLog"
        .StackTrace = ""
        .OccurredAt = Now
    End With

    ' エラーハンドリングモジュール呼び出し (modErrorの使用を想定)
    HandleError errDetail

    ' デバッグ出力 (一時的な対応)
    Debug.Print "Error writing to database: " & Err.Description
End Sub

'---------------------------------
' ADOコネクション取得
'---------------------------------
' データベースへの接続を遅延させる
Private Function GetADOConnection(settings As Variant) As Object
    Static conn As Object ' 静的変数に変更

    If conn Is Nothing Then
        Set conn = CreateObject("ADODB.Connection")
        On Error Resume Next
        conn.Open GetDbConnectionString() ' 接続文字列取得関数
        If Err.Number <> 0 Then
            ' 接続エラー処理
            Set conn = Nothing
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    End If

    Set GetADOConnection = conn
End Function


' ======================
' 1.8.x EventLogDestination クラス (クラスモジュール)
' ======================
Option Explicit

Implements ILogDestination

'---------------------------------
' ログ出力 (イベントログ)
'---------------------------------
Private Sub ILogDestination_WriteLog(ByVal message As String, ByVal logLevel As LogLevelEnum, settings As Variant)
    On Error Resume Next

    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    Dim eventType As Integer

    ' settings からイベントソースを取得
    Dim eventSource As String
    eventSource = settings.LogEventSource

    Select Case logLevel
        Case LOG_LEVEL_ERROR, LOG_LEVEL_FATAL
            eventType = 1 ' ERROR
        Case LOG_LEVEL_WARNING
            eventType = 2 ' WARNING
        Case Else
            eventType = 4 ' INFORMATION
    End Select

    objShell.LogEvent eventType, message, eventSource
    Set objShell = Nothing
    On Error GoTo 0
End Sub


' ======================
' 1.8.x ConsoleLogDestination クラス (クラスモジュール)
' ======================
Option Explicit

Implements ILogDestination

'---------------------------------
' ログ出力 (コンソール)
'---------------------------------
Private Sub ILogDestination_WriteLog(ByVal message As String, ByVal logLevel As LogLevelEnum, settings As Variant)
    Debug.Print Format$(Now, DEFAULT_DATE_FORMAT) & " [" & GetLogLevelName(logLevel) & "] " & message
End Sub

'---------------------------------
' ログレベル名取得（国際化対応）
'---------------------------------
' リソースファイルからログレベル名を取得するように変更
Private Function GetLogLevelName(ByVal logLevel As LogLevelEnum) As String
    Select Case logLevel
        Case LOG_LEVEL_DEBUG: GetLogLevelName = GetResourceString("LogLevelDebug") ' "DEBUG"
        Case LOG_LEVEL_INFO:    GetLogLevelName = GetResourceString("LogLevelInfo") ' "INFO"
        Case LOG_LEVEL_WARNING: GetLogLevelName = GetResourceString("LogLevelWarning") ' "WARNING"
        Case LOG_LEVEL_ERROR:   GetLogLevelName = GetResourceString("LogLevelError") ' "ERROR"
        Case LOG_LEVEL_FATAL:   GetLogLevelName = GetResourceString("LogLevelFatal") ' "FATAL"
        Case Else:              GetLogLevelName = GetResourceString("LogLevelUnknown") ' "UNKNOWN"
    End Select
End Function

'---------------------------------
' リソース文字列取得（ダミー実装）
'---------------------------------
' 本来は外部ファイルやデータベースからリソース文字列を取得する
Private Function GetResourceString(ByVal resourceKey As String) As String
    Select Case resource
        Case "LogLevelDebug": GetResourceString = "DEBUG"
        Case "LogLevelInfo": GetResourceString = "INFO"
        Case "LogLevelWarning": GetResourceString = "WARNING"
        Case "LogLevelError": GetResourceString = "ERROR"
        Case "LogLevelFatal": GetResourceString = "FATAL"
        Case "LogLevelUnknown": GetResourceString = "UNKNOWN"
        Case Else: GetResourceString = ""
    End Select
End Function


' ======================
' 1.8.x EmailLogDestination クラス (クラスモジュール)
' ======================
Option Explicit

Implements ILogDestination

'---------------------------------
' ログ出力 (メール) - 将来的に実装
'---------------------------------
Private Sub ILogDestination_WriteLog(ByVal message As String, ByVal logLevel As LogLevelEnum, settings As Variant)
    ' 将来的にメール通知を実装
    Debug.Print "Email notification not yet implemented."
End Sub


' ======================
' 1.8.x Queue クラス (クラスモジュール)
' ======================
Option Explicit

Private queueArray As Variant
Private head As Long
Private tail As Long
Private size As Long

Private Sub Class_Initialize()
    ReDim queueArray(0 To 9) ' 初期サイズは10
    head = 0
    tail = 0
    size = 0
End Sub

Public Sub Enqueue(item As Variant)
    If size = UBound(queueArray) + 1 Then
        ' 配列が満杯の場合はリサイズ
        ReDim Preserve queueArray(0 To UBound(queueArray) * 2)
    End If
    queueArray(tail) = item
    tail = (tail + 1) Mod (UBound(queueArray) + 1)
    size = size + 1
End Sub

Public Function Dequeue() As Variant
    If size = 0 Then
        Err.Raise 9, "Queue", "Queue is empty" ' エラー番号9は「添え字が範囲外」
    End If
    Dequeue = queueArray(head)
    head = (head + 1) Mod (UBound(queueArray) + 1)
    size = size - 1
End Function

Public Property Get Count() As Long
    Count = size
End Property
