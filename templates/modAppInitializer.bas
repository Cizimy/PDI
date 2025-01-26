Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "modAppInitializer"

' ======================
' 初期化状態管理
' ======================
Private Type InitializationState
    ConfigInitialized As Boolean
    LoggingInitialized As Boolean
    DatabaseInitialized As Boolean
    ErrorHandlersInitialized As Boolean
End Type

Private mInitState As InitializationState

' ======================
' 公開メソッド
' ======================
Public Sub InitializeApplication()
    On Error GoTo ErrorHandler
    
    ' 設定の初期化（最初に実行）
    If Not mInitState.ConfigInitialized Then
        modConfig.InitializeModule
        mInitState.ConfigInitialized = True
    End If
    
    ' ロギングシステムの初期化
    If Not mInitState.LoggingInitialized Then
        InitializeLogging
        mInitState.LoggingInitialized = True
    End If
    
    ' データベース関連の初期化
    If Not mInitState.DatabaseInitialized Then
        InitializeDatabase
        mInitState.DatabaseInitialized = True
    End If
    
    ' エラーハンドラーの初期化（最後に実行）
    If Not mInitState.ErrorHandlersInitialized Then
        InitializeErrorHandlers
        mInitState.ErrorHandlersInitialized = True
    End If
    
    Exit Sub

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "アプリケーションの初期化中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "InitializeApplication"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
End Sub

' ======================
' プライベートメソッド
' ======================
Private Sub InitializeLogging()
    On Error GoTo ErrorHandler
    
    ' ロガー設定の初期化
    Dim loggerSettings As DefaultLoggerSettings
    Set loggerSettings = New DefaultLoggerSettings
    loggerSettings.Initialize modConfig.Settings.DatabaseConnectionString
    
    ' デフォルトロガーの設定
    With New clsLogger
        .Configure loggerSettings
        .Log MODULE_NAME, "ロギングシステムが初期化されました", 0
    End With
    
    Exit Sub

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "ロギングシステムの初期化中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "InitializeLogging"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
End Sub

Private Sub InitializeDatabase()
    On Error GoTo ErrorHandler
    
    ' データベースユーティリティの初期化
    modDatabaseUtils.InitializeModule
    
    ' データベース接続プールの初期化
    With New ConnectionPool
        ' IDatabaseConfigインターフェースを通じて接続文字列を取得
        .Initialize modConfig
    End With
    
    Exit Sub

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "データベースシステムの初期化中にエラーが発生しました: " & Err.Description
        .Category = ECDatabase
        .Source = MODULE_NAME
        .ProcedureName = "InitializeDatabase"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
End Sub

Private Sub InitializeErrorHandlers()
    On Error GoTo ErrorHandler
    
    ' エラーハンドラーの初期化
    ' この時点で設定とロギングは初期化済みであることが保証される
    With DatabaseConnectionErrorHandler.Create
(modConfig)
        ' 必要な初期化処理があれば実行
    End With
    
    Exit Sub

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "エラーハンドラーの初期化中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "InitializeErrorHandlers"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
End Sub