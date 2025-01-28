Option Explicit

' インターフェース実装
Implements IErrorHandler
Implements IDisposable

' プライベート変数
Private mErrorPriorities As Object    ' Dictionary<ErrorCode, ErrorPriority>
Private mHandlingStrategies As Object ' Dictionary<ErrorCode, IErrorStrategy>
Private mMaxRetryCount As Long
Private mRetryInterval As Long
Private mRetryBehavior As RetryBehavior
Private mHandlingStats As HandlingStatistics
Private mActiveResources As Collection
Private mLogger As Object            ' FileSystemObject
Private mSettings As Object          ' Dictionary<String, Variant>
Private mSyncLock As Object          ' 同期用オブジェクト
Private mRecoveryChain As RecoveryChain ' 回復チェーン
Private mMetricsCollector As Object       ' パフォーマンスメトリクス収集用
Private mInTransaction As Boolean    ' トランザクション状態
Private mTransactionLevel As Long    ' トランザクションのネストレベル
Private mADOConnection As Object     ' ADOデータベース接続
Private mIsTransactionSupported As Boolean ' トランザクションサポートフラグ
Private mIsAsyncSupported As Boolean      ' 非同期サポートフラグ
Private mCustomBackoffStrategy As IBackoffStrategy ' カスタムバックオフ戦略

' イベント定義
Public Event ErrorHandlingStarted(ByVal errorInfo As ErrorInfo, ByVal priority As ErrorPriority)
Public Event ErrorHandlingCompleted(ByVal errorInfo As ErrorInfo, ByVal result As HandlingResult)
Public Event RetryAttempted(ByVal errorInfo As ErrorInfo, ByVal attempt As Long, ByVal backoffInterval As Long, ByVal result As HandlingResult)
Public Event ResourceReleased(ByVal resourceType As String, ByVal success As Boolean)
Public Event ErrorRecoveryStarted(ByVal errorInfo As ErrorInfo, ByVal strategy As IErrorStrategy)
Public Event ErrorRecoveryCompleted(ByVal errorInfo As ErrorInfo, ByVal success As Boolean)
Public Event ResourceCleanupStarted(ByVal resourceCount As Long)
Public Event ResourceCleanupCompleted(ByVal cleanedCount As Long)
Public Event MetricsUpdated(ByVal metricName As String, ByVal value As Variant)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' インターフェース実装
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' IErrorHandler実装
Private Function IErrorHandler_HandleError(ByRef errorInfo As ErrorInfo, Optional ByVal options As HandlingOptions) As HandlingResult
    On Error GoTo ErrorHandler
    
    ' 同期処理
    MonitorEnter mSyncLock
    
    ' エラー情報の検証
    If Not ValidateErrorInfo(errorInfo) Then
        MonitorExit mSyncLock
        Return hrUnhandled
    End If
    
    ' 優先度の判定
    Dim priority As ErrorPriority
    priority = GetErrorPriority(errorInfo.ErrorCode)
    
    ' イベント通知
    RaiseEvent ErrorHandlingStarted(errorInfo, priority)
    
    ' メトリクスの更新
    UpdateMetrics "HandlingStarted", Now
    
    ' 処理戦略の取得と実行
    Dim strategy As IErrorStrategy
    Set strategy = GetHandlingStrategy(errorInfo.ErrorCode)
    
    If Not strategy Is Nothing Then
        RaiseEvent ErrorRecoveryStarted(errorInfo, strategy)
        
        ' 戦略の実行
        If strategy.RecoverFromError(errorInfo) Then
            UpdateHandlingStats errorInfo, True
            RaiseEvent ErrorRecoveryCompleted(errorInfo, True)
            MonitorExit mSyncLock
            Return hrSuccess
        End If
        
        RaiseEvent ErrorRecoveryCompleted(errorInfo, False)
    End If
    
    ' リトライ処理
    If (options And hoRetry) = hoRetry Then
        Dim retryResult As HandlingResult
        retryResult = AttemptRetry(errorInfo)
        If retryResult <> hrFailed Then
            MonitorExit mSyncLock
            Return retryResult
        End If
    End If
    
    ' 失敗時の処理
    UpdateHandlingStats errorInfo, False
    UpdateMetrics "HandlingFailed", Now
    
    ' イベント通知
    RaiseEvent ErrorHandlingCompleted(errorInfo, hrFailed)
    
    MonitorExit mSyncLock
    Return hrFailed
    
ErrorHandler:
    ' エラー発生時はログに記録し、hrUnhandledを返す
    LogError "HandleError", Err.Description & " (ErrorCode: " & Err.Number & ")"
    
    If mSyncLock Then
        MonitorExit mSyncLock
    End If
    
    Return hrUnhandled
End Function

Private Function IErrorHandler_GetHandlingStatistics(ByVal errorManager As clsErrorManager) As HandlingStatistics
    On Error GoTo ErrorHandler
    
    MonitorEnter mSyncLock
    
    ' 統計情報の更新
    UpdateStatisticsFromManager errorManager
    
    ' 統計情報のコピーを返す
    Set IErrorHandler_GetHandlingStatistics = mHandlingStats.Clone()
    
    MonitorExit mSyncLock
    Exit Function
    
ErrorHandler:
    LogError "GetHandlingStatistics", Err.Description
    Set IErrorHandler_GetHandlingStatistics = Nothing
    
    If mSyncLock Then
        MonitorExit mSyncLock
    End If
End Function

Private Function IErrorHandler_GetPerformanceMetrics(ByVal errorManager As clsErrorManager) As Collection
    On Error GoTo ErrorHandler
    
    MonitorEnter mSyncLock
    
    Dim metrics As New Collection
    
    With mMetricsCollector
        ' 基本メトリクス
        metrics.Add Array("TotalHandled", .Item("TotalHandled"))
        metrics.Add Array("SuccessRate", CalculateSuccessRate())
        metrics.Add Array("AverageHandlingTime", CalculateAverageHandlingTime())
        metrics.Add Array("RetryRate", CalculateRetryRate())
    End With
    
    Set IErrorHandler_GetPerformanceMetrics = metrics
    
    MonitorExit mSyncLock
    Exit Function
    
ErrorHandler:
    LogError "GetPerformanceMetrics", Err.Description
    Set IErrorHandler_GetPerformanceMetrics = New Collection
    
    If mSyncLock Then
        MonitorExit mSyncLock
    End If
End Function

Private Function IErrorHandler_ReleaseResource(ByVal resource As Object, Optional ByVal options As ReleaseOptions) As Boolean
    On Error GoTo ErrorHandler
    
    MonitorEnter mSyncLock
    
    If resource Is Nothing Then
        MonitorExit mSyncLock
        Exit Function
    End If
    
    ' リソース解放前の検証
    If Not ValidateResource(resource) Then
        MonitorExit mSyncLock
        Exit Function
    End If
    
    ' リソースの解放処理
    If TypeOf resource Is IDisposable Then
        resource.Dispose
    ElseIf TypeOf resource Is Object Then
        Set resource = Nothing
    End If
    
    ' アクティブリソースリストから削除
    RemoveFromActiveResources resource
    
    RaiseEvent ResourceReleased(TypeName(resource), True)
    IErrorHandler_ReleaseResource = True
    
    MonitorExit mSyncLock
    Exit Function
    
ErrorHandler:
    LogError "ReleaseResource", Err.Description
    RaiseEvent ResourceReleased(TypeName(resource), False)
    IErrorHandler_ReleaseResource = False
    
    If mSyncLock Then
        MonitorExit mSyncLock
    End If
End Function

Private Sub IErrorHandler_SetErrorPriority(ByVal errorCode As ErrorCode, ByVal priority As ErrorPriority)
    On Error GoTo ErrorHandler
    
    MonitorEnter mSyncLock
    
    If mErrorPriorities.Exists(errorCode) Then
        mErrorPriorities(errorCode) = priority
    Else
        mErrorPriorities.Add errorCode, priority
    End If
    
    MonitorExit mSyncLock
    Exit Sub
    
ErrorHandler:
    LogError "SetErrorPriority", Err.Description
    
    If mSyncLock Then
        MonitorExit mSyncLock
    End If
End Sub

Private Sub IErrorHandler_SetHandlingStrategy(ByVal errorCode As ErrorCode, ByVal strategy As IErrorStrategy)
    On Error GoTo ErrorHandler
    
    MonitorEnter mSyncLock
    
    If mHandlingStrategies.Exists(errorCode) Then
        mHandlingStrategies(errorCode) = strategy
    Else
        mHandlingStrategies.Add errorCode, strategy
    End If
    
    MonitorExit mSyncLock
    Exit Sub
    
ErrorHandler:
    LogError "SetHandlingStrategy", Err.Description
    
    If mSyncLock Then
        MonitorExit mSyncLock
    End If
End Sub

Private Function IErrorHandler_Validate() As Boolean
    On Error GoTo ErrorHandler
    
    MonitorEnter mSyncLock
    
    ' 基本的な検証
    If mErrorPriorities Is Nothing Then Exit Function
    If mHandlingStrategies Is Nothing Then Exit Function
    If mHandlingStats Is Nothing Then Exit Function
    If mActiveResources Is Nothing Then Exit Function
    
    ' 設定の検証
    If Not ValidateSettings Then Exit Function
    
    IErrorHandler_Validate = True
    
    MonitorExit mSyncLock
    Exit Function
    
ErrorHandler:
    LogError "Validate", Err.Description
    IErrorHandler_Validate = False
    
    If mSyncLock Then
        MonitorExit mSyncLock
    End If
End Function

Private Sub IErrorHandler_Cleanup()
    On Error Resume Next
    
    MonitorEnter mSyncLock
    
    ' リソースの解放
    RaiseEvent ResourceCleanupStarted(mActiveResources.Count)
    
    Dim cleanedCount As Long
    cleanedCount = 0
    
    Dim resource As Object
    For Each resource In mActiveResources
        If IErrorHandler_ReleaseResource(resource) Then
            cleanedCount = cleanedCount + 1
        End If
    Next
    
    RaiseEvent ResourceCleanupCompleted(cleanedCount)
    
    ' オブジェクトの解放
    Set mErrorPriorities = Nothing
    Set mHandlingStrategies = Nothing
    Set mActiveResources = Nothing
    Set mHandlingStats = Nothing
    Set mLogger = Nothing
    Set mSettings = Nothing
    Set mSyncLock = Nothing
    Set mCustomBackoffStrategy = Nothing
    
    MonitorExit mSyncLock
End Sub

' IDisposable実装
Private Sub IDisposable_Dispose()
    IErrorHandler_Cleanup
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 初期化処理
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    ' 同期オブジェクトの初期化
    Set mSyncLock = New Object
    
    ' コレクションの初期化
    Set mErrorPriorities = CreateObject("Scripting.Dictionary")
    Set mHandlingStrategies = CreateObject("Scripting.Dictionary")
    Set mActiveResources = New Collection
    Set mHandlingStats = New HandlingStatistics
    Set mMetricsCollector = CreateObject("Scripting.Dictionary")
    
    ' メトリクスの初期化
    InitializeMetrics
    
    ' 設定の読み込み
    LoadSettings
    
    ' ロガーの初期化
    InitializeLogger
    
    ' 機能フラグの初期化
    mIsTransactionSupported = True
    mIsAsyncSupported = False  ' デフォルトでは非同期処理は無効
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' エラー処理とリトライ処理
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' リトライ対象の処理を実行
Private Function ExecuteRetry(ByRef errorInfo As ErrorInfo) As Boolean
    On Error GoTo ErrorHandler
    
    ' エラーコードに基づいて適切なリトライ処理を実行
    Select Case errorInfo.ErrorCode
        Case ecDatabaseError
            ' データベース接続の再試行
            ExecuteRetry = RetryDatabaseOperation(errorInfo)
            
        Case ecNetworkError
            ' ネットワーク操作の再試行
            ExecuteRetry = RetryNetworkOperation(errorInfo)
            
        Case ecFileNotFound, ecFileAccessDenied
            ' ファイルアクセスの再試行
            ExecuteRetry = RetryFileOperation(errorInfo)
            
        Case ecTimeout
            ' タイムアウトした操作の再試行
            ExecuteRetry = RetryTimedOutOperation(errorInfo)
            
        Case Else
            ' その他のエラーは、デフォルトの再試行処理を実行
            ExecuteRetry = ExecuteDefaultRetry(errorInfo)
    End Select
    
    Exit Function
    
ErrorHandler:
    LogError "ExecuteRetry", Err.Description & " (ErrorCode: " & Err.Number & ")"
    ExecuteRetry = False
End Function

' データベース操作の再試行
Private Function RetryDatabaseOperation(ByRef errorInfo As ErrorInfo) As Boolean
    On Error GoTo ErrorHandler
    
    ' データベース接続の再確立を試みる
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    With conn
        .ConnectionTimeout = mSettings("DEFAULT_TIMEOUT_MS") / 1000
        .Open errorInfo.AdditionalInfo("ConnectionString")
        
        ' 接続成功を確認
        If .State = 1 Then ' adStateOpen
            .Close
            RetryDatabaseOperation = True
            Exit Function
        End If
    End With
    
    RetryDatabaseOperation = False
    Exit Function
    
ErrorHandler:
    LogError "RetryDatabaseOperation", Err.Description
    RetryDatabaseOperation = False
End Function

' ネットワーク操作の再試行
Private Function RetryNetworkOperation(ByRef errorInfo As ErrorInfo) As Boolean
    On Error GoTo ErrorHandler
    
    ' ネットワーク接続の確認
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    With http
        .Open "HEAD", errorInfo.AdditionalInfo("URL"), False
        .send
        
        ' レスポンスコードの確認
        If .Status >= 200 And .Status < 300 Then
            RetryNetworkOperation = True
            Exit Function
        End If
    End With
    
    RetryNetworkOperation = False
    Exit Function
    
ErrorHandler:
    LogError "RetryNetworkOperation", Err.Description
    RetryNetworkOperation = False
End Function

' ファイル操作の再試行
Private Function RetryFileOperation(ByRef errorInfo As ErrorInfo) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' ファイルの存在確認
    If fso.FileExists(errorInfo.AdditionalInfo("FilePath")) Then
        ' ファイルアクセス権限の確認
        Dim file As Object
        Set file = fso.GetFile(errorInfo.AdditionalInfo("FilePath"))
        
        ' 読み取りテスト
        Dim ts As Object
        Set ts = file.OpenAsTextStream(1) ' ForReading
        ts.Close
        
        RetryFileOperation = True
        Exit Function
    End If
    
    RetryFileOperation = False
    Exit Function
    
ErrorHandler:
    LogError "RetryFileOperation", Err.Description
    RetryFileOperation = False
End Function

' タイムアウト操作の再試行
Private Function RetryTimedOutOperation(ByRef errorInfo As ErrorInfo) As Boolean
    On Error GoTo ErrorHandler
    
    ' タイムアウト値を一時的に延長して再試行
    Dim originalTimeout As Long
    originalTimeout = mSettings("DEFAULT_TIMEOUT_MS")
    mSettings("DEFAULT_TIMEOUT_MS") = originalTimeout * 2
    
    ' 元の操作を再実行
    RetryTimedOutOperation = ExecuteDefaultRetry(errorInfo)
    
    ' タイムアウト値を元に戻す
    mSettings("DEFAULT_TIMEOUT_MS") = originalTimeout
    Exit Function
    
ErrorHandler:
    LogError "RetryTimedOutOperation", Err.Description
    RetryTimedOutOperation = False
End Function

' リトライ処理の実行
Private Function AttemptRetry(ByRef errorInfo As ErrorInfo) As HandlingResult
    Dim attempt As Long
    Dim backoffInterval As Long
    Dim startTime As Date
    
    ' リトライ開始時刻を記録
    startTime = Now
    
    For attempt = 1 To mMaxRetryCount
        ' バックオフ間隔の計算
        backoffInterval = CalculateBackoffInterval(attempt)
        
        ' タイムアウトチェック
        If DateDiff("s", startTime, Now) > mSettings("DEFAULT_TIMEOUT_MS") / 1000 Then
            LogError "AttemptRetry", "Retry timeout exceeded"
            RaiseEvent RetryAttempted(errorInfo, attempt, backoffInterval, hrFailed)
            AttemptRetry = hrFailed
            Exit Function
        End If
        
        ' バックオフ待機
        Sleep backoffInterval
        
        ' リトライ実行
        If ExecuteRetry(errorInfo) Then
            ' 成功時の統計更新
            UpdateRetryStatistics errorInfo, attempt, True
            
            ' イベント通知
            RaiseEvent RetryAttempted(errorInfo, attempt, backoffInterval, hrSuccess)
            
            AttemptRetry = hrSuccess
            Exit Function
        End If
        
        ' 失敗時の統計更新
        UpdateRetryStatistics errorInfo, attempt, False
        
        ' イベント通知
        RaiseEvent RetryAttempted(errorInfo, attempt, backoffInterval, hrFailed)
        
        ' エラー情報の更新
        With errorInfo
            .RetryCount = attempt
            .LastRetryTimestamp = Now
        End With
    Next attempt
    
    ' 全リトライ失敗
    AttemptRetry = hrFailed
End Function

' リトライ統計の更新
Private Sub UpdateRetryStatistics(ByRef errorInfo As ErrorInfo, ByVal attempt As Long, ByVal success As Boolean)
    With mHandlingStats
        If success Then
            .RetrySuccessCount = .RetrySuccessCount + 1
        Else
            .RetryFailureCount = .RetryFailureCount + 1
        End If
        
        ' 平均リトライ回数の更新
        .TotalRetryAttempts = .TotalRetryAttempts + 1
        .AverageRetryCount = .TotalRetryAttempts / (.RetrySuccessCount + .RetryFailureCount)
    End With
End Sub

' デフォルトの再試行処理
Private Function ExecuteDefaultRetry(ByRef errorInfo As ErrorInfo) As Boolean
    On Error GoTo ErrorHandler
    
    ' エラー情報から再試行可能な操作を特定
    If errorInfo.AdditionalInfo.Exists("RetryableOperation") Then
        Dim operation As IRetryableOperation
        Set operation = errorInfo.AdditionalInfo("RetryableOperation")
        
        ' 操作の再実行
        ExecuteDefaultRetry = operation.Execute()
    Else
        ExecuteDefaultRetry = False
    End If
    
    Exit Function
    
ErrorHandler:
    LogError "ExecuteDefaultRetry", Err.Description
    ExecuteDefaultRetry = False
End Function

' トランザクション開始
Private Sub BeginTransaction()
    On Error GoTo ErrorHandler
    
    If Not mIsTransactionSupported Then
        Err.Raise vbObjectError + 1, "BeginTransaction", "Transactions are not supported"
    End If
    
    If mInTransaction Then
        ' ネストされたトランザクション
        mTransactionLevel = mTransactionLevel + 1
    Else
        ' 新しいトランザクション
        mInTransaction = True
        mTransactionLevel = 1
        
        ' ADOトランザクションの開始
        If Not mADOConnection Is Nothing Then
            mADOConnection.BeginTrans
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    LogError "BeginTransaction", Err.Description
    Err.Raise Err.Number, "BeginTransaction", "Failed to begin transaction: " & Err.Description
End Sub

' トランザクションのコミット
Private Sub CommitTransaction()
    On Error GoTo ErrorHandler
    
    If Not mInTransaction Then
        Err.Raise vbObjectError + 2, "CommitTransaction", "No active transaction"
    End If
    
    mTransactionLevel = mTransactionLevel - 1
    
    If mTransactionLevel = 0 Then
        ' 最外部のトランザクション
        If Not mADOConnection Is Nothing Then
            mADOConnection.CommitTrans
        End If
        mInTransaction = False
    End If
    
    Exit Sub
    
ErrorHandler:
    LogError "CommitTransaction", Err.Description
    Err.Raise Err.Number, "CommitTransaction", "Failed to commit transaction: " & Err.Description
End Sub

' トランザクションのロールバック
Private Sub RollbackTransaction()
    On Error GoTo ErrorHandler
    
    If Not mInTransaction Then
        Exit Sub ' 既にロールバック済みまたはトランザクションなし
    End If
    
    If Not mADOConnection Is Nothing Then
        mADOConnection.RollbackTrans
    End If
    
    mInTransaction = False
    mTransactionLevel = 0
    
    Exit Sub
    
ErrorHandler:
    LogError "RollbackTransaction", Err.Description
    ' ロールバックのエラーは上位に伝播させない（既に例外処理中の可能性があるため）
End Sub

' トランザクション状態の確認
Private Function IsInTransaction() As Boolean
    IsInTransaction = mInTransaction
End Function

' トランザクションレベルの取得
Private Function GetTransactionLevel() As Long
    GetTransactionLevel = mTransactionLevel
End Function

' ADO接続の設定
Private Sub SetADOConnection(ByVal connection As Object)
    Set mADOConnection = connection
End Sub

' 非同期サポートの設定
Public Property Let AsyncSupported(ByVal value As Boolean)
    mIsAsyncSupported = value
End Property

Public Property Get AsyncSupported() As Boolean
    AsyncSupported = mIsAsyncSupported
End Property

' トランザクションサポートの設定
Public Property Let TransactionSupported(ByVal value As Boolean)
    mIsTransactionSupported = value
End Property

Public Property Get TransactionSupported() As Boolean
    TransactionSupported = mIsTransactionSupported
End Property

' パフォーマンスメトリクスの取得
Public Function GetPerformanceMetrics() As Collection
    Dim metrics As New Collection
    
    With mMetricsCollector
        ' 平均処理時間の計算
        If .Exists("HandlingTimes") And .Item("HandlingTimes").Count > 0 Then
            Dim totalTime As Long
            Dim time As Variant
            For Each time In .Item("HandlingTimes")
                totalTime = totalTime + CLng(time)
            Next
            metrics.Add Array("AverageHandlingTime", totalTime / .Item("HandlingTimes").Count)
        End If
        
        ' 成功率の計算
        If .Exists("SuccessCount") And .Exists("TotalCount") Then
            Dim successRate As Double
            If .Item("TotalCount") > 0 Then
                successRate = .Item("SuccessCount") / .Item("TotalCount") * 100
            End If
            metrics.Add Array("SuccessRate", successRate)
        End If
        
        ' エラーコード別の統計
        Dim key As Variant
        For Each key In .Keys
            If Left(key, 10) = "ErrorCode_" Then
                metrics.Add Array(key, .Item(key))
            End If
        Next
    End With
    
    Set GetPerformanceMetrics = metrics
End Function

' エラーパターンの分析
Public Function AnalyzeErrorPatterns(ByVal timeWindow As Long) As Collection
    Dim patterns As New Collection
    Dim currentTime As Date
    currentTime = Now
    
    With mMetricsCollector
        ' エラーコードの出現パターンを分析
        Dim errorCounts As Object
        Set errorCounts = CreateObject("Scripting.Dictionary")
        
        ' エラー履歴から時間枠内のエラーを集計
        Dim history As Collection
        Set history = mHandlingStats.GetErrorHistory()
        
        Dim errorInfo As ErrorInfo
        For Each errorInfo In history
            If DateDiff("s", errorInfo.Timestamp, currentTime) <= timeWindow Then
                Dim errorKey As String
                errorKey = "ErrorCode_" & errorInfo.ErrorCode
                
                If errorCounts.Exists(errorKey) Then
                    errorCounts(errorKey) = errorCounts(errorKey) + 1
                Else
                    errorCounts.Add errorKey, 1
                End If
            End If
        Next
        
        ' パターンの検出
        Dim key As Variant
        For Each key In errorCounts.Keys
            Dim count As Long
            count = errorCounts(key)
            
            ' 頻出エラーの検出（閾値: 全体の10%以上）
            If count >= (history.Count * 0.1) Then
                patterns.Add Array("FrequentError", key, count)
            End If
            
            ' バースト的なエラーの検出（短時間での急増）
            If IsErrorBurst(key, timeWindow) Then
                patterns.Add Array("ErrorBurst", key, count)
            End If
        Next
    End With
    
    Set AnalyzeErrorPatterns = patterns
End Function

' エラーのバースト検出
Private Function IsErrorBurst(ByVal errorKey As String, ByVal timeWindow As Long) As Boolean
    Const BURST_THRESHOLD As Long = 5 ' 5回以上のエラーを急増とみなす
    Const BURST_INTERVAL As Long = 60 ' 1分以内
    
    Dim history As Collection
    Set history = mHandlingStats.GetErrorHistory()
    
    Dim errorCount As Long
    Dim lastErrorTime As Date
    lastErrorTime = Now
    
    Dim errorInfo As ErrorInfo
    For Each errorInfo In history
        If "ErrorCode_" & errorInfo.ErrorCode = errorKey Then
            If DateDiff("s", errorInfo.Timestamp, lastErrorTime) <= BURST_INTERVAL Then
                errorCount = errorCount + 1
                If errorCount >= BURST_THRESHOLD Then
                    IsErrorBurst = True
                    Exit Function
                End If
            Else
                errorCount = 1
            End If
            lastErrorTime = errorInfo.Timestamp
        End If
    Next
    
    IsErrorBurst = False
End Function

' バックオフ間隔の計算
Private Function CalculateBackoffInterval(ByVal attempt As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim interval As Long
    
    Select Case mRetryBehavior
        Case rbLinear
            ' 線形バックオフ: 試行回数に比例して増加
            interval = mRetryInterval * attempt
            
        Case rbExponential
            ' 指数バックオフ: 2の累乗で増加
            interval = mRetryInterval * (2 ^ (attempt - 1))
            
        Case rbFibonacci
            ' フィボナッチバックオフ: フィボナッチ数列に基づいて増加
            interval = CalculateFibonacciInterval(attempt)
            
        Case rbCustom
            ' カスタムバックオフ: 独自の計算ロジック
            If Not mCustomBackoffStrategy Is Nothing Then
                interval = mCustomBackoffStrategy.CalculateInterval(attempt, mRetryInterval)
            Else
                interval = mRetryInterval
            End If
            
        Case Else
            ' デフォルト: 固定間隔
            interval = mRetryInterval
    End Select
    
    ' ジッター（ランダムな変動）の追加
    If mSettings("USE_JITTER") Then
        Dim jitterRange As Long
        jitterRange = interval * 0.1 ' 10%のジッター範囲
        interval = interval + Int((Rnd * jitterRange * 2) - jitterRange)
    End If
    
    ' 範囲の制限
    interval = WorksheetFunction.Max(mSettings("MIN_RETRY_INTERVAL"), _
               WorksheetFunction.Min(interval, mSettings("MAX_RETRY_INTERVAL")))
    
    CalculateBackoffInterval = interval
    Exit Function
    
ErrorHandler:
    LogError "CalculateBackoffInterval", Err.Description
    CalculateBackoffInterval = mRetryInterval ' エラー時はデフォルト間隔を返す
End Function

' フィボナッチ数列に基づくリトライ間隔を計算
Private Function CalculateFibonacciInterval(ByVal attempt As Long) As Long
    On Error GoTo ErrorHandler
    
    Dim a As Long, b As Long, temp As Long
    Dim i As Long
    
    ' フィボナッチ数列の初期値
    a = 0
    b = mRetryInterval
    
    ' n番目のフィボナッチ数を計算
    For i = 1 To attempt
        temp = a
        a = b
        b = temp + b
        
        ' オーバーフロー防止
        If b < 0 Or b > mSettings("MAX_RETRY_INTERVAL") Then
            a = mSettings("MAX_RETRY_INTERVAL")
            Exit For
        End If
    Next i
    
    CalculateFibonacciInterval = a
    Exit Function
    
ErrorHandler:
    LogError "CalculateFibonacciInterval", Err.Description
    CalculateFibonacciInterval = mRetryInterval ' エラー時はデフォルト間隔を返す
End Function

' コンテキストの復元
Private Sub RestoreContext(ByVal context As ExecutionContext)
    On Error GoTo ErrorHandler
    
    If context Is Nothing Then Exit Sub
    
    ' エラー情報の復元
    If Not context.ErrorInfo Is Nothing Then
        With context.ErrorInfo
            .RetryCount = 0 ' リトライカウントをリセット
            .LastRetryTimestamp = Now
        End With
    End If
    
    Exit Sub
    
ErrorHandler:
    LogError "RestoreContext", Err.Description
End Sub

' 設定の読み込み
Private Sub LoadSettings()
    On Error GoTo ErrorHandler
    
    Set mSettings = CreateObject("Scripting.Dictionary")
    
    ' デフォルト設定
    With mSettings
        ' リトライ関連
        .Add "MAX_RETRY_COUNT", 3
        .Add "DEFAULT_RETRY_INTERVAL", 1000
        .Add "MIN_RETRY_INTERVAL", 100
        .Add "MAX_RETRY_INTERVAL", 30000
        .Add "USE_JITTER", True
        
        ' タイムアウト関連
        .Add "DEFAULT_TIMEOUT_MS", 30000
        .Add "CONNECTION_TIMEOUT_MS", 5000
        .Add "OPERATION_TIMEOUT_MS", 30000
        
        ' リソース管理
        .Add "MAX_ACTIVE_RESOURCES", 100
        .Add "RESOURCE_CLEANUP_INTERVAL", 300000
        .Add "MAX_RESOURCE_AGE_MS", 3600000
        
        ' ログ関連
        .Add "LOG_FILE_PATH", "error_handler.log"
        .Add "MAX_LOG_SIZE", 5242880
        .Add "LOG_ROTATION_COUNT", 5
        .Add "LOG_LEVEL", "INFO"
        
        ' エラー分析
        .Add "ERROR_HISTORY_SIZE", 1000
        .Add "PATTERN_ANALYSIS_WINDOW", 3600
        .Add "BURST_DETECTION_THRESHOLD", 5
        .Add "BURST_DETECTION_INTERVAL", 60
        
        ' パフォーマンス
        .Add "METRICS_COLLECTION_ENABLED", True
        .Add "METRICS_SAMPLE_INTERVAL", 60000
        .Add "METRICS_HISTORY_SIZE", 1000
        
        ' 回復戦略
        .Add "MAX_RECOVERY_CHAIN_LENGTH", 5
        .Add "RECOVERY_TIMEOUT_MS", 60000
        .Add "FALLBACK_ENABLED", True
    End With
    
    ' 設定ファイルからの読み込み（存在する場合）
    LoadSettingsFromFile
    
    ' 環境変数からの設定上書き
    LoadSettingsFromEnvironment
    
    ' 設定の検証
    ValidateSettings
    
    Exit Sub
    
ErrorHandler:
    LogError "LoadSettings", "Failed to load settings: " & Err.Description
    ' 最小限の設定で続行
    InitializeMinimalSettings
End Sub

' 設定ファイルからの読み込み
Private Sub LoadSettingsFromFile()
    On Error GoTo ErrorHandler
    
    Dim settingsPath As String
    settingsPath = App.Path & "\error_handler_settings.json"
    
    If Dir(settingsPath) <> "" Then
        ' JSONファイルから設定を読み込む
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        Dim file As Object
        Set file = fso.OpenTextFile(settingsPath, 1) ' ForReading
        
        Dim content As String
        content = file.ReadAll
        file.Close
        
        ' JSONのパース（簡易実装）
        ParseSettingsJson content
    End If
    Exit Sub
    
ErrorHandler:
    LogError "LoadSettingsFromFile", "Failed to load settings file: " & Err.Description
End Sub

' 環境変数からの設定読み込み
Private Sub LoadSettingsFromEnvironment()
    On Error Resume Next
    
    ' 環境変数のプレフィックス
    Const ENV_PREFIX As String = "ERROR_HANDLER_"
    
    ' 主要な設定の環境変数チェック
    CheckEnvironmentSetting "MAX_RETRY_COUNT", ENV_PREFIX & "MAX_RETRY_COUNT"
    CheckEnvironmentSetting "DEFAULT_RETRY_INTERVAL", ENV_PREFIX & "RETRY_INTERVAL"
    CheckEnvironmentSetting "LOG_LEVEL", ENV_PREFIX & "LOG_LEVEL"
    CheckEnvironmentSetting "LOG_FILE_PATH", ENV_PREFIX & "LOG_FILE"
End Sub

' 環境変数の設定値チェック
Private Sub CheckEnvironmentSetting(ByVal settingKey As String, ByVal envKey As String)
    Dim envValue As String
    envValue = Environ(envKey)
    
    If envValue <> "" Then
        ' 数値の場合は変換
        If IsNumeric(mSettings(settingKey)) Then
            mSettings(settingKey) = CLng(envValue)
        Else
            mSettings(settingKey) = envValue
        End If
    End If
End Sub

' 設定の検証
Private Sub ValidateSettings()
    ' 必須設定の存在確認
    Dim requiredSettings As Variant
    requiredSettings = Array("MAX_RETRY_COUNT", "DEFAULT_RETRY_INTERVAL", "LOG_FILE_PATH")
    
    Dim setting As Variant
    For Each setting In requiredSettings
        If Not mSettings.Exists(setting) Then
            Err.Raise vbObjectError + 1, "ValidateSettings", "Required setting missing: " & setting
        End If
    Next
    
    ' 値の範囲チェック
    With mSettings
        .Item("MAX_RETRY_COUNT") = ValidateRange(.Item("MAX_RETRY_COUNT"), 1, 10)
        .Item("DEFAULT_RETRY_INTERVAL") = ValidateRange(.Item("DEFAULT_RETRY_INTERVAL"), 100, 60000)
        .Item("MAX_LOG_SIZE") = ValidateRange(.Item("MAX_LOG_SIZE"), 1048576, 104857600)
    End With
End Sub

' 値の範囲チェック
Private Function ValidateRange(ByVal value As Long, ByVal minValue As Long, ByVal maxValue As Long) As Long
    ValidateRange = WorksheetFunction.Max(minValue, WorksheetFunction.Min(value, maxValue))
End Function

' 最小限の設定で初期化
Private Sub InitializeMinimalSettings()
    mSettings.RemoveAll
    
    With mSettings
        .Add "MAX_RETRY_COUNT", 3
        .Add "DEFAULT_RETRY_INTERVAL", 1000
        .Add "LOG_FILE_PATH", "error_handler.log"
        .Add "MAX_LOG_SIZE", 5242880
    End With
End Sub