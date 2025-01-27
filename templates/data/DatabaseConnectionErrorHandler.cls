VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseConnectionErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' モジュール情報
Private Const MODULE_NAME As String = "DatabaseConnectionErrorHandler"

' 定数定義
Private Const DEFAULT_RETRY_INTERVAL As Long = 1000    ' ミリ秒
Private Const MAX_RETRY_COUNT As Long = 3
Private Const MAX_POOL_SIZE As Long = 100              ' 最大プールサイズ
Private Const MIN_POOL_SIZE As Long = 5                ' 最小プールサイズ
Private Const POOL_CLEANUP_INTERVAL As Long = 300000   ' プールクリーンアップ間隔（5分）
Private Const DIAGNOSTIC_TIMEOUT As Long = 5000        ' 診断タイムアウト（5秒）
Private Const MAX_FALLBACK_ATTEMPTS As Long = 2        ' 最大フォールバック試行回数
Private Const METRICS_INTERVAL As Long = 60000         ' メトリクス収集間隔（1分）

' イベント定義
Public Event RetryAttempted(ByVal attempt As Long, ByVal success As Boolean)
Public Event FallbackActivated(ByVal reason As String, ByVal connectionString As String)
Public Event PoolStateChanged(ByVal activeConnections As Long, ByVal maxPoolSize As Long)
Public Event DiagnosticsCompleted(ByVal results As Collection)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)
Public Event ConnectionRecovered(ByVal originalError As ErrorInfo, ByVal recoveryTime As Double)
Public Event MetricsCollected(ByVal metrics As Collection)

' プライベート変数
Private mConnectionString As String
Private mFallbackConnectionString As String
Private mLock As ILock
Private mLogger As ILogger
Private mEmergencyLogger As IEmergencyLogger
Private mConfig As IDatabaseConfig
Private mUserNotifier As IUserNotifier
Private mPerformanceMonitor As IPerformanceMonitor
Private mConnectionPool As Collection
Private mLastCleanupTime As Date
Private mLastMetricsTime As Date
Private mRetryStrategies As Collection
Private mDiagnosticResults As Collection

Implements IErrorHandler

' コンストラクタ
Public Function Create( _
    ByVal lock As ILock, _
    ByVal logger As ILogger, _
    ByVal emergencyLogger As IEmergencyLogger, _
    ByVal config As IDatabaseConfig, _
    ByVal userNotifier As IUserNotifier, _
    ByVal performanceMonitor As IPerformanceMonitor _
) As DatabaseConnectionErrorHandler
    
    Dim instance As New DatabaseConnectionErrorHandler
    
    ' 依存性の注入と検証
    With instance
        ValidateDependency lock, "ILock"
        Set .Lock = lock
        
        ValidateDependency logger, "ILogger"
        Set .Logger = logger
        
        ValidateDependency emergencyLogger, "IEmergencyLogger"
        Set .EmergencyLogger = emergencyLogger
        
        ValidateDependency config, "IDatabaseConfig"
        Set .Config = config
        
        ValidateDependency userNotifier, "IUserNotifier"
        Set .UserNotifier = userNotifier
        
        ValidateDependency performanceMonitor, "IPerformanceMonitor"
        Set .PerformanceMonitor = performanceMonitor
        
        .Initialize
    End With
    
    Set Create = instance
End Function

' プロパティ
Public Property Set Lock(ByVal value As ILock)
    Set mLock = value
End Property

Public Property Set Logger(ByVal value As ILogger)
    Set mLogger = value
End Property

Public Property Set EmergencyLogger(ByVal value As IEmergencyLogger)
    Set mEmergencyLogger = value
End Property

Public Property Set Config(ByVal value As IDatabaseConfig)
    Set mConfig = value
    mConnectionString = mConfig.GetConnectionString
    mFallbackConnectionString = mConfig.GetFallbackConnectionString
End Property

Public Property Set UserNotifier(ByVal value As IUserNotifier)
    Set mUserNotifier = value
End Property

Public Property Set PerformanceMonitor(ByVal value As IPerformanceMonitor)
    Set mPerformanceMonitor = value
End Property

' 初期化
Private Sub Initialize()
    Set mConnectionPool = New Collection
    Set mRetryStrategies = New Collection
    Set mDiagnosticResults = New Collection
    mLastCleanupTime = Now
    mLastMetricsTime = Now
    InitializeRetryStrategies
End Sub

' IErrorHandler実装
Private Function IErrorHandler_HandleError(ByRef errorDetail As ErrorInfo) As Boolean
    Dim proceed As Boolean
    proceed = False
    
    mPerformanceMonitor.StartMeasurement "HandleError"
    
    ' エラー情報をログに記録
    LogError errorDetail
    
    ' エラー情報の補完
    CompleteErrorInfo errorDetail
    
    ' 診断の実行
    ExecuteDiagnostics errorDetail
    
    ' リトライ処理の実行
    If AttemptRetry(errorDetail) Then
        proceed = True
        errorDetail.RecoverySuccessful = True
        RaiseEvent ConnectionRecovered(errorDetail, mPerformanceMonitor.GetElapsedTime)
    Else
        ' フォールバック処理の実行
        If AttemptFallback(errorDetail) Then
            proceed = True
            errorDetail.RecoverySuccessful = True
        Else
            errorDetail.RecoverySuccessful = False
            NotifyUser errorDetail, vbCritical
        End If
    End If
    
    ' メトリクスの収集
    CollectMetrics
    
    ' プールの管理
    ManageConnectionPool
    
    mPerformanceMonitor.EndMeasurement "HandleError"
    
    IErrorHandler_HandleError = proceed
End Function

' プライベートメソッド
Private Sub InitializeRetryStrategies()
    With mRetryStrategies
        .Add New LinearRetryStrategy
        .Add New ExponentialRetryStrategy
        .Add New FibonacciRetryStrategy
    End With
End Sub

Private Sub CompleteErrorInfo(ByRef errorDetail As ErrorInfo)
    With errorDetail
        .AddAdditionalInfo "ConnectionString", mConnectionString
        .AddAdditionalInfo "MaxRetries", CStr(MAX_RETRY_COUNT)
        .AddAdditionalInfo "PoolSize", CStr(mConnectionPool.Count)
        .AddAdditionalInfo "LastDiagnostic", GetLastDiagnosticResult
        .RecoveryAttempted = True
    End With
End Sub

Private Function AttemptRetry(ByRef errorDetail As ErrorInfo) As Boolean
    Dim strategy As IRetryStrategy
    Set strategy = SelectRetryStrategy(errorDetail)
    
    Dim retryCount As Long
    Dim success As Boolean
    
    For retryCount = 1 To MAX_RETRY_COUNT
        Wait strategy.GetNextInterval(retryCount)
        
        errorDetail.AddAdditionalInfo "CurrentRetry", CStr(retryCount)
        
        If TryDatabaseConnection Then
            success = True
            Exit For
        End If
        
        RaiseEvent RetryAttempted(retryCount, False)
        LogRetryAttempt retryCount, errorDetail
    Next retryCount
    
    If success Then RaiseEvent RetryAttempted(retryCount, True)
    
    AttemptRetry = success
End Function

Private Function AttemptFallback(ByRef errorDetail As ErrorInfo) As Boolean
    If Len(mFallbackConnectionString) = 0 Then Exit Function
    
    Dim originalConnection As String
    originalConnection = mConnectionString
    mConnectionString = mFallbackConnectionString
    
    RaiseEvent FallbackActivated("Primary connection failed", mFallbackConnectionString)
    
    Dim success As Boolean
    success = TryDatabaseConnection
    
    If Not success Then mConnectionString = originalConnection
    
    AttemptFallback = success
End Function

Private Sub ExecuteDiagnostics(ByRef errorDetail As ErrorInfo)
    mPerformanceMonitor.StartMeasurement "Diagnostics"
    
    Dim results As New Collection
    
    ' ネットワーク診断
    results.Add DiagnoseNetwork
    
    ' サーバー診断
    results.Add DiagnoseServer
    
    ' 認証診断
    results.Add DiagnoseAuthentication
    
    Set mDiagnosticResults = results
    RaiseEvent DiagnosticsCompleted(results)
    
    mPerformanceMonitor.EndMeasurement "Diagnostics"
End Sub

Private Sub ManageConnectionPool()
    If DateDiff("s", mLastCleanupTime, Now) > POOL_CLEANUP_INTERVAL / 1000 Then
        CleanupConnectionPool
        mLastCleanupTime = Now
    End If
    
    RaiseEvent PoolStateChanged(mConnectionPool.Count, MAX_POOL_SIZE)
End Sub

Private Sub CollectMetrics()
    If DateDiff("s", mLastMetricsTime, Now) > METRICS_INTERVAL / 1000 Then
        Dim metrics As New Collection
        
        With metrics
            .Add mPerformanceMonitor.GetAllMeasurements
            .Add GetConnectionPoolMetrics
            .Add GetRetryMetrics
            .Add GetDiagnosticMetrics
        End With
        
        RaiseEvent MetricsCollected(metrics)
        mLastMetricsTime = Now
    End If
End Sub

Private Function SelectRetryStrategy(ByRef errorDetail As ErrorInfo) As IRetryStrategy
    ' エラーの種類に応じて最適な戦略を選択
    ' 実装は省略
End Function

Private Function DiagnoseNetwork() As DiagnosticResult
    ' ネットワーク診断の実装
    ' 実装は省略
End Function

Private Function DiagnoseServer() As DiagnosticResult
    ' サーバー診断の実装
    ' 実装は省略
End Function

Private Function DiagnoseAuthentication() As DiagnosticResult
    ' 認証診断の実装
    ' 実装は省略
End Function

Private Sub CleanupConnectionPool()
    ' 接続プールのクリーンアップ
    ' 実装は省略
End Sub

Private Function GetConnectionPoolMetrics() As Collection
    ' 接続プールのメトリクス取得
    ' 実装は省略
End Function

Private Function GetRetryMetrics() As Collection
    ' リトライのメトリクス取得
    ' 実装は省略
End Function

Private Function GetDiagnosticMetrics() As Collection
    ' 診断のメトリクス取得
    ' 実装は省略
End Function

Private Function GetLastDiagnosticResult() As String
    ' 最後の診断結果を取得
    ' 実装は省略
End Function

Private Sub ValidateDependency(ByVal dependency As Object, ByVal name As String)
    If dependency Is Nothing Then
        Err.Raise 5, "DatabaseConnectionErrorHandler.Create", name & " is required"
    End If
End Sub

' 既存のメソッド（TryDatabaseConnection, LogError, NotifyUser, Wait）は変更なし