VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventLogImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements IEventLog

' 定数定義
Private Const MODULE_NAME As String = "EventLogImpl"
Private Const MAX_MESSAGE_LENGTH As Long = 32766     ' メッセージの最大長
Private Const MAX_BATCH_SIZE As Long = 1000         ' 最大バッチサイズ
Private Const CACHE_DURATION_MS As Long = 300000    ' キャッシュの有効期間（5分）
Private Const MAX_RETRY_COUNT As Long = 3           ' 最大リトライ回数
Private Const BACKUP_INTERVAL_MS As Long = 3600000  ' バックアップ間隔（1時間）
Private Const METRICS_INTERVAL_MS As Long = 60000   ' メトリクス収集間隔（1分）
Private Const CLEANUP_INTERVAL_MS As Long = 300000  ' クリーンアップ間隔（5分）

' イベント定義
Public Event EventLogged(ByVal source As String, ByVal eventType As EventLogType)
Public Event BatchProcessed(ByVal source As String, ByVal successCount As Long, ByVal failureCount As Long)
Public Event SourceRegistered(ByVal source As String, ByVal success As Boolean)
Public Event BackupCreated(ByVal backupPath As String)
Public Event FilterApplied(ByVal filterCriteria As String, ByVal matchCount As Long)
Public Event SecurityAlert(ByVal alertType As String, ByVal details As String)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)
Public Event ResourceExhausted(ByVal resourceType As String, ByVal details As String)

' メンバ変数
Private mIsInitialized As Boolean
Private mLock As clsLock
Private mPerformanceMonitor As clsPerformanceMonitor
Private mLastError As String
Private mSourceCache As Object ' Scripting.Dictionary
Private mLastBackupTime As Date
Private mLastMetricsTime As Date
Private mLastCleanupTime As Date
Private mEventFilters As Collection
Private mSecurityContext As SecurityContext

' 初期化・終了処理
Private Sub Class_Initialize()
    Set mLock = New clsLock
    Set mPerformanceMonitor = New clsPerformanceMonitor
    Set mSourceCache = CreateObject("Scripting.Dictionary")
    Set mEventFilters = New Collection
    Set mSecurityContext = New SecurityContext
    mLastBackupTime = Now
    mLastMetricsTime = Now
    mLastCleanupTime = Now
    mIsInitialized = True
End Sub

Private Sub Class_Terminate()
    Set mLock = Nothing
    Set mPerformanceMonitor = Nothing
    Set mSourceCache = Nothing
    Set mEventFilters = Nothing
    Set mSecurityContext = Nothing
    mIsInitialized = False
End Sub

' IEventLog実装
Private Function IEventLog_WriteToEventLog(ByVal source As String, _
                                         ByVal message As String, _
                                         ByVal eventType As EventLogType, _
                                         Optional ByVal options As WriteOptions) As Boolean
    If Not mIsInitialized Then Exit Function
    
    mPerformanceMonitor.StartMeasurement "WriteToEventLog"
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    ' セキュリティチェック
    If Not ValidateSecurityContext(source) Then
        LogSecurityAlert "不正なイベントログアクセス試行", source
        GoTo CleanUp
    End If
    
    ' ソースの存在確認（キャッシュ利用）
    If Not VerifyEventSource(source) Then
        LogError "イベントソースが存在しません: " & source
        GoTo CleanUp
    End If
    
    ' メッセージの検証と整形
    message = ValidateAndFormatMessage(message)
    If message = "" Then GoTo CleanUp
    
    ' イベントログ書き込み（リトライ機能付き）
    If WriteEventLogWithRetry(source, message, eventType, options) Then
        RaiseEvent EventLogged(source, eventType)
        IEventLog_WriteToEventLog = True
    End If
    
    ' 定期的な処理
    PerformPeriodicTasks
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "WriteToEventLog"
    Exit Function

ErrorHandler:
    mLastError = "イベントログへの書き込み中にエラーが発生しました: " & Err.Description
    LogError mLastError
    IEventLog_WriteToEventLog = False
    Resume CleanUp
End Function

Private Function IEventLog_WriteBatch(ByVal entries As Collection, Optional ByVal options As BatchOptions) As BatchResult
    Dim result As New BatchResult
    
    If entries.Count > MAX_BATCH_SIZE Then
        LogError "バッチサイズが上限を超えています: " & entries.Count
        Set IEventLog_WriteBatch = result
        Exit Function
    End If
    
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    Dim entry As EventLogEntry
    For Each entry In entries
        If IEventLog_WriteToEventLog(entry.Source, entry.Message, entry.EventType) Then
            result.SuccessCount = result.SuccessCount + 1
        Else
            result.FailureCount = result.FailureCount + 1
        End If
    Next
    
    RaiseEvent BatchProcessed(entries(1).Source, result.SuccessCount, result.FailureCount)
    
CleanUp:
    mLock.ReleaseLock
    Set IEventLog_WriteBatch = result
    Exit Function

ErrorHandler:
    LogError "バッチ処理中にエラーが発生しました: " & Err.Description
    Resume CleanUp
End Function

Private Function IEventLog_FilterEvents(ByVal criteria As String, Optional ByVal options As FilterOptions) As Collection
    Dim results As New Collection
    
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    ' フィルタリング処理
    ApplyEventFilters results, criteria, options
    
    RaiseEvent FilterApplied(criteria, results.Count)
    
CleanUp:
    mLock.ReleaseLock
    Set IEventLog_FilterEvents = results
    Exit Function

ErrorHandler:
    LogError "イベントのフィルタリング中にエラーが発生しました: " & Err.Description
    Resume CleanUp
End Function

Private Function IEventLog_CreateBackup(Optional ByVal options As BackupOptions) As String
    Dim backupPath As String
    
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    backupPath = CreateEventLogBackup(options)
    If backupPath <> "" Then
        RaiseEvent BackupCreated(backupPath)
        mLastBackupTime = Now
    End If
    
CleanUp:
    mLock.ReleaseLock
    IEventLog_CreateBackup = backupPath
    Exit Function

ErrorHandler:
    LogError "バックアップ作成中にエラーが発生しました: " & Err.Description
    Resume CleanUp
End Function

Private Function IEventLog_GetPerformanceMetrics() As Collection
    Dim metrics As New Collection
    
    With metrics
        .Add mPerformanceMonitor.GetAllMeasurements()
        .Add GetEventLogMetrics()
        .Add GetCacheMetrics()
        .Add GetSecurityMetrics()
    End With
    
    Set IEventLog_GetPerformanceMetrics = metrics
End Function

' プライベートメソッド
Private Function ValidateSecurityContext(ByVal source As String) As Boolean
    ' セキュリティコンテキストの検証
    ' 実装は省略
End Function

Private Sub LogSecurityAlert(ByVal alertType As String, ByVal details As String)
    RaiseEvent SecurityAlert(alertType, details)
    ' 実装は省略
End Sub

Private Function VerifyEventSource(ByVal source As String) As Boolean
    ' イベントソースの検証（キャッシュ利用）
    ' 実装は省略
End Function

Private Function ValidateAndFormatMessage(ByVal message As String) As String
    ' メッセージの検証と整形
    ' 実装は省略
End Function

Private Function WriteEventLogWithRetry(ByVal source As String, _
                                      ByVal message As String, _
                                      ByVal eventType As EventLogType, _
                                      ByVal options As WriteOptions) As Boolean
    ' リトライ機能付きイベントログ書き込み
    ' 実装は省略
End Function

Private Sub PerformPeriodicTasks()
    ' メトリクス収集
    If DateDiff("s", mLastMetricsTime, Now) > METRICS_INTERVAL_MS / 1000 Then
        CollectMetrics
        mLastMetricsTime = Now
    End If
    
    ' バックアップ
    If DateDiff("s", mLastBackupTime, Now) > BACKUP_INTERVAL_MS / 1000 Then
        IEventLog_CreateBackup
    End If
    
    ' クリーンアップ
    If DateDiff("s", mLastCleanupTime, Now) > CLEANUP_INTERVAL_MS / 1000 Then
        PerformCleanup
        mLastCleanupTime = Now
    End If
End Sub

Private Sub CollectMetrics()
    ' メトリクス収集処理
    ' 実装は省略
End Sub

Private Sub PerformCleanup()
    ' クリーンアップ処理
    ' 実装は省略
End Sub

Private Function CreateEventLogBackup(ByVal options As BackupOptions) As String
    ' バックアップ作成処理
    ' 実装は省略
End Function

Private Sub ApplyEventFilters(ByRef results As Collection, _
                            ByVal criteria As String, _
                            ByVal options As FilterOptions)
    ' フィルタリング処理
    ' 実装は省略
End Sub

Private Function GetEventLogMetrics() As Collection
    ' イベントログのメトリクス取得
    ' 実装は省略
End Function

Private Function GetCacheMetrics() As Collection
    ' キャッシュのメトリクス取得
    ' 実装は省略
End Function

Private Function GetSecurityMetrics() As Collection
    ' セキュリティのメトリクス取得
    ' 実装は省略
End Function

' テストサポート機能
#If DEBUG Then
    Public Function ValidateState() As Boolean
        ValidateState = mIsInitialized And _
                       Not (mLock Is Nothing) And _
                       Not (mPerformanceMonitor Is Nothing) And _
                       Not (mSourceCache Is Nothing) And _
                       Not (mEventFilters Is Nothing) And _
                       Not (mSecurityContext Is Nothing)
    End Function
    
    Public Function GetPerformanceMonitor() As clsPerformanceMonitor
        Set GetPerformanceMonitor = mPerformanceMonitor
    End Function
    
    Public Function TestEventLogAccess(ByVal testSource As String) As Boolean
        TestEventLogAccess = IEventLog_WriteToEventLog(testSource, _
                                                      "EventLogImpl Test Message @ " & Now, _
                                                      EVENTLOG_INFORMATION_TYPE)
    End Function
#End If