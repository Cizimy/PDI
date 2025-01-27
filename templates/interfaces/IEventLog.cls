VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEventLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

'@Description("Windowsイベントログの操作と管理を提供するインターフェース")

' 定数定義
Private Const MAX_MESSAGE_LENGTH As Long = 32766   ' メッセージの最大長
Private Const MAX_SOURCE_LENGTH As Long = 255      ' ソース名の最大長
Private Const MAX_BATCH_SIZE As Long = 1000        ' 最大バッチサイズ
Private Const CACHE_DURATION_MS As Long = 300000   ' キャッシュの有効期間（5分）
Private Const MAX_RETRY_COUNT As Long = 3          ' 最大リトライ回数
Private Const CLEANUP_INTERVAL_MS As Long = 3600000 ' クリーンアップ間隔（1時間）

' イベントログの種類定義
Public Enum EventLogType
    EVENTLOG_SUCCESS = 0
    EVENTLOG_ERROR_TYPE = 1
    EVENTLOG_WARNING_TYPE = 2
    EVENTLOG_INFORMATION_TYPE = 4
    EVENTLOG_AUDIT_SUCCESS = 8
    EVENTLOG_AUDIT_FAILURE = 16
End Enum

' イベント定義
Public Event EventLogged(ByVal source As String, ByVal eventType As EventLogType)
Public Event SourceRegistered(ByVal source As String, ByVal success As Boolean)
Public Event LogCleared(ByVal source As String, ByVal entryCount As Long)
Public Event BatchProcessed(ByVal source As String, ByVal successCount As Long, ByVal failureCount As Long)
Public Event LogFull(ByVal source As String, ByVal usedSpace As Long)
Public Event BackupCreated(ByVal backupPath As String)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)

''' <summary>
''' イベントログにメッセージを書き込みます
''' </summary>
''' <param name="source">イベントソース名</param>
''' <param name="message">メッセージ</param>
''' <param name="eventType">イベントの種類</param>
''' <param name="options">書き込みオプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function WriteToEventLog(ByVal source As String, _
                              ByVal message As String, _
                              ByVal eventType As EventLogType, _
                              Optional ByVal options As WriteOptions) As Boolean
End Function

''' <summary>
''' バッチでイベントログに書き込みます
''' </summary>
''' <param name="entries">イベントログエントリのコレクション</param>
''' <param name="options">バッチオプション（オプション）</param>
''' <returns>バッチ処理結果</returns>
Public Function WriteBatch(ByVal entries As Collection, _
                         Optional ByVal options As BatchOptions) As BatchResult
End Function

''' <summary>
''' イベントソースの存在を確認します
''' </summary>
''' <param name="source">イベントソース名</param>
''' <param name="options">確認オプション（オプション）</param>
''' <returns>存在する場合True、存在しない場合False</returns>
Public Function EventSourceExists(ByVal source As String, _
                                Optional ByVal options As SourceOptions) As Boolean
End Function

''' <summary>
''' イベントソースを登録します
''' </summary>
''' <param name="source">イベントソース名</param>
''' <param name="options">登録オプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function RegisterEventSource(ByVal source As String, _
                                  Optional ByVal options As RegisterOptions) As Boolean
End Function

''' <summary>
''' イベントログをクリアします
''' </summary>
''' <param name="source">イベントソース名</param>
''' <param name="options">クリアオプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function ClearEventLog(ByVal source As String, _
                            Optional ByVal options As ClearOptions) As Boolean
End Function

''' <summary>
''' イベントログをバックアップします
''' </summary>
''' <param name="backupPath">バックアップパス</param>
''' <param name="options">バックアップオプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function BackupEventLog(ByVal backupPath As String, _
                             Optional ByVal options As BackupOptions) As Boolean
End Function

''' <summary>
''' イベントログをフィルタリングします
''' </summary>
''' <param name="criteria">フィルタ条件</param>
''' <param name="options">フィルタオプション（オプション）</param>
''' <returns>フィルタリングされたエントリのコレクション</returns>
Public Function FilterEvents(ByVal criteria As String, _
                           Optional ByVal options As FilterOptions) As Collection
End Function

''' <summary>
''' イベントログの統計を取得します
''' </summary>
''' <param name="source">イベントソース名</param>
''' <param name="options">統計オプション（オプション）</param>
''' <returns>統計情報</returns>
Public Function GetEventStatistics(ByVal source As String, _
                                 Optional ByVal options As StatisticsOptions) As EventStatistics
End Function

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <returns>パフォーマンス情報のコレクション</returns>
Public Function GetPerformanceMetrics() As Collection
End Function

''' <summary>
''' イベントログの状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function

''' <summary>
''' リソースを解放します
''' </summary>
Public Sub Cleanup()
End Sub