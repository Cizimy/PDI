VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

'@Description("ログの記録、管理、分析を提供し、複数の出力先とフォーマットをサポートするインターフェース")

' 定数定義
Private Const MAX_MESSAGE_LENGTH As Long = 8192     ' メッセージの最大長
Private Const MAX_QUEUE_SIZE As Long = 10000        ' キューの最大サイズ
Private Const FLUSH_INTERVAL_MS As Long = 1000      ' フラッシュ間隔（1秒）
Private Const MAX_BATCH_SIZE As Long = 1000         ' バッチサイズの最大値
Private Const MAX_FILE_SIZE_MB As Long = 100        ' ログファイルの最大サイズ（MB）
Private Const MAX_RETENTION_DAYS As Long = 90       ' ログの最大保持期間（日）
Private Const COMPRESSION_THRESHOLD_MB As Long = 50  ' 圧縮開始サイズ（MB）

' イベント定義
Public Event LogEntryWritten(ByVal entry As LogEntry, ByVal destination As LogDestination)
Public Event LogLevelChanged(ByVal oldLevel As LogLevel, ByVal newLevel As LogLevel)
Public Event DestinationChanged(ByVal oldDestination As LogDestination, ByVal newDestination As LogDestination)
Public Event LogRotated(ByVal oldFile As String, ByVal newFile As String)
Public Event LogCompressed(ByVal originalSize As Long, ByVal compressedSize As Long)
Public Event QueueThresholdReached(ByVal queueSize As Long, ByVal threshold As Long)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)
Public Event StorageAlert(ByVal availableSpace As Long, ByVal threshold As Long)

''' <summary>
''' ログを記録します
''' </summary>
''' <param name="moduleName">モジュール名</param>
''' <param name="message">ログメッセージ</param>
''' <param name="level">ログレベル（オプション）</param>
''' <param name="errorCode">エラーコード（オプション）</param>
''' <param name="options">ログオプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - パラメータの検証（NULL、空文字、最大長など）
''' - 出力先の状態確認
''' - エラー発生時の代替処理
''' - リトライメカニズムの実装
''' - キューの管理
''' - パフォーマンスの監視
''' </remarks>
Public Sub Log(ByVal moduleName As String, _
               ByVal message As String, _
               Optional ByVal level As LogLevel = llInfo, _
               Optional ByVal errorCode As ErrorCode, _
               Optional ByVal options As LogOptions)
End Sub

''' <summary>
''' バッチログを記録します
''' </summary>
''' <param name="entries">ログエントリのコレクション</param>
''' <param name="options">バッチオプション（オプション）</param>
Public Sub LogBatch(ByVal entries As Collection, Optional ByVal options As BatchLogOptions)
End Sub

''' <summary>
''' ログレベルを設定します
''' </summary>
''' <param name="level">ログレベル</param>
''' <param name="options">レベルオプション（オプション）</param>
Public Property Let LogLevel(ByVal level As LogLevel, Optional ByVal options As LevelOptions)
End Property

''' <summary>
''' 現在のログレベルを取得します
''' </summary>
Public Property Get LogLevel() As LogLevel
End Property

''' <summary>
''' ログの出力先を設定します
''' </summary>
''' <param name="destination">ログの出力先</param>
''' <param name="options">出力先オプション（オプション）</param>
Public Property Let LogDestination(ByVal destination As LogDestination, Optional ByVal options As DestinationOptions)
End Property

''' <summary>
''' 現在のログ出力先を取得します
''' </summary>
Public Property Get LogDestination() As LogDestination
End Property

''' <summary>
''' ログフィルターを設定します
''' </summary>
''' <param name="filter">フィルター条件</param>
''' <param name="options">フィルターオプション（オプション）</param>
Public Sub SetLogFilter(ByVal filter As String, Optional ByVal options As FilterOptions)
End Sub

''' <summary>
''' ログをフラッシュします
''' </summary>
''' <param name="options">フラッシュオプション（オプション）</param>
Public Sub Flush(Optional ByVal options As FlushOptions)
End Sub

''' <summary>
''' ログをローテーションします
''' </summary>
''' <param name="options">ローテーションオプション（オプション）</param>
Public Sub RotateLog(Optional ByVal options As RotationOptions)
End Sub

''' <summary>
''' ログを圧縮します
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="options">圧縮オプション（オプション）</param>
Public Sub CompressLog(ByVal filePath As String, Optional ByVal options As CompressionOptions)
End Sub

''' <summary>
''' ログを検索します
''' </summary>
''' <param name="searchCriteria">検索条件</param>
''' <param name="options">検索オプション（オプション）</param>
''' <returns>検索結果のコレクション</returns>
Public Function SearchLogs(ByVal searchCriteria As String, Optional ByVal options As SearchOptions) As Collection
End Function

''' <summary>
''' ログの統計情報を取得します
''' </summary>
''' <param name="options">統計オプション（オプション）</param>
''' <returns>統計情報</returns>
Public Function GetLogStatistics(Optional ByVal options As StatisticsOptions) As LogStatistics
End Function

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <returns>パフォーマンス情報のコレクション</returns>
Public Function GetPerformanceMetrics() As Collection
End Function

''' <summary>
''' ログの状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function

''' <summary>
''' リソースを解放します
''' </summary>
Public Sub Cleanup()
End Sub