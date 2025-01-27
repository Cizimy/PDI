VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPerformanceCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

'@Description("高精度なパフォーマンス測定と監視機能を提供するインターフェース")

' 定数定義
Private Const MAX_SAMPLES As Long = 1000            ' 最大サンプル数
Private Const MIN_INTERVAL_MS As Long = 100         ' 最小測定間隔（ミリ秒）
Private Const CACHE_DURATION_MS As Long = 1000      ' キャッシュの有効期間（1秒）
Private Const DEFAULT_THRESHOLD As Double = 0.8      ' デフォルトの閾値（80%）
Private Const MAX_HISTORY_ENTRIES As Long = 1000    ' 履歴の最大エントリ数
Private Const CLEANUP_INTERVAL_MS As Long = 60000   ' クリーンアップ間隔（1分）

' イベント定義
Public Event CounterStarted(ByVal counterName As String, ByVal startTime As Date)
Public Event CounterStopped(ByVal counterName As String, ByVal elapsedTime As Double)
Public Event ThresholdExceeded(ByVal counterName As String, ByVal currentValue As Double, ByVal threshold As Double)
Public Event SampleCollected(ByVal counterName As String, ByVal value As Double)
Public Event StatisticsUpdated(ByVal counterName As String, ByVal statistics As PerformanceStatistics)
Public Event AnomalyDetected(ByVal counterName As String, ByVal anomalyType As String, ByVal value As Double)
Public Event CacheUpdated(ByVal counterName As String, ByVal isCached As Boolean)

''' <summary>
''' パフォーマンスカウンターの現在値を取得します
''' </summary>
''' <param name="performanceCount">カウンター値を格納する変数</param>
''' <param name="options">取得オプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
''' <remarks>
''' エラー処理要件：
''' - カウンターの有効性確認
''' - オーバーフロー検出
''' - プロセッサ間の一貫性確認
''' - キャッシュの活用
''' </remarks>
Public Function QueryPerformanceCounter(ByRef performanceCount As Currency, Optional ByVal options As CounterOptions) As Boolean
End Function

''' <summary>
''' パフォーマンスカウンターの周波数を取得します
''' </summary>
''' <param name="frequency">周波数を格納する変数</param>
''' <param name="options">取得オプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
''' <remarks>
''' エラー処理要件：
''' - システムサポートの確認
''' - 値の安定性確認
''' - キャッシュの活用
''' </remarks>
Public Function QueryPerformanceFrequency(ByRef frequency As Currency, Optional ByVal options As FrequencyOptions) As Boolean
End Function

''' <summary>
''' パフォーマンス測定を開始します
''' </summary>
''' <param name="counterName">カウンター名</param>
''' <param name="options">測定オプション（オプション）</param>
''' <returns>測定ID</returns>
Public Function StartMeasurement(ByVal counterName As String, Optional ByVal options As MeasurementOptions) As String
End Function

''' <summary>
''' パフォーマンス測定を停止します
''' </summary>
''' <param name="measurementId">測定ID</param>
''' <returns>測定結果</returns>
Public Function StopMeasurement(ByVal measurementId As String) As MeasurementResult
End Function

''' <summary>
''' パフォーマンス閾値を設定します
''' </summary>
''' <param name="counterName">カウンター名</param>
''' <param name="threshold">閾値</param>
''' <param name="options">閾値オプション（オプション）</param>
Public Sub SetThreshold(ByVal counterName As String, ByVal threshold As Double, Optional ByVal options As ThresholdOptions)
End Sub

''' <summary>
''' パフォーマンス統計を取得します
''' </summary>
''' <param name="counterName">カウンター名</param>
''' <param name="options">統計オプション（オプション）</param>
''' <returns>統計情報</returns>
Public Function GetStatistics(ByVal counterName As String, Optional ByVal options As StatisticsOptions) As PerformanceStatistics
End Function

''' <summary>
''' パフォーマンス履歴を取得します
''' </summary>
''' <param name="counterName">カウンター名</param>
''' <param name="options">履歴オプション（オプション）</param>
''' <returns>履歴データのコレクション</returns>
Public Function GetHistory(ByVal counterName As String, Optional ByVal options As HistoryOptions) As Collection
End Function

''' <summary>
''' 異常値を検出します
''' </summary>
''' <param name="counterName">カウンター名</param>
''' <param name="options">検出オプション（オプション）</param>
''' <returns>異常値のコレクション</returns>
Public Function DetectAnomalies(ByVal counterName As String, Optional ByVal options As AnomalyOptions) As Collection
End Function

''' <summary>
''' パフォーマンスレポートを生成します
''' </summary>
''' <param name="options">レポートオプション（オプション）</param>
''' <returns>レポートデータ</returns>
Public Function GenerateReport(Optional ByVal options As ReportOptions) As PerformanceReport
End Function

''' <summary>
''' カウンターをリセットします
''' </summary>
''' <param name="counterName">カウンター名（オプション）</param>
Public Sub Reset(Optional ByVal counterName As String)
End Sub

''' <summary>
''' カウンターの状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function

''' <summary>
''' リソースを解放します
''' </summary>
Public Sub Cleanup()
End Sub