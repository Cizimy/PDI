VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPerformanceMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("システム全体のパフォーマンスを監視し、分析、レポート生成を提供するインターフェース")

' 定数定義
Private Const MAX_MEASUREMENTS As Long = 1000       ' 最大測定数
Private Const SAMPLING_INTERVAL_MS As Long = 1000   ' サンプリング間隔（1秒）
Private Const ALERT_THRESHOLD As Double = 0.8       ' アラート閾値（80%）
Private Const MAX_HISTORY_DAYS As Long = 30         ' 履歴保持期間（日）
Private Const CLEANUP_INTERVAL_MS As Long = 3600000 ' クリーンアップ間隔（1時間）
Private Const DEFAULT_TIMEOUT_MS As Long = 30000    ' デフォルトタイムアウト（30秒）

' イベント定義
Public Event MeasurementStarted(ByVal measurementName As String, ByVal startTime As Date)
Public Event MeasurementCompleted(ByVal measurementName As String, ByVal duration As Double)
Public Event ThresholdExceeded(ByVal metricName As String, ByVal currentValue As Double, ByVal threshold As Double)
Public Event ResourceAlert(ByVal resourceType As String, ByVal usage As Double, ByVal capacity As Double)
Public Event PerformanceAnomaly(ByVal metricName As String, ByVal value As Double, ByVal expectedRange As String)
Public Event MonitoringStatusChanged(ByVal isEnabled As Boolean, ByVal reason As String)
Public Event ReportGenerated(ByVal reportType As String, ByVal timestamp As Date)

''' <summary>
''' 計測を開始します
''' </summary>
''' <param name="measurementName">計測名</param>
''' <param name="category">カテゴリ（オプション）</param>
''' <param name="options">計測オプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - 計測名の検証
''' - 重複計測の検出
''' - リソース使用量の確認
''' - タイムアウトの設定
''' </remarks>
Public Sub Start(ByVal measurementName As String, Optional ByVal category As String, Optional ByVal options As MeasurementOptions)
End Sub

''' <summary>
''' 計測を終了します
''' </summary>
''' <param name="measurementName">計測名</param>
''' <param name="options">終了オプション（オプション）</param>
''' <returns>計測結果</returns>
Public Function Stop(ByVal measurementName As String, Optional ByVal options As StopOptions) As MeasurementResult
End Function

''' <summary>
''' 指定した計測の結果を取得します
''' </summary>
''' <param name="measurementName">計測名</param>
''' <param name="options">取得オプション（オプション）</param>
''' <returns>計測結果</returns>
Public Function GetMeasurement(ByVal measurementName As String, Optional ByVal options As GetOptions) As MeasurementResult
End Function

''' <summary>
''' すべての計測結果を取得します
''' </summary>
''' <param name="filter">フィルター条件（オプション）</param>
''' <param name="options">取得オプション（オプション）</param>
''' <returns>計測結果のコレクション</returns>
Public Function GetAllMeasurements(Optional ByVal filter As String, Optional ByVal options As GetAllOptions) As Collection
End Function

''' <summary>
''' パフォーマンス統計を取得します
''' </summary>
''' <param name="metricName">メトリクス名</param>
''' <param name="options">統計オプション（オプション）</param>
''' <returns>統計情報</returns>
Public Function GetStatistics(ByVal metricName As String, Optional ByVal options As StatisticsOptions) As PerformanceStatistics
End Function

''' <summary>
''' リソース使用状況を監視します
''' </summary>
''' <param name="resourceType">リソースタイプ</param>
''' <param name="options">監視オプション（オプション）</param>
Public Sub MonitorResource(ByVal resourceType As String, Optional ByVal options As MonitorOptions)
End Sub

''' <summary>
''' パフォーマンスレポートを生成します
''' </summary>
''' <param name="reportType">レポートタイプ</param>
''' <param name="options">レポートオプション（オプション）</param>
''' <returns>レポートデータ</returns>
Public Function GenerateReport(ByVal reportType As String, Optional ByVal options As ReportOptions) As PerformanceReport
End Function

''' <summary>
''' アラート条件を設定します
''' </summary>
''' <param name="metricName">メトリクス名</param>
''' <param name="condition">アラート条件</param>
''' <param name="options">アラートオプション（オプション）</param>
Public Sub SetAlertCondition(ByVal metricName As String, ByVal condition As String, Optional ByVal options As AlertOptions)
End Sub

''' <summary>
''' すべての計測結果をクリアします
''' </summary>
''' <param name="options">クリアオプション（オプション）</param>
Public Sub Clear(Optional ByVal options As ClearOptions)
End Sub

''' <summary>
''' パフォーマンス監視が有効かどうかを取得します
''' </summary>
Public Property Get IsEnabled() As Boolean
End Property

''' <summary>
''' パフォーマンス監視を有効にします
''' </summary>
''' <param name="options">有効化オプション（オプション）</param>
Public Sub Enable(Optional ByVal options As EnableOptions)
End Sub

''' <summary>
''' パフォーマンス監視を無効にします
''' </summary>
''' <param name="options">無効化オプション（オプション）</param>
Public Sub Disable(Optional ByVal options As DisableOptions)
End Sub

''' <summary>
''' 監視設定をエクスポートします
''' </summary>
''' <param name="filePath">出力ファイルパス</param>
''' <param name="options">エクスポートオプション（オプション）</param>
Public Sub ExportSettings(ByVal filePath As String, Optional ByVal options As ExportOptions)
End Sub

''' <summary>
''' 監視設定をインポートします
''' </summary>
''' <param name="filePath">入力ファイルパス</param>
''' <param name="options">インポートオプション（オプション）</param>
Public Sub ImportSettings(ByVal filePath As String, Optional ByVal options As ImportOptions)
End Sub

''' <summary>
''' 監視の状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function

''' <summary>
''' リソースを解放します
''' </summary>
Public Sub Cleanup()
End Sub