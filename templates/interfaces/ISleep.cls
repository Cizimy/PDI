VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISleep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

'@Description("高精度なスリープ機能とスケジューリングを提供するインターフェース")

' 定数定義
Private Const MIN_SLEEP_MS As Long = 1             ' 最小スリープ時間（ミリ秒）
Private Const MAX_SLEEP_MS As Long = 3600000       ' 最大スリープ時間（1時間）
Private Const TIMER_RESOLUTION_MS As Long = 1       ' タイマー分解能（ミリ秒）
Private Const MAX_RETRIES As Long = 3              ' 最大リトライ回数
Private Const SPIN_THRESHOLD_MS As Long = 15       ' スピンウェイト閾値（ミリ秒）
Private Const DEFAULT_TIMEOUT_MS As Long = 30000   ' デフォルトタイムアウト（30秒）
Private Const MAX_BATCH_SIZE As Long = 100         ' 最大バッチサイズ
Private Const MAX_GROUP_SIZE As Long = 10          ' 最大グループサイズ
Private Const ENERGY_CHECK_INTERVAL_MS As Long = 60000 ' エネルギーチェック間隔（1分）
Private Const CLEANUP_INTERVAL_MS As Long = 300000 ' クリーンアップ間隔（5分）

' イベント定義
Public Event SleepStarted(ByVal duration As Long, ByVal startTime As Date)
Public Event SleepCompleted(ByVal duration As Long, ByVal actualDuration As Double)
Public Event SleepInterrupted(ByVal duration As Long, ByVal elapsedTime As Double)
Public Event TimerResolutionChanged(ByVal oldResolution As Long, ByVal newResolution As Long)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)
Public Event ScheduleRegistered(ByVal scheduleId As String, ByVal nextExecutionTime As Date)
Public Event ResourceStateChanged(ByVal resourceType As String, ByVal state As String)
Public Event BatchCompleted(ByVal batchId As String, ByVal successCount As Long)
Public Event GroupStateChanged(ByVal groupId As String, ByVal state As String)
Public Event EnergyStateChanged(ByVal oldState As String, ByVal newState As String)

''' <summary>
''' 指定された時間だけ実行を停止します
''' </summary>
''' <param name="milliseconds">スリープ時間（ミリ秒）</param>
''' <param name="options">スリープオプション（オプション）</param>
Public Sub Sleep(ByVal milliseconds As Long, Optional ByVal options As SleepOptions)
End Sub

''' <summary>
''' 高精度スリープを実行します
''' </summary>
''' <param name="milliseconds">スリープ時間（ミリ秒）</param>
''' <param name="options">高精度オプション（オプション）</param>
Public Sub HighPrecisionSleep(ByVal milliseconds As Long, Optional ByVal options As PrecisionOptions)
End Sub

''' <summary>
''' スピンウェイトを実行します
''' </summary>
''' <param name="milliseconds">待機時間（ミリ秒）</param>
''' <param name="options">スピンオプション（オプション）</param>
Public Sub SpinWait(ByVal milliseconds As Long, Optional ByVal options As SpinOptions)
End Sub

''' <summary>
''' スリープをスケジュールします
''' </summary>
''' <param name="milliseconds">スリープ時間（ミリ秒）</param>
''' <param name="scheduleTime">スケジュール時刻</param>
''' <param name="options">スケジュールオプション（オプション）</param>
''' <returns>スケジュールID</returns>
Public Function ScheduleSleep(ByVal milliseconds As Long, ByVal scheduleTime As Date, Optional ByVal options As ScheduleOptions) As String
End Function

''' <summary>
''' 定期的なスリープをスケジュールします
''' </summary>
''' <param name="milliseconds">スリープ時間（ミリ秒）</param>
''' <param name="interval">実行間隔（ミリ秒）</param>
''' <param name="options">定期実行オプション（オプション）</param>
''' <returns>スケジュールID</returns>
Public Function ScheduleRecurringSleep(ByVal milliseconds As Long, ByVal interval As Long, Optional ByVal options As RecurringOptions) As String
End Function

''' <summary>
''' スリープパターンを登録します
''' </summary>
''' <param name="pattern">スリープパターン</param>
''' <param name="options">パターンオプション（オプション）</param>
''' <returns>パターンID</returns>
Public Function RegisterPattern(ByVal pattern As SleepPattern, Optional ByVal options As PatternOptions) As String
End Function

''' <summary>
''' バッチスリープを実行します
''' </summary>
''' <param name="durations">スリープ時間のコレクション</param>
''' <param name="options">バッチオプション（オプション）</param>
''' <returns>バッチ実行結果</returns>
Public Function BatchSleep(ByVal durations As Collection, Optional ByVal options As BatchOptions) As BatchResult
End Function

''' <summary>
''' スリープグループを作成します
''' </summary>
''' <param name="name">グループ名</param>
''' <param name="options">グループオプション（オプション）</param>
''' <returns>グループID</returns>
Public Function CreateGroup(ByVal name As String, Optional ByVal options As GroupOptions) As String
End Function

''' <summary>
''' スリープをグループに追加します
''' </summary>
''' <param name="groupId">グループID</param>
''' <param name="milliseconds">スリープ時間（ミリ秒）</param>
''' <param name="options">追加オプション（オプション）</param>
Public Sub AddToGroup(ByVal groupId As String, ByVal milliseconds As Long, Optional ByVal options As AddOptions)
End Sub

''' <summary>
''' スケジュールされたスリープをキャンセルします
''' </summary>
''' <param name="scheduleId">スケジュールID</param>
Public Sub CancelScheduledSleep(ByVal scheduleId As String)
End Sub

''' <summary>
''' 現在のスリープを中断します
''' </summary>
Public Sub Interrupt()
End Sub

''' <summary>
''' エネルギー設定を構成します
''' </summary>
''' <param name="settings">エネルギー設定</param>
Public Sub ConfigureEnergy(ByVal settings As EnergySettings)
End Sub

''' <summary>
''' タイマーの分解能を設定します
''' </summary>
''' <param name="resolution">分解能（ミリ秒）</param>
Public Property Let TimerResolution(ByVal resolution As Long)
End Property

''' <summary>
''' タイマーの分解能を取得します
''' </summary>
Public Property Get TimerResolution() As Long
End Property

''' <summary>
''' スリープ状態を取得します
''' </summary>
Public Property Get State() As SleepState
End Property

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <returns>パフォーマンス情報のコレクション</returns>
Public Function GetPerformanceMetrics() As Collection
End Function

''' <summary>
''' スケジュール一覧を取得します
''' </summary>
''' <returns>スケジュール情報のコレクション</returns>
Public Function GetSchedules() As Collection
End Function

''' <summary>
''' グループ一覧を取得します
''' </summary>
''' <returns>グループ情報のコレクション</returns>
Public Function GetGroups() As Collection
End Function

''' <summary>
''' エネルギー使用状況を取得します
''' </summary>
''' <returns>エネルギー情報</returns>
Public Function GetEnergyUsage() As EnergyUsage
End Function

''' <summary>
''' スリープの状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function

''' <summary>
''' リソースを解放します
''' </summary>
Public Sub Cleanup()
End Sub