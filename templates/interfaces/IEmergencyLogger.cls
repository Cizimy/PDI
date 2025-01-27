VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEmergencyLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("緊急時のエラーログ出力とフォールバック機能を提供するインターフェース")

' 定数定義
Private Const MAX_BUFFER_SIZE As Long = 1000        ' バッファの最大サイズ
Private Const FLUSH_INTERVAL_MS As Long = 1000      ' フラッシュ間隔（1秒）
Private Const MAX_RETRY_COUNT As Long = 3           ' 最大リトライ回数
Private Const MAX_MESSAGE_LENGTH As Long = 8192     ' メッセージの最大長
Private Const FALLBACK_PATH As String = "emergency.log" ' フォールバックログパス
Private Const ALERT_THRESHOLD As Long = 100         ' アラート閾値
Private Const CLEANUP_INTERVAL_MS As Long = 3600000 ' クリーンアップ間隔（1時間）

' イベント定義
Public Event EmergencyLogWritten(ByVal message As String, ByVal severity As EmergencyLevel)
Public Event FallbackActivated(ByVal reason As String, ByVal fallbackPath As String)
Public Event BufferThresholdReached(ByVal currentSize As Long, ByVal maxSize As Long)
Public Event LogFlushed(ByVal entryCount As Long, ByVal success As Boolean)
Public Event NotificationSent(ByVal recipient As String, ByVal messageType As String)
Public Event ResourceExhausted(ByVal resourceType As String, ByVal details As String)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)

''' <summary>
''' 緊急時のエラーをログに記録します
''' </summary>
''' <param name="message">エラーメッセージ</param>
''' <param name="errorInfo">エラー情報（オプション）</param>
''' <param name="options">ログオプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - メッセージの検証
''' - バッファの管理
''' - フォールバックの適用
''' - リソースの監視
''' - 優先度の確認
''' </remarks>
Public Sub LogEmergencyError(ByVal message As String, Optional ByRef errorInfo As ErrorInfo, Optional ByVal options As EmergencyLogOptions)
End Sub

''' <summary>
''' バッファをフラッシュします
''' </summary>
''' <param name="options">フラッシュオプション（オプション）</param>
Public Sub FlushBuffer(Optional ByVal options As FlushOptions)
End Sub

''' <summary>
''' フォールバックパスを設定します
''' </summary>
''' <param name="path">フォールバックパス</param>
''' <param name="options">パスオプション（オプション）</param>
Public Sub SetFallbackPath(ByVal path As String, Optional ByVal options As FallbackOptions)
End Sub

''' <summary>
''' 緊急通知を送信します
''' </summary>
''' <param name="message">通知メッセージ</param>
''' <param name="recipients">受信者リスト</param>
''' <param name="options">通知オプション（オプション）</param>
Public Sub SendEmergencyNotification(ByVal message As String, ByVal recipients As Collection, Optional ByVal options As NotificationOptions)
End Sub

''' <summary>
''' ログの優先度を設定します
''' </summary>
''' <param name="level">優先度レベル</param>
''' <param name="options">優先度オプション（オプション）</param>
Public Sub SetPriorityLevel(ByVal level As EmergencyLevel, Optional ByVal options As PriorityOptions)
End Sub

''' <summary>
''' バッファサイズを設定します
''' </summary>
''' <param name="size">バッファサイズ</param>
Public Property Let BufferSize(ByVal size As Long)
End Property

''' <summary>
''' 現在のバッファサイズを取得します
''' </summary>
Public Property Get BufferSize() As Long
End Property

''' <summary>
''' フラッシュ間隔を設定します
''' </summary>
''' <param name="intervalMs">フラッシュ間隔（ミリ秒）</param>
Public Property Let FlushInterval(ByVal intervalMs As Long)
End Property

''' <summary>
''' 現在のフラッシュ間隔を取得します
''' </summary>
Public Property Get FlushInterval() As Long
End Property

''' <summary>
''' ログ履歴を取得します
''' </summary>
''' <param name="options">履歴オプション（オプション）</param>
''' <returns>ログ履歴のコレクション</returns>
Public Function GetLogHistory(Optional ByVal options As HistoryOptions) As Collection
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