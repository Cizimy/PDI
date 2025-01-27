VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

'@Description("リソースの排他制御とデッドロック検出を提供するインターフェース")

' 定数定義
Private Const DEFAULT_TIMEOUT_MS As Long = 30000    ' デフォルトタイムアウト（30秒）
Private Const MAX_RETRY_COUNT As Long = 3           ' 最大リトライ回数
Private Const MIN_SPIN_COUNT As Long = 1000         ' 最小スピンカウント
Private Const MAX_RECURSION_DEPTH As Long = 100     ' 最大再帰深度
Private Const DEADLOCK_CHECK_INTERVAL_MS As Long = 1000 ' デッドロック検出間隔（1秒）
Private Const MAX_WAITING_THREADS As Long = 100     ' 最大待機スレッド数
Private Const CLEANUP_INTERVAL_MS As Long = 60000   ' クリーンアップ間隔（1分）

' イベント定義
Public Event LockAcquired(ByVal lockId As String, ByVal waitTime As Double)
Public Event LockReleased(ByVal lockId As String, ByVal holdTime As Double)
Public Event DeadlockDetected(ByVal lockId As String, ByVal waitingThreads As Collection)
Public Event TimeoutOccurred(ByVal lockId As String, ByVal timeout As Long)
Public Event PriorityInversion(ByVal lockId As String, ByVal currentPriority As Long, ByVal requiredPriority As Long)
Public Event ResourceContentionDetected(ByVal lockId As String, ByVal contentionLevel As Double)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)

''' <summary>
''' ロックを取得します
''' </summary>
''' <param name="timeout">タイムアウト時間（ミリ秒）。-1は無限待機</param>
''' <param name="options">取得オプション（オプション）</param>
''' <returns>ロックの取得に成功した場合はTrue、タイムアウトした場合はFalse</returns>
''' <remarks>
''' エラー処理要件：
''' - タイムアウトの検証
''' - デッドロックの検出
''' - 優先度逆転の防止
''' - リソース競合の監視
''' - 再入可能性の確認
''' </remarks>
Public Function Acquire(Optional ByVal timeout As Long = -1, _
                       Optional ByVal options As LockOptions) As Boolean
End Function

''' <summary>
''' ロックを解放します
''' </summary>
''' <param name="options">解放オプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - 所有権の確認
''' - 待機スレッドの管理
''' - リソースの解放確認
''' - イベントの発行
''' </remarks>
Public Sub Release(Optional ByVal options As ReleaseOptions)
End Sub

''' <summary>
''' ロックの状態を確認します
''' </summary>
''' <returns>ロックの状態情報</returns>
Public Function GetLockStatus() As LockStatus
End Function

''' <summary>
''' 待機中のスレッドを取得します
''' </summary>
''' <returns>待機スレッドのコレクション</returns>
Public Function GetWaitingThreads() As Collection
End Function

''' <summary>
''' ロックの優先度を設定します
''' </summary>
''' <param name="priority">優先度</param>
Public Property Let Priority(ByVal priority As LockPriority)
End Property

''' <summary>
''' ロックの優先度を取得します
''' </summary>
Public Property Get Priority() As LockPriority
End Property

''' <summary>
''' 再入可能性を設定します
''' </summary>
''' <param name="value">再入可能な場合True</param>
Public Property Let IsReentrant(ByVal value As Boolean)
End Property

''' <summary>
''' 再入可能かどうかを取得します
''' </summary>
Public Property Get IsReentrant() As Boolean
End Property

''' <summary>
''' デッドロック検出を有効/無効にします
''' </summary>
''' <param name="value">有効な場合True</param>
Public Property Let DeadlockDetectionEnabled(ByVal value As Boolean)
End Property

''' <summary>
''' デッドロック検出が有効かどうかを取得します
''' </summary>
Public Property Get DeadlockDetectionEnabled() As Boolean
End Property

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <returns>パフォーマンス情報のコレクション</returns>
Public Function GetPerformanceMetrics() As Collection
End Function

''' <summary>
''' デッドロックを検出します
''' </summary>
''' <param name="options">検出オプション（オプション）</param>
''' <returns>デッドロック情報のコレクション</returns>
Public Function DetectDeadlocks(Optional ByVal options As DeadlockOptions) As Collection
End Function

''' <summary>
''' ロックの状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function

''' <summary>
''' リソースを解放します
''' </summary>
Public Sub Cleanup()
End Sub