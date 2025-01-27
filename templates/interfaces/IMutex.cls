VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMutex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

'@Description("スレッド間の同期と排他制御を提供するインターフェース")

' 定数定義
Private Const DEFAULT_TIMEOUT_MS As Long = 30000    ' デフォルトタイムアウト（30秒）
Private Const MAX_RETRY_COUNT As Long = 3           ' 最大リトライ回数
Private Const MIN_SPIN_COUNT As Long = 1000         ' 最小スピンカウント
Private Const MAX_RECURSION_DEPTH As Long = 100     ' 最大再帰深度
Private Const DEADLOCK_CHECK_INTERVAL_MS As Long = 1000 ' デッドロック検出間隔（1秒）
Private Const MAX_WAITING_THREADS As Long = 100     ' 最大待機スレッド数
Private Const CLEANUP_INTERVAL_MS As Long = 60000   ' クリーンアップ間隔（1分）

' イベント定義
Public Event MutexCreated(ByVal mutexName As String, ByVal initialOwner As Boolean)
Public Event MutexAcquired(ByVal mutexName As String, ByVal waitTime As Double)
Public Event MutexReleased(ByVal mutexName As String, ByVal holdTime As Double)
Public Event DeadlockDetected(ByVal mutexName As String, ByVal waitingThreads As Collection)
Public Event TimeoutOccurred(ByVal mutexName As String, ByVal timeout As Long)
Public Event PriorityInversion(ByVal mutexName As String, ByVal currentPriority As Long, ByVal requiredPriority As Long)
Public Event ResourceContentionDetected(ByVal mutexName As String, ByVal contentionLevel As Double)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)

''' <summary>
''' ミューテックスを作成します
''' </summary>
''' <param name="initialOwner">初期所有者フラグ</param>
''' <param name="name">ミューテックスの名前</param>
''' <param name="options">作成オプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
''' <remarks>
''' エラー処理要件：
''' - 名前の検証
''' - 重複チェック
''' - セキュリティ確認
''' - リソース確保
''' </remarks>
Public Function CreateMutex(ByVal initialOwner As Boolean, _
                          ByVal name As String, _
                          Optional ByVal options As MutexOptions) As Boolean
End Function

''' <summary>
''' ミューテックスを解放します
''' </summary>
''' <param name="options">解放オプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
''' <remarks>
''' エラー処理要件：
''' - 所有権の確認
''' - 待機スレッドの管理
''' - リソースの解放確認
''' - イベントの発行
''' </remarks>
Public Function ReleaseMutex(Optional ByVal options As ReleaseOptions) As Boolean
End Function

''' <summary>
''' ミューテックスの所有権を取得します
''' </summary>
''' <param name="timeoutMilliseconds">タイムアウト時間（ミリ秒）</param>
''' <param name="options">取得オプション（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
Public Function WaitForSingleObject(ByVal timeoutMilliseconds As Long, _
                                  Optional ByVal options As WaitOptions) As Boolean
End Function

''' <summary>
''' ミューテックスの状態を確認します
''' </summary>
''' <returns>ミューテックスの状態情報</returns>
Public Function GetMutexStatus() As MutexStatus
End Function

''' <summary>
''' 待機中のスレッドを取得します
''' </summary>
''' <returns>待機スレッドのコレクション</returns>
Public Function GetWaitingThreads() As Collection
End Function

''' <summary>
''' ミューテックスの優先度を設定します
''' </summary>
''' <param name="priority">優先度</param>
Public Property Let Priority(ByVal priority As MutexPriority)
End Property

''' <summary>
''' ミューテックスの優先度を取得します
''' </summary>
Public Property Get Priority() As MutexPriority
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
''' ミューテックスの状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function

''' <summary>
''' リソースを解放します
''' </summary>
Public Sub Cleanup()
End Sub