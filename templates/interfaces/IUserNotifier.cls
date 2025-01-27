VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IUserNotifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

'@Description("ユーザーへの通知機能を提供し、複数の通知チャネルと表示方法をサポートするインターフェース")

' 定数定義
Private Const MAX_MESSAGE_LENGTH As Long = 2048     ' メッセージの最大長
Private Const MAX_TITLE_LENGTH As Long = 256        ' タイトルの最大長
Private Const DEFAULT_TIMEOUT_MS As Long = 30000    ' デフォルトのタイムアウト時間（30秒）
Private Const MAX_RETRY_COUNT As Long = 3           ' 最大リトライ回数
Private Const MIN_DISPLAY_TIME_MS As Long = 3000    ' 最小表示時間（3秒）
Private Const MAX_QUEUED_NOTIFICATIONS As Long = 100 ' 最大キュー通知数

' イベント定義
Public Event NotificationDisplayed(ByVal notificationId As String, ByVal level As ErrorNotificationLevel)
Public Event NotificationClosed(ByVal notificationId As String, ByVal userResponse As VbMsgBoxResult)
Public Event NotificationQueued(ByVal notificationId As String, ByVal queuePosition As Long)
Public Event NotificationTimeout(ByVal notificationId As String)
Public Event StyleChanged(ByVal oldStyle As VbMsgBoxStyle, ByVal newStyle As VbMsgBoxStyle)
Public Event UserInteractionStarted(ByVal notificationId As String)
Public Event UserInteractionCompleted(ByVal notificationId As String, ByVal interactionTime As Long)
Public Event NotificationError(ByVal errorInfo As ErrorInfo)
Public Event ResourceExhausted(ByVal resourceType As String, ByVal currentUsage As Long)

''' <summary>
''' エラー情報をユーザーに通知します
''' </summary>
''' <param name="errorDetail">エラー詳細情報</param>
''' <param name="buttons">表示するボタンの種類（オプション）</param>
''' <param name="title">ダイアログのタイトル（オプション）</param>
''' <param name="options">通知オプション（オプション）</param>
''' <returns>ユーザーの選択結果</returns>
''' <remarks>
''' エラー処理要件：
''' - エラー詳細のNULLチェック
''' - メッセージ長の制限
''' - UI表示失敗時の代替通知手段
''' - ユーザー応答のタイムアウト処理
''' - 多重表示の防止
''' - システムリソースの監視
''' </remarks>
Public Function Notify(ByRef errorDetail As ErrorInfo, _
                      Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
                      Optional ByVal title As String = "エラー", _
                      Optional ByVal options As NotificationOptions) As VbMsgBoxResult
End Function

''' <summary>
''' バッチ通知を実行します
''' </summary>
''' <param name="notifications">通知のコレクション</param>
''' <param name="options">バッチオプション（オプション）</param>
''' <returns>バッチ処理の結果</returns>
Public Function NotifyBatch(ByVal notifications As Collection, Optional ByVal options As BatchNotificationOptions) As NotificationBatchResult
End Function

''' <summary>
''' 通知の表示方法を設定します
''' </summary>
''' <param name="style">通知スタイル</param>
''' <param name="options">スタイルオプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - 無効なスタイル値の検出
''' - スタイル変更時の整合性確認
''' - 現在表示中の通知への影響考慮
''' - イベントの適切な発行
''' </remarks>
Public Property Let NotificationStyle(ByVal style As VbMsgBoxStyle, Optional ByVal options As StyleOptions)
End Property

''' <summary>
''' 現在の通知表示方法を取得します
''' </summary>
''' <returns>通知スタイル</returns>
Public Property Get NotificationStyle() As VbMsgBoxStyle
End Property

''' <summary>
''' デフォルトのダイアログタイトルを設定します
''' </summary>
''' <param name="title">タイトル</param>
''' <param name="locale">ロケール（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - タイトル文字列の検証
''' - 長さ制限の確認
''' - 不正な文字の検出
''' - 設定変更の永続化
''' - 多言語対応の確認
''' </remarks>
Public Property Let DefaultTitle(ByVal title As String, Optional ByVal locale As String)
End Property

''' <summary>
''' 現在のデフォルトダイアログタイトルを取得します
''' </summary>
''' <param name="locale">ロケール（オプション）</param>
''' <returns>タイトル</returns>
Public Function GetDefaultTitle(Optional ByVal locale As String) As String
End Function

''' <summary>
''' 通知テンプレートを設定します
''' </summary>
''' <param name="templateName">テンプレート名</param>
''' <param name="template">テンプレート内容</param>
''' <param name="locale">ロケール（オプション）</param>
Public Sub SetNotificationTemplate(ByVal templateName As String, ByVal template As String, Optional ByVal locale As String)
End Sub

''' <summary>
''' 通知テンプレートを取得します
''' </summary>
''' <param name="templateName">テンプレート名</param>
''' <param name="locale">ロケール（オプション）</param>
Public Function GetNotificationTemplate(ByVal templateName As String, Optional ByVal locale As String) As String
End Function

''' <summary>
''' 通知の優先順位を設定します
''' </summary>
''' <param name="notificationId">通知ID</param>
''' <param name="priority">優先順位</param>
Public Sub SetNotificationPriority(ByVal notificationId As String, ByVal priority As ValidationPriority)
End Sub

''' <summary>
''' 通知履歴を取得します
''' </summary>
''' <param name="options">履歴オプション（オプション）</param>
''' <returns>通知履歴のコレクション</returns>
Public Function GetNotificationHistory(Optional ByVal options As HistoryOptions) As Collection
End Function

''' <summary>
''' 通知キューの状態を取得します
''' </summary>
''' <returns>キューの状態情報</returns>
Public Function GetQueueStatus() As NotificationQueueStatus
End Function

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <returns>パフォーマンス情報のコレクション</returns>
Public Function GetPerformanceMetrics() As Collection
End Function

''' <summary>
''' 通知をキャンセルします
''' </summary>
''' <param name="notificationId">通知ID</param>
Public Sub CancelNotification(ByVal notificationId As String)
End Sub

''' <summary>
''' すべての通知をクリアします
''' </summary>
''' <param name="options">クリアオプション（オプション）</param>
Public Sub ClearAllNotifications(Optional ByVal options As ClearOptions)
End Sub

''' <summary>
''' 通知の状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function

''' <summary>
''' リソースを解放します
''' </summary>
Public Sub Cleanup()
End Sub