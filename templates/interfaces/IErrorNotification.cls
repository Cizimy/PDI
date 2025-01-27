VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IErrorNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("エラー通知機能を定義するインターフェース")

' イベント定義
Public Event NotificationSent(ByVal errorInfo As ErrorInfo, ByVal level As ErrorNotificationLevel)
Public Event NotificationDelivered(ByVal errorInfo As ErrorInfo, ByVal target As INotificationTarget)
Public Event NotificationFailed(ByVal errorInfo As ErrorInfo, ByVal target As INotificationTarget, ByVal reason As String)
Public Event NotificationAcknowledged(ByVal errorInfo As ErrorInfo, ByVal target As INotificationTarget)

''' <summary>
''' エラー通知レベルを設定します
''' </summary>
''' <remarks>
''' エラー処理要件：
''' - 無効な通知レベルの検証
''' - 現在の通知処理への影響考慮
''' - 設定変更のログ記録
''' - イベントの適切な発行
''' </remarks>
Public Property Let NotificationLevel(ByVal level As ErrorNotificationLevel)
End Property

''' <summary>
''' 現在のエラー通知レベルを取得します
''' </summary>
Public Property Get NotificationLevel() As ErrorNotificationLevel
End Property

''' <summary>
''' エラーを通知します
''' </summary>
''' <param name="errorInfo">エラー情報</param>
''' <param name="options">通知オプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - 通知レベルに応じた処理
''' - 通知失敗時の代替処理
''' - リソースの適切な解放
''' - 通知の重複防止
''' - イベントの適切な発行
''' 
''' 実装時の注意点：
''' 1. 通知処理自体のエラーを適切に処理すること
''' 2. パフォーマンスに影響を与えないよう注意すること
''' 3. UI スレッドをブロックしないこと
''' 4. 機密情報の適切な処理を行うこと
''' 5. 通知の優先順位を考慮すること
''' </remarks>
Public Sub NotifyError(ByRef errorInfo As ErrorInfo, Optional ByVal options As NotificationOptions)
End Sub

''' <summary>
''' 通知先を追加します
''' </summary>
''' <param name="target">通知先オブジェクト</param>
''' <param name="priority">優先度（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - 通知先の有効性確認
''' - 重複登録の防止
''' - リソース制限の考慮
''' - 優先度の検証
''' </remarks>
Public Sub AddNotificationTarget(ByVal target As INotificationTarget, Optional ByVal priority As ValidationPriority = vpNormal)
End Sub

''' <summary>
''' 通知先を削除します
''' </summary>
''' <param name="target">通知先オブジェクト</param>
''' <remarks>
''' エラー処理要件：
''' - 存在しない通知先の処理
''' - 削除中の通知への影響考慮
''' - リソースの適切な解放
''' - イベントの適切な発行
''' </remarks>
Public Sub RemoveNotificationTarget(ByVal target As INotificationTarget)
End Sub

''' <summary>
''' すべての通知先をクリアします
''' </summary>
''' <remarks>
''' エラー処理要件：
''' - 進行中の通知への影響考慮
''' - リソースの適切な解放
''' - 状態の整合性維持
''' - イベントの適切な発行
''' </remarks>
Public Sub ClearNotificationTargets()
End Sub

''' <summary>
''' 通知のフィルタリング条件を設定します
''' </summary>
''' <param name="filterExpression">フィルタリング条件の式</param>
''' <remarks>
''' エラー処理要件：
''' - 式の構文検証
''' - 無効な条件の検出
''' - パフォーマンスへの影響考慮
''' - 既存フィルタの更新
''' </remarks>
Public Sub SetNotificationFilter(ByVal filterExpression As String)
End Sub

''' <summary>
''' 通知履歴を取得します
''' </summary>
''' <param name="options">取得オプション（オプション）</param>
''' <returns>通知履歴のコレクション</returns>
''' <remarks>
''' エラー処理要件：
''' - メモリ使用量の考慮
''' - 大量データの効率的な処理
''' - 履歴の整合性確認
''' - フィルタリングの適用
''' </remarks>
Public Function GetNotificationHistory(Optional ByVal options As NotificationHistoryOptions) As Collection
End Function

''' <summary>
''' 通知テンプレートを設定します
''' </summary>
''' <param name="templateName">テンプレート名</param>
''' <param name="template">テンプレート内容</param>
Public Sub SetNotificationTemplate(ByVal templateName As String, ByVal template As String)
End Sub

''' <summary>
''' 通知テンプレートを取得します
''' </summary>
''' <param name="templateName">テンプレート名</param>
Public Function GetNotificationTemplate(ByVal templateName As String) As String
End Function

''' <summary>
''' 通知のグループ化ルールを設定します
''' </summary>
''' <param name="groupingRule">グループ化ルール</param>
Public Sub SetGroupingRule(ByVal groupingRule As INotificationGroupingRule)
End Sub

''' <summary>
''' 通知の重複排除ルールを設定します
''' </summary>
''' <param name="deduplicationRule">重複排除ルール</param>
Public Sub SetDeduplicationRule(ByVal deduplicationRule As INotificationDeduplicationRule)
End Sub

''' <summary>
''' 通知の配信ポリシーを設定します
''' </summary>
''' <param name="policy">配信ポリシー</param>
Public Sub SetDeliveryPolicy(ByVal policy As INotificationDeliveryPolicy)
End Sub

''' <summary>
''' 通知のバッチ処理を有効/無効にします
''' </summary>
Public Property Let BatchProcessingEnabled(ByVal value As Boolean)
End Property

''' <summary>
''' 通知のバッチ処理が有効かどうかを取得します
''' </summary>
Public Property Get BatchProcessingEnabled() As Boolean
End Property

''' <summary>
''' 通知の進捗状況を取得します
''' </summary>
Public Property Get Progress() As Long
End Property

''' <summary>
''' キャンセル状態を設定/取得します
''' </summary>
Public Property Get IsCancellationRequested() As Boolean
End Property

Public Property Let IsCancellationRequested(ByVal value As Boolean)
End Property