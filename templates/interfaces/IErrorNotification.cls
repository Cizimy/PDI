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

''' <summary>
''' エラー通知レベルを設定します
''' </summary>
''' <remarks>
''' エラー処理要件：
''' - 無効な通知レベルの検証
''' - 現在の通知処理への影響考慮
''' - 設定変更のログ記録
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
''' <remarks>
''' エラー処理要件：
''' - 通知レベルに応じた処理
''' - 通知失敗時の代替処理
''' - リソースの適切な解放
''' - 通知の重複防止
''' 
''' 実装時の注意点：
''' 1. 通知処理自体のエラーを適切に処理すること
''' 2. パフォーマンスに影響を与えないよう注意すること
''' 3. UI スレッドをブロックしないこと
''' 4. 機密情報の適切な処理を行うこと
''' </remarks>
Public Sub NotifyError(ByRef errorInfo As ErrorInfo)
End Sub

''' <summary>
''' 通知先を追加します
''' </summary>
''' <param name="target">通知先オブジェクト</param>
''' <remarks>
''' エラー処理要件：
''' - 通知先の有効性確認
''' - 重複登録の防止
''' - リソース制限の考慮
''' </remarks>
Public Sub AddNotificationTarget(ByVal target As INotificationTarget)
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
''' </remarks>
Public Sub SetNotificationFilter(ByVal filterExpression As String)
End Sub

''' <summary>
''' 通知履歴を取得します
''' </summary>
''' <returns>通知履歴のコレクション</returns>
''' <remarks>
''' エラー処理要件：
''' - メモリ使用量の考慮
''' - 大量データの効率的な処理
''' - 履歴の整合性確認
''' </remarks>
Public Function GetNotificationHistory() As Collection
End Function