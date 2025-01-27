VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAsyncValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("非同期バリデーションを定義するインターフェース")

' イベント定義
Public Event ValidationStarted(ByVal validationId As String, ByVal context As IValidationContext)
Public Event ValidationProgress(ByVal validationId As String, ByVal progressPercentage As Long, ByVal currentOperation As String)
Public Event ValidationCompleted(ByVal validationId As String, ByVal success As Boolean, ByVal results As Collection)
Public Event ValidationError(ByVal validationId As String, ByVal errorInfo As ErrorInfo)
Public Event ValidationCancelled(ByVal validationId As String)
Public Event ResourceExhausted(ByVal validationId As String, ByVal resourceType As String)
Public Event ValidationQueued(ByVal validationId As String, ByVal queuePosition As Long)

''' <summary>
''' 非同期バリデーションを開始します
''' </summary>
''' <param name="value">検証する値</param>
''' <param name="context">バリデーションコンテキスト</param>
''' <param name="priority">実行優先度（オプション）</param>
''' <returns>バリデーションID（キャンセル時に使用）</returns>
''' <remarks>
''' エラー処理要件：
''' - パラメータの検証
''' - リソースの可用性確認
''' - 同時実行数の制限
''' - タイムアウト設定
''' - キューの管理
''' </remarks>
Public Function StartAsyncValidation(ByVal value As Variant, ByVal context As IValidationContext, Optional ByVal priority As ValidationPriority = vpNormal) As String
End Function

''' <summary>
''' 非同期バリデーションをキャンセルします
''' </summary>
''' <param name="validationId">バリデーションID</param>
''' <param name="gracefulShutdown">グレースフルシャットダウンを試みるかどうか</param>
''' <remarks>
''' エラー処理要件：
''' - IDの有効性確認
''' - リソースの適切な解放
''' - 実行中の処理の安全な停止
''' - 依存タスクの処理
''' </remarks>
Public Sub CancelValidation(ByVal validationId As String, Optional ByVal gracefulShutdown As Boolean = True)
End Sub

''' <summary>
''' バリデーションの状態を取得します
''' </summary>
''' <param name="validationId">バリデーションID</param>
''' <returns>現在の状態情報</returns>
Public Function GetValidationStatus(ByVal validationId As String) As ValidationStatus
End Function

''' <summary>
''' タイムアウト時間を設定します
''' </summary>
''' <param name="milliseconds">タイムアウト時間（ミリ秒）</param>
''' <remarks>
''' エラー処理要件：
''' - 値の範囲チェック
''' - 実行中の処理への影響考慮
''' - タイムアウト処理の整合性確認
''' </remarks>
Public Property Let TimeoutMilliseconds(ByVal milliseconds As Long)
End Property

''' <summary>
''' タイムアウト時間を取得します
''' </summary>
Public Property Get TimeoutMilliseconds() As Long
End Property

''' <summary>
''' 最大同時実行数を設定します
''' </summary>
''' <param name="count">最大同時実行数</param>
''' <remarks>
''' エラー処理要件：
''' - 値の範囲チェック
''' - システムリソースの考慮
''' - 実行中タスクへの影響評価
''' </remarks>
Public Property Let MaxConcurrentValidations(ByVal count As Long)
End Property

''' <summary>
''' 最大同時実行数を取得します
''' </summary>
Public Property Get MaxConcurrentValidations() As Long
End Property

''' <summary>
''' 進捗通知の間隔を設定します
''' </summary>
''' <param name="milliseconds">通知間隔（ミリ秒）</param>
''' <remarks>
''' エラー処理要件：
''' - 値の範囲チェック
''' - パフォーマンスへの影響考慮
''' - イベント発行の最適化
''' </remarks>
Public Property Let ProgressNotificationInterval(ByVal milliseconds As Long)
End Property

''' <summary>
''' 進捗通知の間隔を取得します
''' </summary>
Public Property Get ProgressNotificationInterval() As Long
End Property

''' <summary>
''' すべての非同期バリデーションを待機します
''' </summary>
''' <param name="timeoutMilliseconds">タイムアウト時間（ミリ秒）</param>
''' <returns>すべての処理が完了した場合True</returns>
''' <remarks>
''' エラー処理要件：
''' - タイムアウト処理
''' - リソースの監視
''' - デッドロック防止
''' - 部分的な完了の処理
''' </remarks>
Public Function WaitForAll(Optional ByVal timeoutMilliseconds As Long = -1) As Boolean
End Function

''' <summary>
''' 実行中のバリデーションの一覧を取得します
''' </summary>
''' <returns>実行中のバリデーションIDのコレクション</returns>
Public Function GetRunningValidations() As Collection
End Function

''' <summary>
''' バリデーション結果を非同期で取得します
''' </summary>
''' <param name="validationId">バリデーションID</param>
''' <returns>バリデーション結果（完了していない場合はNothing）</returns>
Public Function TryGetResults(ByVal validationId As String) As Variant
End Function

''' <summary>
''' バリデーションキューの状態を取得します
''' </summary>
''' <returns>キューの状態情報</returns>
Public Function GetQueueStatus() As ValidationQueueStatus
End Function

''' <summary>
''' バリデーションの優先度を変更します
''' </summary>
''' <param name="validationId">バリデーションID</param>
''' <param name="newPriority">新しい優先度</param>
''' <remarks>
''' エラー処理要件：
''' - IDの有効性確認
''' - 優先度の範囲チェック
''' - キューの再構築
''' </remarks>
Public Sub ChangePriority(ByVal validationId As String, ByVal newPriority As ValidationPriority)
End Sub

''' <summary>
''' エラー発生時のリトライ戦略を設定します
''' </summary>
''' <param name="strategy">リトライ戦略</param>
Public Property Let RetryStrategy(ByVal strategy As IErrorStrategy)
End Property

''' <summary>
''' エラー発生時のリトライ戦略を取得します
''' </summary>
Public Property Get RetryStrategy() As IErrorStrategy
End Property

''' <summary>
''' バリデーション結果の永続化を設定します
''' </summary>
''' <param name="enabled">永続化を有効にするかどうか</param>
Public Property Let PersistResults(ByVal enabled As Boolean)
End Property

''' <summary>
''' バリデーション結果の永続化状態を取得します
''' </summary>
Public Property Get PersistResults() As Boolean
End Property

''' <summary>
''' リソース使用量の制限を設定します
''' </summary>
''' <param name="resourceType">リソースの種類</param>
''' <param name="limit">制限値</param>
Public Sub SetResourceLimit(ByVal resourceType As String, ByVal limit As Long)
End Sub

''' <summary>
''' リソース使用量の制限を取得します
''' </summary>
''' <param name="resourceType">リソースの種類</param>
Public Function GetResourceLimit(ByVal resourceType As String) As Long
End Function