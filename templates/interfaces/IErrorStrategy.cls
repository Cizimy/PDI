VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IErrorStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("エラーからの回復戦略を定義するインターフェース")

' イベント定義
Public Event RecoveryAttemptStarted(ByVal errorInfo As ErrorInfo)
Public Event RecoveryAttemptCompleted(ByVal errorInfo As ErrorInfo, ByVal success As Boolean)
Public Event RecoveryProgressChanged(ByVal errorInfo As ErrorInfo, ByVal progressPercentage As Long)
Public Event FallbackStrategyActivated(ByVal errorInfo As ErrorInfo, ByVal fallbackStrategy As IErrorStrategy)

''' <summary>
''' エラーからの回復を試みます
''' </summary>
''' <param name="errorInfo">エラー情報</param>
''' <param name="context">回復コンテキスト（オプション）</param>
''' <returns>回復成功時True、失敗時False</returns>
''' <remarks>
''' エラー処理要件：
''' - リトライ回数の管理
''' - タイムアウト処理
''' - リソースの適切な解放
''' - 回復不可能な状態の検出
''' - イベントの適切な発行
''' 
''' 実装時の注意点：
''' 1. 再帰的な回復処理を避けること
''' 2. 回復処理自体のエラーを適切に処理すること
''' 3. リソースリークを防ぐこと
''' 4. デッドロックを防ぐこと
''' 5. 回復の進捗を通知すること
''' </remarks>
Public Function RecoverFromError(ByRef errorInfo As ErrorInfo, Optional ByVal context As Variant) As Boolean
End Function

''' <summary>
''' 最大リトライ回数を設定します
''' </summary>
''' <remarks>
''' - 0以上の値を設定すること
''' - デフォルト値は3
''' - 実行中の回復処理への影響を考慮
''' </remarks>
Public Property Let MaxRetryCount(ByVal value As Long)
End Property

''' <summary>
''' 最大リトライ回数を取得します
''' </summary>
Public Property Get MaxRetryCount() As Long
End Property

''' <summary>
''' 回復戦略の優先度を取得します
''' </summary>
''' <remarks>
''' - 値が大きいほど優先度が高い
''' - 複数の回復戦略がある場合、優先度の高い順に実行される
''' - 同じ優先度の場合、登録順で実行
''' </remarks>
Public Property Get Priority() As Long
End Property

''' <summary>
''' 回復戦略の優先度を設定します
''' </summary>
Public Property Let Priority(ByVal value As Long)
End Property

''' <summary>
''' リトライ間隔（ミリ秒）を設定します
''' </summary>
''' <remarks>
''' - 0以上の値を設定すること
''' - デフォルト値は1000（1秒）
''' - 指数バックオフを考慮
''' </remarks>
Public Property Let RetryInterval(ByVal value As Long)
End Property

''' <summary>
''' リトライ間隔（ミリ秒）を取得します
''' </summary>
Public Property Get RetryInterval() As Long
End Property

''' <summary>
''' 回復戦略の説明を取得します
''' </summary>
''' <remarks>
''' - ログ出力やデバッグ用
''' - 具体的な回復方法を説明する文字列を返すこと
''' - 多言語対応を考慮
''' </remarks>
Public Property Get Description() As String
End Property

''' <summary>
''' フォールバック戦略を設定します
''' </summary>
''' <param name="strategy">フォールバック戦略</param>
''' <remarks>
''' エラー処理要件：
''' - 循環参照の検出
''' - 戦略の有効性確認
''' - リソースの確保
''' </remarks>
Public Sub SetFallbackStrategy(ByVal strategy As IErrorStrategy)
End Sub

''' <summary>
''' フォールバック戦略を取得します
''' </summary>
Public Property Get FallbackStrategy() As IErrorStrategy
End Property

''' <summary>
''' 回復処理の進捗状況を取得します
''' </summary>
Public Property Get Progress() As Long
End Property

''' <summary>
''' エラー種別ごとの回復方法を設定します
''' </summary>
''' <param name="errorCode">エラーコード</param>
''' <param name="recoveryAction">回復アクション</param>
Public Sub SetRecoveryAction(ByVal errorCode As ErrorCode, ByVal recoveryAction As IRecoveryAction)
End Sub

''' <summary>
''' エラー種別ごとの回復方法を取得します
''' </summary>
''' <param name="errorCode">エラーコード</param>
Public Function GetRecoveryAction(ByVal errorCode As ErrorCode) As IRecoveryAction
End Function

''' <summary>
''' 回復履歴を取得します
''' </summary>
''' <returns>回復試行の履歴コレクション</returns>
Public Function GetRecoveryHistory() As Collection
End Function

''' <summary>
''' 回復戦略の状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function

''' <summary>
''' リソースを解放します
''' </summary>
Public Sub Cleanup()
End Sub

''' <summary>
''' 回復処理をキャンセルします
''' </summary>
Public Sub CancelRecovery()
End Sub

''' <summary>
''' キャンセル状態を取得します
''' </summary>
Public Property Get IsCancellationRequested() As Boolean
End Property

''' <summary>
''' 非同期回復をサポートしているかどうかを取得します
''' </summary>
Public Property Get SupportsAsyncRecovery() As Boolean
End Property

''' <summary>
''' トランザクションをサポートしているかどうかを取得します
''' </summary>
Public Property Get SupportsTransactions() As Boolean
End Property