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

''' <summary>
''' エラーからの回復を試みます
''' </summary>
''' <param name="errorInfo">エラー情報</param>
''' <returns>回復成功時True、失敗時False</returns>
''' <remarks>
''' エラー処理要件：
''' - リトライ回数の管理
''' - タイムアウト処理
''' - リソースの適切な解放
''' - 回復不可能な状態の検出
''' 
''' 実装時の注意点：
''' 1. 再帰的な回復処理を避けること
''' 2. 回復処理自体のエラーを適切に処理すること
''' 3. リソースリークを防ぐこと
''' 4. デッドロックを防ぐこと
''' </remarks>
Public Function RecoverFromError(ByRef errorInfo As ErrorInfo) As Boolean
End Function

''' <summary>
''' 最大リトライ回数を設定します
''' </summary>
''' <remarks>
''' - 0以上の値を設定すること
''' - デフォルト値は3
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
''' </remarks>
Public Property Get Description() As String
End Property