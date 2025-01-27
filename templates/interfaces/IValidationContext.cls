VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValidationContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("バリデーション実行時のコンテキスト情報を管理するインターフェース")

''' <summary>
''' コンテキストデータを設定します
''' </summary>
''' <param name="key">データのキー</param>
''' <param name="value">設定する値</param>
''' <remarks>
''' エラー処理要件：
''' - キーの一意性確認
''' - 値の型チェック
''' - メモリ使用量の監視
''' - 無効なキーの検出
''' 
''' 実装時の注意点：
''' 1. キーは大文字小文字を区別すること
''' 2. 値のディープコピーを保存すること
''' 3. 循環参照を検出すること
''' </remarks>
Public Property Let ContextData(ByVal key As String, ByVal value As Variant)
End Property

''' <summary>
''' コンテキストデータを取得します
''' </summary>
''' <param name="key">データのキー</param>
''' <returns>格納されている値</returns>
''' <remarks>
''' エラー処理要件：
''' - 存在しないキーの処理
''' - 型の互換性確認
''' - デフォルト値の提供
''' </remarks>
Public Property Get ContextData(ByVal key As String) As Variant
End Property

''' <summary>
''' カスタムエラーメッセージを設定します
''' </summary>
''' <param name="rule">バリデーションルール</param>
''' <param name="message">カスタムメッセージ</param>
''' <remarks>
''' エラー処理要件：
''' - メッセージの妥当性確認
''' - ルールの存在確認
''' - メッセージの多言語対応
''' </remarks>
Public Sub SetCustomErrorMessage(ByVal rule As IValidationRule, ByVal message As String)
End Sub

''' <summary>
''' カスタムエラーメッセージを取得します
''' </summary>
''' <param name="rule">バリデーションルール</param>
''' <returns>カスタムメッセージ（未設定の場合は空文字列）</returns>
Public Function GetCustomErrorMessage(ByVal rule As IValidationRule) As String
End Function

''' <summary>
''' バリデーション優先度を設定します
''' </summary>
''' <param name="rule">バリデーションルール</param>
''' <param name="priority">優先度</param>
''' <remarks>
''' エラー処理要件：
''' - 優先度の範囲チェック
''' - ルールの存在確認
''' - 優先度の整合性確認
''' </remarks>
Public Sub SetValidationPriority(ByVal rule As IValidationRule, ByVal priority As ValidationPriority)
End Sub

''' <summary>
''' バリデーション優先度を取得します
''' </summary>
''' <param name="rule">バリデーションルール</param>
''' <returns>優先度（未設定の場合はvpNormal）</returns>
Public Function GetValidationPriority(ByVal rule As IValidationRule) As ValidationPriority
End Function

''' <summary>
''' コンテキストをクリアします
''' </summary>
''' <remarks>
''' エラー処理要件：
''' - リソースの適切な解放
''' - 進行中のバリデーションへの影響考慮
''' - 状態の整合性維持
''' </remarks>
Public Sub Clear()
End Sub

''' <summary>
''' コンテキストの状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
''' <remarks>
''' エラー処理要件：
''' - データの整合性確認
''' - リソースの状態確認
''' - 循環参照の検出
''' </remarks>
Public Function Validate() As Boolean
End Function

''' <summary>
''' コンテキストのスナップショットを作成します
''' </summary>
''' <returns>現在の状態のコピー</returns>
''' <remarks>
''' エラー処理要件：
''' - メモリ使用量の考慮
''' - ディープコピーの実行
''' - 参照の適切な処理
''' </remarks>
Public Function CreateSnapshot() As IValidationContext
End Function