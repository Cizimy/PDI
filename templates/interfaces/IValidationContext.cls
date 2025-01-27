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

' イベント定義
Public Event ContextChanged(ByVal key As String, ByVal oldValue As Variant, ByVal newValue As Variant)
Public Event ValidationStarted(ByVal validatorId As String)
Public Event ValidationCompleted(ByVal validatorId As String, ByVal success As Boolean)

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
''' 4. 変更通知イベントを発行すること
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
''' <param name="messageType">メッセージの種類（情報/警告/エラー）</param>
''' <remarks>
''' エラー処理要件：
''' - メッセージの妥当性確認
''' - ルールの存在確認
''' - メッセージの多言語対応
''' - メッセージタイプの検証
''' </remarks>
Public Sub SetCustomErrorMessage(ByVal rule As IValidationRule, ByVal message As String, Optional ByVal messageType As ErrorNotificationLevel = enlError)
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
''' - 依存関係との整合性確認
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
''' - イベントの適切な発行
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
''' - メモリリークの検出
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
''' - イベントハンドラの適切な複製
''' </remarks>
Public Function CreateSnapshot() As IValidationContext
End Function

''' <summary>
''' バリデーションの依存関係を設定します
''' </summary>
''' <param name="dependentRule">依存するルール</param>
''' <param name="requiredRule">必要とされるルール</param>
''' <remarks>
''' エラー処理要件：
''' - 循環依存の検出
''' - ルールの存在確認
''' - 依存関係の整合性確認
''' - 優先度との整合性確認
''' </remarks>
Public Sub SetRuleDependency(ByVal dependentRule As IValidationRule, ByVal requiredRule As IValidationRule)
End Sub

''' <summary>
''' バリデーションの実行順序を取得します
''' </summary>
''' <returns>実行順序に並べられたルールのコレクション</returns>
''' <remarks>
''' エラー処理要件：
''' - 循環依存のチェック
''' - 優先度の考慮
''' - 依存関係の解決
''' - 実行順序の最適化
''' </remarks>
Public Function GetExecutionOrder() As Collection
End Function

''' <summary>
''' バリデーション結果をマージします
''' </summary>
''' <param name="results">マージする結果のコレクション</param>
''' <remarks>
''' エラー処理要件：
''' - 結果の整合性確認
''' - 重複の処理
''' - コンフリクトの解決
''' - メモリ使用量の最適化
''' </remarks>
Public Sub MergeResults(ByVal results As Collection)
End Sub

''' <summary>
''' バリデーションの進行状況を取得します
''' </summary>
''' <returns>進行状況（0-100）</returns>
Public Property Get Progress() As Long
End Property

''' <summary>
''' バリデーションのキャンセル状態を設定/取得します
''' </summary>
Public Property Get IsCancellationRequested() As Boolean
End Property

Public Property Let IsCancellationRequested(ByVal value As Boolean)
End Property