VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValidationRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' バリデーションルールインターフェース
' ======================

''' <summary>
''' 値の妥当性を検証します
''' </summary>
''' <param name="value">検証する値</param>
''' <returns>妥当な場合True</returns>
''' <remarks>
''' エラー処理要件：
''' - 入力値のNULLチェック
''' - データ型の互換性確認
''' - メモリ制限の考慮
''' - 検証ルールの整合性確認
''' 
''' 実装時の注意点：
''' 1. 検証失敗時は必ずErrorMessageを設定すること
''' 2. パフォーマンスに影響を与える検証は最後に実行すること
''' 3. 副作用を発生させないこと
''' </remarks>
Public Function Validate(ByVal value As Variant) As Boolean
End Function

''' <summary>
''' 検証エラーメッセージを取得します
''' </summary>
''' <returns>エラーメッセージ</returns>
''' <remarks>
''' エラー処理要件：
''' - メッセージの初期化状態確認
''' - 文字列長の制限
''' - 特殊文字のエスケープ
''' - 多言語対応
''' 
''' 実装時の注意点：
''' 1. エラーがない場合は空文字列を返すこと
''' 2. メッセージは具体的で理解しやすい内容にすること
''' 3. 機密情報を含めないこと
''' 4. フォーマット指定子を適切に使用すること
''' </remarks>
Public Property Get ErrorMessage() As String
End Property