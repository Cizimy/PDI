VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' バリデーターインターフェース
' ======================

''' <remarks>
''' このインターフェースを実装するクラスは、以下のエラー処理を実装する必要があります：
'''
''' 1. 入力値の検証
'''    - NULL値の処理
'''    - データ型の検証
'''    - 範囲チェック
'''    - フォーマット検証
'''    - 文字エンコーディングの考慮
'''
''' 2. ビジネスルールの検証
'''    - 依存関係のチェック
'''    - 整合性の確認
'''    - カスタムルールの適用
'''    - 条件付きバリデーション
'''
''' 3. パフォーマンスとリソース
'''    - 大量データの効率的な検証
'''    - メモリ使用量の最適化
'''    - キャッシュの活用
'''    - タイムアウト処理
'''
''' 4. エラー情報の管理
'''    - 詳細なエラー情報の提供
'''    - 多言語対応
'''    - エラーの階層化
'''    - エラーの集約
'''
''' 5. セキュリティ考慮事項
'''    - 入力値のサニタイズ
'''    - インジェクション攻撃の防止
'''    - 機密情報の保護
'''    - アクセス制御の確認
''' </remarks>

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
''' - 再帰的な検証の制御
''' - 検証ルールの整合性確認
''' 
''' 実装時の注意点：
''' 1. 検証失敗時は必ずErrorMessageを設定すること
''' 2. 複数のエラーが検出された場合は適切に集約すること
''' 3. パフォーマンスに影響を与える検証は最後に実行すること
''' 4. 副作用を発生させないこと
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

''' <summary>
''' 検証ルールを追加します
''' </summary>
''' <param name="rule">検証ルール</param>
''' <remarks>
''' エラー処理要件：
''' - ルールの妥当性確認
''' - 重複ルールの検出
''' - ルール間の依存関係確認
''' - メモリ使用量の監視
''' </remarks>
Public Sub AddRule(ByVal rule As Variant)
End Sub

''' <summary>
''' 検証ルールをクリアします
''' </summary>
''' <remarks>
''' エラー処理要件：
''' - リソースの適切な解放
''' - 実行中の検証への影響考慮
''' - 状態の整合性維持
''' </remarks>
Public Sub ClearRules()
End Sub

''' <summary>
''' 検証の詳細結果を取得します
''' </summary>
''' <returns>検証結果の詳細情報</returns>
''' <remarks>
''' エラー処理要件：
''' - 結果の初期化状態確認
''' - メモリ使用量の考慮
''' - 大量データの効率的な処理
''' </remarks>
Public Function GetValidationDetails() As Variant
End Function