VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICompositeValidationRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("複合バリデーションルールを定義するインターフェース")

' イベント定義
Public Event RuleAdded(ByVal rule As IValidationRule)
Public Event RuleRemoved(ByVal rule As IValidationRule)
Public Event ValidationStarted(ByVal context As IValidationContext)
Public Event ValidationCompleted(ByVal context As IValidationContext, ByVal success As Boolean)

''' <summary>
''' バリデーションルールを追加します
''' </summary>
''' <param name="rule">追加するルール</param>
''' <param name="priority">優先順位（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - ルールの有効性確認
''' - 循環参照の検出
''' - メモリ使用量の監視
''' - 重複ルールの検出
''' 
''' 実装時の注意点：
''' 1. ルールの優先順位を考慮すること
''' 2. ルール間の依存関係を管理すること
''' 3. パフォーマンスへの影響を考慮すること
''' 4. イベントを適切に発行すること
''' </remarks>
Public Sub AddRule(ByVal rule As IValidationRule, Optional ByVal priority As ValidationPriority = vpNormal)
End Sub

''' <summary>
''' バリデーションルールを削除します
''' </summary>
''' <param name="rule">削除するルール</param>
''' <remarks>
''' エラー処理要件：
''' - 存在しないルールの処理
''' - 依存関係の更新
''' - リソースの適切な解放
''' - イベントの適切な発行
''' </remarks>
Public Sub RemoveRule(ByVal rule As IValidationRule)
End Sub

''' <summary>
''' ルールの優先順位を設定します
''' </summary>
''' <param name="rule">対象ルール</param>
''' <param name="priority">優先順位</param>
''' <remarks>
''' エラー処理要件：
''' - ルールの存在確認
''' - 優先順位の範囲チェック
''' - 優先順位の一意性確保
''' - 依存関係との整合性確認
''' </remarks>
Public Sub SetRulePriority(ByVal rule As IValidationRule, ByVal priority As ValidationPriority)
End Sub

''' <summary>
''' バリデーション結果の詳細を取得します
''' </summary>
''' <returns>バリデーション結果のコレクション</returns>
''' <remarks>
''' エラー処理要件：
''' - 結果の整合性確認
''' - メモリ使用量の最適化
''' - エラー情報の集約
''' - 結果の優先順位付け
''' </remarks>
Public Function GetValidationDetails() As Collection
End Function

''' <summary>
''' すべてのルールをクリアします
''' </summary>
''' <remarks>
''' エラー処理要件：
''' - リソースの適切な解放
''' - 状態の整合性維持
''' - 進行中の検証への影響考慮
''' - イベントの適切な発行
''' </remarks>
Public Sub ClearRules()
End Sub

''' <summary>
''' ルール間の依存関係を設定します
''' </summary>
''' <param name="dependentRule">依存するルール</param>
''' <param name="dependencyRule">依存されるルール</param>
''' <param name="dependencyType">依存関係の種類</param>
''' <remarks>
''' エラー処理要件：
''' - 循環依存の検出
''' - ルールの存在確認
''' - 依存関係の整合性確認
''' - 依存関係の種類の検証
''' </remarks>
Public Sub SetRuleDependency(ByVal dependentRule As IValidationRule, ByVal dependencyRule As IValidationRule, Optional ByVal dependencyType As DependencyType = dtRequired)
End Sub

''' <summary>
''' 指定したルールの依存関係を取得します
''' </summary>
''' <param name="rule">対象ルール</param>
''' <returns>依存関係のコレクション</returns>
Public Function GetRuleDependencies(ByVal rule As IValidationRule) As Collection
End Function

''' <summary>
''' ルールの実行順序を取得します
''' </summary>
''' <returns>実行順序に並べられたルールのコレクション</returns>
''' <remarks>
''' - 優先順位と依存関係を考慮した実行順序
''' - 循環依存がある場合はエラー
''' - パフォーマンスの最適化
''' - 並列実行の可能性を考慮
''' </remarks>
Public Function GetExecutionOrder() As Collection
End Function

''' <summary>
''' 複合ルールの検証を実行します
''' </summary>
''' <param name="value">検証する値</param>
''' <param name="context">バリデーションコンテキスト</param>
''' <returns>すべてのルールが成功した場合True</returns>
''' <remarks>
''' エラー処理要件：
''' - 部分的な失敗の処理
''' - タイムアウト処理
''' - リソースの監視
''' - イベントの適切な発行
''' </remarks>
Public Function ValidateComposite(ByVal value As Variant, ByVal context As IValidationContext) As Boolean
End Function

''' <summary>
''' 非同期検証をサポートしているかどうかを取得します
''' </summary>
Public Property Get SupportsAsyncValidation() As Boolean
End Property

''' <summary>
''' 並列検証をサポートしているかどうかを取得します
''' </summary>
Public Property Get SupportsParallelValidation() As Boolean
End Property

''' <summary>
''' ルールグループを作成します
''' </summary>
''' <param name="groupName">グループ名</param>
''' <returns>作成されたルールグループ</returns>
''' <remarks>
''' エラー処理要件：
''' - グループ名の一意性確認
''' - リソースの確保
''' - メモリ使用量の監視
''' </remarks>
Public Function CreateRuleGroup(ByVal groupName As String) As IValidationGroup
End Function

''' <summary>
''' ルールをグループに追加します
''' </summary>
''' <param name="rule">追加するルール</param>
''' <param name="groupName">グループ名</param>
''' <remarks>
''' エラー処理要件：
''' - グループの存在確認
''' - ルールの有効性確認
''' - 重複の防止
''' </remarks>
Public Sub AddRuleToGroup(ByVal rule As IValidationRule, ByVal groupName As String)
End Sub

''' <summary>
''' 検証の進行状況を取得します
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