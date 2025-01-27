VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValidationGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("複数のバリデーターをグループ化し、実行順序と依存関係を管理するインターフェース")

' イベント定義
Public Event GroupValidationStarted(ByVal groupName As String, ByVal context As IValidationContext)
Public Event GroupValidationCompleted(ByVal groupName As String, ByVal success As Boolean, ByVal results As Collection)
Public Event ValidatorAdded(ByVal validator As IValidator)
Public Event ValidatorRemoved(ByVal validator As IValidator)
Public Event GroupDependencyChanged(ByVal dependentGroup As IValidationGroup)
Public Event ExecutionOrderChanged(ByVal newOrder As Long)
Public Event ValidationError(ByVal errorInfo As ErrorInfo)
Public Event GroupStateChanged(ByVal groupName As String, ByVal newState As ValidationGroupState)

' 定数定義
Private Const MAX_VALIDATORS As Long = 100  ' 1グループあたりの最大バリデーター数

''' <summary>
''' グループ名を設定します
''' </summary>
''' <remarks>
''' エラー処理要件：
''' - 名前の一意性確認
''' - 無効な文字の検出
''' - 長さ制限の確認
''' - 既存の処理への影響考慮
''' </remarks>
Public Property Let GroupName(ByVal value As String)
End Property

''' <summary>
''' グループ名を取得します
''' </summary>
Public Property Get GroupName() As String
End Property

''' <summary>
''' グループの優先順位を設定します
''' </summary>
''' <remarks>
''' エラー処理要件：
''' - 優先順位の範囲チェック
''' - グループ間の整合性確認
''' - 実行中の検証への影響考慮
''' </remarks>
Public Property Let Priority(ByVal value As Long)
End Property

''' <summary>
''' グループの優先順位を取得します
''' </summary>
Public Property Get Priority() As Long
End Property

''' <summary>
''' バリデーターを追加します
''' </summary>
''' <param name="validator">追加するバリデーター</param>
''' <remarks>
''' エラー処理要件：
''' - バリデーターの有効性確認
''' - 重複チェック
''' - メモリ使用量の監視
''' - グループ内の整合性確認
''' - 最大バリデーター数の制限
''' - バリデーター間の依存関係確認
''' - イベントの適切な発行
''' </remarks>
Public Sub AddValidator(ByVal validator As IValidator)
End Sub

''' <summary>
''' バリデーターを削除します
''' </summary>
''' <param name="validator">削除するバリデーター</param>
''' <remarks>
''' エラー処理要件：
''' - 存在確認
''' - リソースの解放
''' - 依存関係の更新
''' - 実行中の検証への影響考慮
''' - イベントの適切な発行
''' </remarks>
Public Sub RemoveValidator(ByVal validator As IValidator)
End Sub

''' <summary>
''' グループ全体の検証を実行します
''' </summary>
''' <param name="context">バリデーションコンテキスト</param>
''' <returns>すべてのバリデーターが成功した場合True</returns>
''' <remarks>
''' エラー処理要件：
''' - 部分的な失敗の処理
''' - エラー情報の集約
''' - パフォーマンスの最適化
''' - タイムアウト処理
''' - 非同期実行のサポート
''' - キャンセル処理
''' - イベントの適切な発行
''' </remarks>
Public Function ValidateGroup(ByVal context As IValidationContext) As Boolean
End Function

''' <summary>
''' グループの依存関係を設定します
''' </summary>
''' <param name="dependentGroup">依存するグループ</param>
''' <remarks>
''' エラー処理要件：
''' - 循環依存の検出
''' - グループの存在確認
''' - 依存関係の整合性確認
''' - 実行順序との整合性確認
''' - イベントの適切な発行
''' </remarks>
Public Sub SetGroupDependency(ByVal dependentGroup As IValidationGroup)
End Sub

''' <summary>
''' グループの検証順序を設定します
''' </summary>
''' <param name="order">実行順序（1から開始）</param>
''' <remarks>
''' エラー処理要件：
''' - 順序の一意性確認
''' - 範囲チェック
''' - 依存関係との整合性確認
''' - 実行中の検証への影響考慮
''' - イベントの適切な発行
''' </remarks>
Public Sub SetExecutionOrder(ByVal order As Long)
End Sub

''' <summary>
''' グループの検証結果を取得します
''' </summary>
''' <returns>検証結果のコレクション</returns>
''' <remarks>
''' エラー処理要件：
''' - 結果の整合性確認
''' - メモリ使用量の最適化
''' - エラー情報の集約
''' - 結果のフィルタリング
''' - 結果の優先順位付け
''' </remarks>
Public Function GetGroupValidationResults() As Collection
End Function

''' <summary>
''' グループのメタデータを設定します
''' </summary>
''' <param name="key">メタデータのキー</param>
''' <param name="value">メタデータの値</param>
''' <remarks>
''' エラー処理要件：
''' - キーの一意性確認
''' - 値の型チェック
''' - メモリ使用量の監視
''' - キーの命名規則確認
''' - 値の永続化要件確認
''' </remarks>
Public Sub SetMetadata(ByVal key As String, ByVal value As Variant)
End Sub

''' <summary>
''' グループのメタデータを取得します
''' </summary>
''' <param name="key">メタデータのキー</param>
''' <returns>メタデータの値</returns>
Public Function GetMetadata(ByVal key As String) As Variant
End Function

''' <summary>
''' グループの状態を取得します
''' </summary>
''' <returns>現在の状態</returns>
''' <remarks>
''' エラー処理要件：
''' - 状態の整合性確認
''' - 遷移履歴の管理
''' - イベントの適切な発行
''' </remarks>
Public Property Get State() As ValidationGroupState
End Property

''' <summary>
''' グループの検証をキャンセルします
''' </summary>
''' <remarks>
''' エラー処理要件：
''' - 実行状態の確認
''' - リソースの適切な解放
''' - 依存グループへの影響考慮
''' - イベントの適切な発行
''' </remarks>
Public Sub CancelValidation()
End Sub

''' <summary>
''' グループの検証を一時停止します
''' </summary>
''' <remarks>
''' エラー処理要件：
''' - 実行状態の確認
''' - 中間状態の保存
''' - リソースの維持
''' - イベントの適切な発行
''' </remarks>
Public Sub PauseValidation()
End Sub

''' <summary>
''' グループの検証を再開します
''' </summary>
''' <remarks>
''' エラー処理要件：
''' - 状態の復元確認
''' - リソースの再確保
''' - 依存関係の再検証
''' - イベントの適切な発行
''' </remarks>
Public Sub ResumeValidation()
End Sub

''' <summary>
''' グループの進捗状況を取得します
''' </summary>
''' <returns>進捗率（0-100）</returns>
Public Property Get Progress() As Long
End Property

''' <summary>
''' グループのバリデーターをフィルタリングします
''' </summary>
''' <param name="criteria">フィルタリング条件</param>
''' <returns>フィルタリングされたバリデーターのコレクション</returns>
''' <remarks>
''' エラー処理要件：
''' - 条件式の構文検証
''' - パフォーマンスの最適化
''' - メモリ使用量の監視
''' - 結果の整合性確認
''' </remarks>
Public Function FilterValidators(ByVal criteria As String) As Collection
End Function