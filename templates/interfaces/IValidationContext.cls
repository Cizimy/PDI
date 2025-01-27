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

'@Description("バリデーション実行時のコンテキスト情報を管理し、実行状態とリソースを制御するインターフェース")

' 定数定義
Private Const MAX_CONTEXT_ENTRIES As Long = 1000    ' コンテキストエントリの最大数
Private Const MAX_ERROR_LENGTH As Long = 1024       ' エラーメッセージの最大長
Private Const CACHE_TIMEOUT_MS As Long = 300000     ' キャッシュのタイムアウト時間（5分）
Private Const MAX_SNAPSHOT_COUNT As Long = 10       ' 保持する最大スナップショット数

' イベント定義
Public Event ContextChanged(ByVal key As String, ByVal oldValue As Variant, ByVal newValue As Variant)
Public Event ValidationStarted(ByVal validatorId As String, ByVal contextSnapshot As IValidationContext)
Public Event ValidationCompleted(ByVal validatorId As String, ByVal success As Boolean, ByVal results As Collection)
Public Event ValidationProgress(ByVal validatorId As String, ByVal progressPercentage As Long)
Public Event ContextStateChanged(ByVal oldState As ValidationContextState, ByVal newState As ValidationContextState)
Public Event ResourceExhausted(ByVal resourceType As String, ByVal currentUsage As Long)
Public Event CacheUpdated(ByVal key As String, ByVal isCached As Boolean)
Public Event SnapshotCreated(ByVal snapshotId As String)
Public Event ValidationError(ByVal errorInfo As ErrorInfo)

''' <summary>
''' コンテキストデータを設定します
''' </summary>
''' <param name="key">データのキー</param>
''' <param name="value">設定する値</param>
''' <param name="cacheOptions">キャッシュオプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - キーの一意性確認
''' - 値の型チェック
''' - メモリ使用量の監視
''' - 無効なキーの検出
''' - キャッシュ制限の確認
''' - 永続化要件の確認
''' 
''' 実装時の注意点：
''' 1. キーは大文字小文字を区別すること
''' 2. 値のディープコピーを保存すること
''' 3. 循環参照を検出すること
''' 4. 変更通知イベントを発行すること
''' 5. キャッシュ戦略を適用すること
''' </remarks>
Public Property Let ContextData(ByVal key As String, ByVal value As Variant, Optional ByVal cacheOptions As ValidationCacheOptions)
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
''' - キャッシュの有効性確認
''' - 読み取りタイムアウトの処理
''' </remarks>
Public Property Get ContextData(ByVal key As String) As Variant
End Property

''' <summary>
''' カスタムエラーメッセージを設定します
''' </summary>
''' <param name="rule">バリデーションルール</param>
''' <param name="message">カスタムメッセージ</param>
''' <param name="messageType">メッセージの種類（情報/警告/エラー）</param>
''' <param name="options">メッセージオプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - メッセージの妥当性確認
''' - ルールの存在確認
''' - メッセージの多言語対応
''' - メッセージタイプの検証
''' - メッセージ長の制限
''' - プレースホルダーの検証
''' </remarks>
Public Sub SetCustomErrorMessage(ByVal rule As IValidationRule, ByVal message As String, Optional ByVal messageType As ErrorNotificationLevel = enlError, Optional ByVal options As ValidationMessageOptions)
End Sub

''' <summary>
''' カスタムエラーメッセージを取得します
''' </summary>
''' <param name="rule">バリデーションルール</param>
''' <param name="locale">ロケール（オプション）</param>
''' <returns>カスタムメッセージ（未設定の場合は空文字列）</returns>
Public Function GetCustomErrorMessage(ByVal rule As IValidationRule, Optional ByVal locale As String) As String
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
''' - 実行中の検証への影響考慮
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
''' <param name="options">クリアオプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - リソースの適切な解放
''' - 進行中のバリデーションへの影響考慮
''' - 状態の整合性維持
''' - イベントの適切な発行
''' - キャッシュの適切な処理
''' - スナップショットの管理
''' </remarks>
Public Sub Clear(Optional ByVal options As ValidationClearOptions)
End Sub

''' <summary>
''' コンテキストの状態を検証します
''' </summary>
''' <param name="validationLevel">検証レベル（オプション）</param>
''' <returns>有効な場合True</returns>
''' <remarks>
''' エラー処理要件：
''' - データの整合性確認
''' - リソースの状態確認
''' - 循環参照の検出
''' - メモリリークの検出
''' - キャッシュの有効性確認
''' - パフォーマンス指標の確認
''' </remarks>
Public Function Validate(Optional ByVal validationLevel As ValidationLevel = vlNormal) As Boolean
End Function

''' <summary>
''' コンテキストのスナップショットを作成します
''' </summary>
''' <param name="options">スナップショットオプション（オプション）</param>
''' <returns>現在の状態のコピー</returns>
''' <remarks>
''' エラー処理要件：
''' - メモリ使用量の考慮
''' - ディープコピーの実行
''' - 参照の適切な処理
''' - イベントハンドラの適切な複製
''' - スナップショット数の制限
''' - 古いスナップショットの管理
''' </remarks>
Public Function CreateSnapshot(Optional ByVal options As ValidationSnapshotOptions) As IValidationContext
End Function

''' <summary>
''' バリデーションの依存関係を設定します
''' </summary>
''' <param name="dependentRule">依存するルール</param>
''' <param name="requiredRule">必要とされるルール</param>
''' <param name="dependencyType">依存関係の種類（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - 循環依存の検出
''' - ルールの存在確認
''' - 依存関係の整合性確認
''' - 優先度との整合性確認
''' - 実行順序への影響確認
''' </remarks>
Public Sub SetRuleDependency(ByVal dependentRule As IValidationRule, ByVal requiredRule As IValidationRule, Optional ByVal dependencyType As ValidationDependencyType = vdtRequired)
End Sub

''' <summary>
''' バリデーションの実行順序を取得します
''' </summary>
''' <param name="optimizationStrategy">最適化戦略（オプション）</param>
''' <returns>実行順序に並べられたルールのコレクション</returns>
''' <remarks>
''' エラー処理要件：
''' - 循環依存のチェック
''' - 優先度の考慮
''' - 依存関係の解決
''' - 実行順序の最適化
''' - パフォーマンスの考慮
''' </remarks>
Public Function GetExecutionOrder(Optional ByVal optimizationStrategy As ValidationOptimizationStrategy = vosDefault) As Collection
End Function

''' <summary>
''' バリデーション結果をマージします
''' </summary>
''' <param name="results">マージする結果のコレクション</param>
''' <param name="options">マージオプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - 結果の整合性確認
''' - 重複の処理
''' - コンフリクトの解決
''' - メモリ使用量の最適化
''' - 優先順位の考慮
''' - イベントの発行
''' </remarks>
Public Sub MergeResults(ByVal results As Collection, Optional ByVal options As ValidationMergeOptions)
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

''' <summary>
''' コンテキストの状態を取得します
''' </summary>
''' <returns>現在の状態</returns>
Public Property Get State() As ValidationContextState
End Property

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <returns>パフォーマンス情報のコレクション</returns>
Public Function GetPerformanceMetrics() As Collection
End Function

''' <summary>
''' リソース使用状況を取得します
''' </summary>
''' <returns>リソース使用情報のコレクション</returns>
Public Function GetResourceUsage() As Collection
End Function

''' <summary>
''' 非同期バリデーションをサポートしているかどうかを取得します
''' </summary>
Public Property Get SupportsAsyncValidation() As Boolean
End Property

''' <summary>
''' キャッシュ戦略を設定します
''' </summary>
''' <param name="strategy">キャッシュ戦略</param>
Public Property Let CacheStrategy(ByVal strategy As IValidationCacheStrategy)
End Property