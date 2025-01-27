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

'@Description("バリデーションの実行と結果管理を担当するインターフェース")

' 定数定義
Private Const MAX_RULES As Long = 100               ' 最大ルール数
Private Const MAX_ERROR_LENGTH As Long = 1024       ' エラーメッセージの最大長
Private Const CACHE_DURATION_MS As Long = 300000    ' キャッシュの有効期間（5分）
Private Const DEFAULT_TIMEOUT_MS As Long = 30000    ' デフォルトのタイムアウト時間（30秒）
Private Const MAX_BATCH_SIZE As Long = 1000         ' 最大バッチサイズ
Private Const MIN_MEMORY_THRESHOLD As Long = 102400 ' 最小メモリ閾値（100KB）

' イベント定義
Public Event ValidationStarted(ByVal value As Variant, ByVal context As IValidationContext)
Public Event ValidationCompleted(ByVal value As Variant, ByVal success As Boolean, ByVal metrics As ValidationMetrics)
Public Event RuleExecutionStarted(ByVal rule As IValidationRule, ByVal context As IValidationContext)
Public Event RuleExecutionCompleted(ByVal rule As IValidationRule, ByVal success As Boolean, ByVal executionTime As Double)
Public Event ValidationError(ByVal errorInfo As ErrorInfo, ByVal rule As IValidationRule)
Public Event ResourceExhausted(ByVal resourceType As String, ByVal currentUsage As Long)
Public Event CacheUpdated(ByVal cacheKey As String, ByVal isCached As Boolean)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)

''' <summary>
''' 値の妥当性を検証します
''' </summary>
''' <param name="value">検証する値</param>
''' <param name="context">検証コンテキスト（オプション）</param>
''' <param name="options">検証オプション（オプション）</param>
''' <returns>妥当な場合True</returns>
''' <remarks>
''' エラー処理要件：
''' - 入力値のNULLチェック
''' - データ型の互換性確認
''' - メモリ制限の考慮
''' - 再帰的な検証の制御
''' - 検証ルールの整合性確認
''' - タイムアウト処理
''' - キャッシュの活用
''' 
''' 実装時の注意点：
''' 1. 検証失敗時は必ずErrorMessageを設定すること
''' 2. 複数のエラーが検出された場合は適切に集約すること
''' 3. パフォーマンスに影響を与える検証は最後に実行すること
''' 4. 副作用を発生させないこと
''' 5. メトリクスを収集すること
''' </remarks>
Public Function Validate(ByVal value As Variant, Optional ByVal context As IValidationContext, Optional ByVal options As ValidationOptions) As Boolean
End Function

''' <summary>
''' バッチ検証を実行します
''' </summary>
''' <param name="values">検証する値のコレクション</param>
''' <param name="options">バッチオプション（オプション）</param>
''' <returns>バッチ検証結果</returns>
Public Function ValidateBatch(ByVal values As Collection, Optional ByVal options As BatchValidationOptions) As ValidationBatchResult
End Function

''' <summary>
''' 検証エラーメッセージを取得します
''' </summary>
''' <param name="locale">ロケール（オプション）</param>
''' <returns>エラーメッセージ</returns>
''' <remarks>
''' エラー処理要件：
''' - メッセージの初期化状態確認
''' - 文字列長の制限
''' - 特殊文字のエスケープ
''' - 多言語対応
''' - テンプレートの適用
''' 
''' 実装時の注意点：
''' 1. エラーがない場合は空文字列を返すこと
''' 2. メッセージは具体的で理解しやすい内容にすること
''' 3. 機密情報を含めないこと
''' 4. フォーマット指定子を適切に使用すること
''' 5. メッセージの一貫性を保つこと
''' </remarks>
Public Function GetErrorMessage(Optional ByVal locale As String) As String
End Function

''' <summary>
''' 検証ルールを追加します
''' </summary>
''' <param name="rule">検証ルール</param>
''' <param name="priority">優先度（オプション）</param>
''' <param name="options">追加オプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - ルールの妥当性確認
''' - 重複ルールの検出
''' - ルール間の依存関係確認
''' - メモリ使用量の監視
''' - 最大ルール数の制限
''' - 優先度の検証
''' </remarks>
Public Sub AddRule(ByVal rule As IValidationRule, Optional ByVal priority As ValidationPriority = vpNormal, Optional ByVal options As RuleOptions)
End Sub

''' <summary>
''' 検証ルールを削除します
''' </summary>
''' <param name="rule">削除するルール</param>
''' <param name="options">削除オプション（オプション）</param>
Public Sub RemoveRule(ByVal rule As IValidationRule, Optional ByVal options As RuleRemovalOptions)
End Sub

''' <summary>
''' 検証ルールをクリアします
''' </summary>
''' <param name="options">クリアオプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - リソースの適切な解放
''' - 実行中の検証への影響考慮
''' - 状態の整合性維持
''' - キャッシュの更新
''' </remarks>
Public Sub ClearRules(Optional ByVal options As ClearOptions)
End Sub

''' <summary>
''' 検証の詳細結果を取得します
''' </summary>
''' <param name="options">取得オプション（オプション）</param>
''' <returns>検証結果の詳細情報</returns>
''' <remarks>
''' エラー処理要件：
''' - 結果の初期化状態確認
''' - メモリ使用量の考慮
''' - 大量データの効率的な処理
''' - キャッシュの活用
''' </remarks>
Public Function GetValidationDetails(Optional ByVal options As ValidationDetailsOptions) As ValidationDetails
End Function

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <returns>パフォーマンス情報のコレクション</returns>
Public Function GetPerformanceMetrics() As Collection
End Function

''' <summary>
''' キャッシュ戦略を設定します
''' </summary>
''' <param name="strategy">キャッシュ戦略</param>
Public Property Let CacheStrategy(ByVal strategy As IValidationCacheStrategy)
End Property

''' <summary>
''' タイムアウト時間を設定します
''' </summary>
''' <param name="milliseconds">タイムアウト時間（ミリ秒）</param>
Public Property Let TimeoutMilliseconds(ByVal milliseconds As Long)
End Property

''' <summary>
''' タイムアウト時間を取得します
''' </summary>
Public Property Get TimeoutMilliseconds() As Long
End Property

''' <summary>
''' バリデーターの状態を検証します
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
''' メタデータを設定します
''' </summary>
''' <param name="key">メタデータのキー</param>
''' <param name="value">メタデータの値</param>
Public Sub SetMetadata(ByVal key As String, ByVal value As Variant)
End Sub

''' <summary>
''' メタデータを取得します
''' </summary>
''' <param name="key">メタデータのキー</param>
Public Function GetMetadata(ByVal key As String) As Variant
End Function