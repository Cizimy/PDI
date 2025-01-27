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

'@Description("個別のバリデーションルールを定義し、値の検証とエラー情報の管理を行うインターフェース")

' 定数定義
Private Const MAX_ERROR_LENGTH As Long = 1024        ' エラーメッセージの最大長
Private Const MAX_METADATA_ENTRIES As Long = 100     ' メタデータエントリの最大数
Private Const CACHE_DURATION_MS As Long = 60000      ' キャッシュの有効期間（1分）
Private Const DEFAULT_TIMEOUT_MS As Long = 5000      ' デフォルトのタイムアウト時間（5秒）

' イベント定義
Public Event ValidationStarted(ByVal value As Variant, ByVal context As IValidationContext)
Public Event ValidationCompleted(ByVal value As Variant, ByVal success As Boolean, ByVal executionTime As Double)
Public Event ValidationError(ByVal errorInfo As ErrorInfo)
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
''' - 検証ルールの整合性確認
''' - タイムアウト処理
''' - キャッシュの活用
''' 
''' 実装時の注意点：
''' 1. 検証失敗時は必ずErrorMessageを設定すること
''' 2. パフォーマンスに影響を与える検証は最後に実行すること
''' 3. 副作用を発生させないこと
''' 4. メトリクスを収集すること
''' 5. イベントを適切に発行すること
''' </remarks>
Public Function Validate(ByVal value As Variant, Optional ByVal context As IValidationContext, Optional ByVal options As ValidationOptions) As Boolean
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
''' 検証ルールのメタデータを設定します
''' </summary>
''' <param name="key">メタデータのキー</param>
''' <param name="value">メタデータの値</param>
Public Sub SetMetadata(ByVal key As String, ByVal value As Variant)
End Sub

''' <summary>
''' 検証ルールのメタデータを取得します
''' </summary>
''' <param name="key">メタデータのキー</param>
Public Function GetMetadata(ByVal key As String) As Variant
End Function

''' <summary>
''' 検証ルールの説明を取得します
''' </summary>
''' <param name="locale">ロケール（オプション）</param>
Public Function GetDescription(Optional ByVal locale As String) As String
End Function

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
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
''' 検証ルールの状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function

''' <summary>
''' リソースを解放します
''' </summary>
Public Sub Cleanup()
End Sub