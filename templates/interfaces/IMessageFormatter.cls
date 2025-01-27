VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMessageFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("メッセージのフォーマットとテンプレート管理を提供するインターフェース")

' 定数定義
Private Const MAX_MESSAGE_LENGTH As Long = 8192     ' メッセージの最大長
Private Const MAX_TEMPLATE_SIZE As Long = 32768     ' テンプレートの最大サイズ
Private Const MAX_PLACEHOLDERS As Long = 100        ' プレースホルダーの最大数
Private Const CACHE_DURATION_MS As Long = 300000    ' キャッシュの有効期間（5分）
Private Const MAX_RETRY_COUNT As Long = 3           ' 最大リトライ回数
Private Const DEFAULT_LOCALE As String = "ja-JP"    ' デフォルトロケール
Private Const CLEANUP_INTERVAL_MS As Long = 3600000 ' クリーンアップ間隔（1時間）

' イベント定義
Public Event MessageFormatted(ByVal originalMessage As String, ByVal formattedMessage As String)
Public Event TemplateLoaded(ByVal templateName As String, ByVal success As Boolean)
Public Event LocaleChanged(ByVal oldLocale As String, ByVal newLocale As String)
Public Event PlaceholderResolved(ByVal placeholder As String, ByVal value As String)
Public Event FormatError(ByVal errorType As String, ByVal details As String)
Public Event CacheUpdated(ByVal templateName As String, ByVal isCached As Boolean)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)

''' <summary>
''' メッセージをフォーマットします
''' </summary>
''' <param name="message">基本メッセージ</param>
''' <param name="errorInfo">エラー情報</param>
''' <param name="options">フォーマットオプション（オプション）</param>
''' <returns>フォーマットされたメッセージ</returns>
''' <remarks>
''' エラー処理要件：
''' - メッセージの検証
''' - プレースホルダーの解決
''' - テンプレートの適用
''' - 文字列長の制限
''' - エスケープ処理
''' </remarks>
Public Function FormatMessage(ByVal message As String, _
                            ByRef errorInfo As ErrorInfo, _
                            Optional ByVal options As FormatOptions) As String
End Function

''' <summary>
''' テンプレートを使用してメッセージをフォーマットします
''' </summary>
''' <param name="templateName">テンプレート名</param>
''' <param name="parameters">パラメータのコレクション</param>
''' <param name="options">フォーマットオプション（オプション）</param>
''' <returns>フォーマットされたメッセージ</returns>
Public Function FormatWithTemplate(ByVal templateName As String, _
                                 ByVal parameters As Collection, _
                                 Optional ByVal options As TemplateOptions) As String
End Function

''' <summary>
''' メッセージテンプレートを設定します
''' </summary>
''' <param name="templateName">テンプレート名</param>
''' <param name="template">テンプレート内容</param>
''' <param name="options">テンプレートオプション（オプション）</param>
Public Sub SetTemplate(ByVal templateName As String, _
                      ByVal template As String, _
                      Optional ByVal options As TemplateOptions)
End Sub

''' <summary>
''' メッセージテンプレートを取得します
''' </summary>
''' <param name="templateName">テンプレート名</param>
''' <returns>テンプレート内容</returns>
Public Function GetTemplate(ByVal templateName As String) As String
End Function

''' <summary>
''' ロケールを設定します
''' </summary>
''' <param name="locale">ロケール</param>
''' <param name="options">ロケールオプション（オプション）</param>
Public Sub SetLocale(ByVal locale As String, Optional ByVal options As LocaleOptions)
End Sub

''' <summary>
''' 現在のロケールを取得します
''' </summary>
Public Property Get CurrentLocale() As String
End Property

''' <summary>
''' カスタムプレースホルダーを追加します
''' </summary>
''' <param name="placeholder">プレースホルダー</param>
''' <param name="resolver">解決関数</param>
Public Sub AddPlaceholder(ByVal placeholder As String, ByVal resolver As IPlaceholderResolver)
End Sub

''' <summary>
''' フォーマット設定をエクスポートします
''' </summary>
''' <param name="filePath">出力ファイルパス</param>
''' <param name="options">エクスポートオプション（オプション）</param>
Public Sub ExportSettings(ByVal filePath As String, Optional ByVal options As ExportOptions)
End Sub

''' <summary>
''' フォーマット設定をインポートします
''' </summary>
''' <param name="filePath">入力ファイルパス</param>
''' <param name="options">インポートオプション（オプション）</param>
Public Sub ImportSettings(ByVal filePath As String, Optional ByVal options As ImportOptions)
End Sub

''' <summary>
''' テンプレートをバリデーションします
''' </summary>
''' <param name="template">テンプレート内容</param>
''' <param name="options">検証オプション（オプション）</param>
''' <returns>検証結果</returns>
Public Function ValidateTemplate(ByVal template As String, Optional ByVal options As ValidationOptions) As ValidationResult
End Function

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <returns>パフォーマンス情報のコレクション</returns>
Public Function GetPerformanceMetrics() As Collection
End Function

''' <summary>
''' フォーマッターの状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function

''' <summary>
''' リソースを解放します
''' </summary>
Public Sub Cleanup()
End Sub