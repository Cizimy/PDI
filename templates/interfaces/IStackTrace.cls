VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStackTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("スタックトレースの収集、管理、フォーマットを担当するインターフェース")

' 定数定義
Private Const MAX_STACK_DEPTH As Long = 100        ' 最大スタック深度
Private Const MAX_ENTRY_LENGTH As Long = 1024      ' エントリの最大長
Private Const CACHE_DURATION_MS As Long = 60000    ' キャッシュの有効期間（1分）
Private Const MAX_FORMAT_LENGTH As Long = 10240    ' フォーマット後の最大長（10KB）
Private Const DEFAULT_INDENT As Long = 2           ' デフォルトのインデントサイズ

' イベント定義
Public Event StackEntryPushed(ByVal entry As String, ByVal depth As Long)
Public Event StackEntryPopped(ByVal entry As String, ByVal remainingDepth As Long)
Public Event StackCleared()
Public Event MaxDepthReached()
Public Event StackOverflowPrevented(ByVal attemptedEntry As String)
Public Event PerformanceAlert(ByVal metricType As String, ByVal currentValue As Double, ByVal threshold As Double)

''' <summary>
''' スタックにエントリを追加します
''' </summary>
''' <param name="ClassName">クラス名</param>
''' <param name="MethodName">メソッド名</param>
''' <param name="Parameters">パラメータ（オプション）</param>
''' <param name="options">追加オプション（オプション）</param>
''' <returns>追加成功時True、スタックが最大数を超えた場合False</returns>
''' <remarks>
''' エラー処理要件：
''' - パラメータの検証
''' - スタックオーバーフローの防止
''' - メモリ使用量の監視
''' - 循環参照の検出
''' - イベントの適切な発行
''' </remarks>
Public Function PushStackEntry(ByVal ClassName As String, ByVal MethodName As String, Optional ByVal Parameters As Variant, Optional ByVal options As StackEntryOptions) As Boolean
End Function

''' <summary>
''' スタックから最後のエントリを削除して返します
''' </summary>
''' <param name="options">取得オプション（オプション）</param>
''' <returns>最後のエントリ、スタックが空の場合は空文字列</returns>
''' <remarks>
''' エラー処理要件：
''' - 空スタックの処理
''' - メモリ解放の確認
''' - イベントの発行
''' </remarks>
Public Function PopStackEntry(Optional ByVal options As StackEntryOptions) As String
End Function

''' <summary>
''' 現在のスタックトレースを文字列として取得します
''' </summary>
''' <param name="format">フォーマット指定（オプション）</param>
''' <param name="options">取得オプション（オプション）</param>
''' <returns>スタックトレースの文字列表現</returns>
''' <remarks>
''' エラー処理要件：
''' - フォーマットの検証
''' - メモリ使用量の最適化
''' - 長大なスタックの効率的な処理
''' - 特殊文字のエスケープ
''' </remarks>
Public Function GetStackTrace(Optional ByVal format As String, Optional ByVal options As StackTraceOptions) As String
End Function

''' <summary>
''' 現在のスタックの深さを取得します
''' </summary>
''' <returns>スタック内のエントリ数</returns>
Public Function GetStackDepth() As Long
End Function

''' <summary>
''' スタックを空にします
''' </summary>
''' <param name="options">クリアオプション（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - リソースの適切な解放
''' - イベントの発行
''' - 状態の整合性確認
''' </remarks>
Public Sub ClearStack(Optional ByVal options As ClearOptions)
End Sub

''' <summary>
''' スタックが空かどうかを確認します
''' </summary>
''' <returns>スタックが空の場合True</returns>
Public Function IsEmpty() As Boolean
End Function

''' <summary>
''' スタックの最後のエントリを削除せずに取得します
''' </summary>
''' <param name="options">取得オプション（オプション）</param>
''' <returns>最後のエントリ、スタックが空の場合は空文字列</returns>
Public Function Peek(Optional ByVal options As PeekOptions) As String
End Function

''' <summary>
''' スタックの特定位置のエントリを取得します
''' </summary>
''' <param name="index">取得するエントリのインデックス</param>
''' <returns>指定位置のエントリ、範囲外の場合は空文字列</returns>
Public Function GetEntryAt(ByVal index As Long) As String
End Function

''' <summary>
''' スタックの内容を検索します
''' </summary>
''' <param name="searchPattern">検索パターン</param>
''' <param name="options">検索オプション（オプション）</param>
''' <returns>マッチしたエントリのコレクション</returns>
Public Function SearchStack(ByVal searchPattern As String, Optional ByVal options As SearchOptions) As Collection
End Function

''' <summary>
''' スタックの内容をフィルタリングします
''' </summary>
''' <param name="filterExpression">フィルタ式</param>
''' <param name="options">フィルタオプション（オプション）</param>
''' <returns>フィルタリングされたスタックトレース</returns>
Public Function FilterStackTrace(ByVal filterExpression As String, Optional ByVal options As FilterOptions) As String
End Function

''' <summary>
''' スタックトレースをファイルに保存します
''' </summary>
''' <param name="filePath">保存先のファイルパス</param>
''' <param name="options">保存オプション（オプション）</param>
Public Sub SaveToFile(ByVal filePath As String, Optional ByVal options As SaveOptions)
End Sub

''' <summary>
''' パフォーマンス指標を取得します
''' </summary>
''' <returns>パフォーマンス情報のコレクション</returns>
Public Function GetPerformanceMetrics() As Collection
End Function

''' <summary>
''' スタックトレースのフォーマッタを設定します
''' </summary>
''' <param name="formatter">フォーマッタ</param>
Public Property Let StackTraceFormatter(ByVal formatter As IStackTraceFormatter)
End Property

''' <summary>
''' メモリ使用量を最適化します
''' </summary>
Public Sub OptimizeMemoryUsage()
End Sub

''' <summary>
''' スタックトレースの状態を検証します
''' </summary>
''' <returns>有効な場合True</returns>
Public Function Validate() As Boolean
End Function