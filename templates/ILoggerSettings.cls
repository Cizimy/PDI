VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILoggerSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' ロガー設定インターフェース
' ======================

''' <remarks>
''' このインターフェースを実装するクラスは、以下のエラー処理を実装する必要があります：
'''
''' 1. 設定値の検証
'''    - NULL値のチェック
'''    - 範囲外の値の検出
'''    - 不正なフォーマットの検出
'''    - 依存関係のある設定値の整合性確認
'''
''' 2. 設定ファイルの処理
'''    - ファイルの存在確認
'''    - 読み取り権限の確認
'''    - 破損ファイルの検出と修復
'''    - デフォルト値の適切な提供
'''
''' 3. 設定変更の管理
'''    - 変更のトランザクション管理
'''    - 変更履歴の記録
'''    - 設定の整合性維持
'''    - 他のコンポーネントへの通知
'''
''' 4. エラー発生時の動作
'''    - デフォルト値の使用
'''    - エラーのログ記録
'''    - 上位層への通知
'''    - リカバリ処理の実行
''' </remarks>

''' <summary>
''' ログレベルを取得します
''' </summary>
''' <returns>ログレベル</returns>
''' <remarks>
''' エラー処理要件：
''' - 未初期化状態の検出
''' - 無効な設定値の検出
''' - デフォルト値の提供
''' </remarks>
Public Property Get LogLevel() As LogLevelEnum
End Property

''' <summary>
''' ログ出力先を取得します
''' </summary>
''' <returns>ログ出力先</returns>
''' <remarks>
''' エラー処理要件：
''' - 出力先の有効性確認
''' - 出力先の状態確認
''' - 必要なリソースの確認
''' </remarks>
Public Property Get LogDestination() As ILogDestination
End Property

''' <summary>
''' ログファイルパスを取得します
''' </summary>
''' <returns>ログファイルパス</returns>
''' <remarks>
''' エラー処理要件：
''' - パスの妥当性検証
''' - アクセス権限の確認
''' - ディスク容量の確認
''' - 代替パスの提供
''' </remarks>
Public Property Get LogFilePath() As String
End Property

''' <summary>
''' ログテーブル名を取得します
''' </summary>
''' <returns>ログテーブル名</returns>
''' <remarks>
''' エラー処理要件：
''' - テーブル名の妥当性確認
''' - テーブルの存在確認
''' - テーブル構造の検証
''' - 権限の確認
''' </remarks>
Public Property Get LogTableName() As String
End Property

''' <summary>
''' ログイベントソースを取得します
''' </summary>
''' <returns>イベントソース名</returns>
''' <remarks>
''' エラー処理要件：
''' - イベントソースの存在確認
''' - 登録権限の確認
''' - ソース名の妥当性検証
''' - 代替ソースの提供
''' </remarks>
Public Property Get LogEventSource() As String
End Property

''' <summary>
''' タイマー間隔を取得します
''' </summary>
''' <returns>タイマー間隔（ミリ秒）</returns>
''' <remarks>
''' エラー処理要件：
''' - 範囲チェック（最小値、最大値）
''' - システムリソースの考慮
''' - パフォーマンスへの影響評価
''' </remarks>
Public Property Get TimerInterval() As Long
End Property

''' <summary>
''' データベース接続文字列を取得します
''' </summary>
''' <returns>接続文字列</returns>
''' <remarks>
''' エラー処理要件：
''' - 接続文字列の妥当性確認
''' - 機密情報の保護
''' - 接続テストの実行
''' - フォールバック設定の提供
''' </remarks>
Public Property Get ConnectionString() As String
End Property