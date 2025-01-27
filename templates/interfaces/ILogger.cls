VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' ロガーインターフェース
' ======================

''' <remarks>
''' このインターフェースを実装するクラスは、以下のエラー処理を実装する必要があります：
'''
''' 1. 出力先別のエラー処理
'''    - ファイル出力
'''      * ファイルのロック状態確認
'''      * ディスク容量の確認
'''      * アクセス権限の確認
'''      * 出力失敗時の代替出力先（例：一時ファイル）の使用
'''
'''    - データベース出力
'''      * 接続状態の確認
'''      * トランザクション管理
'''      * デッドロック対策
'''      * 接続失敗時の再試行
'''
'''    - イベントログ出力
'''      * イベントソースの存在確認
'''      * 書き込み権限の確認
'''      * イベントログ容量の確認
'''
'''    - コンソール出力
'''      * 出力バッファのオーバーフロー対策
'''      * 文字コードの互換性確認
'''
''' 2. エラー発生時の動作
'''    - エラーの重大度に応じた適切な処理
'''    - 代替出力先への切り替え
'''    - エラー情報自体のログ記録
'''    - 上位層へのエラー通知
'''
''' 3. パフォーマンスとリソース管理
'''    - ログキューのオーバーフロー対策
'''    - リソースの適切な解放
'''    - 非同期出力時のエラー処理
'''
''' 4. データ検証
'''    - メッセージ長の制限
'''    - 不正な文字の除去
'''    - NULLや空文字の適切な処理
''' </remarks>

''' <summary>
''' ログを記録します
''' </summary>
''' <param name="moduleName">モジュール名</param>
''' <param name="message">ログメッセージ</param>
''' <param name="errorCode">エラーコード（オプション）</param>
''' <remarks>
''' エラー処理要件：
''' - パラメータの検証（NULL、空文字、最大長など）
''' - 出力先の状態確認
''' - エラー発生時の代替処理
''' - リトライメカニズムの実装
''' </remarks>
Public Sub Log(ByVal moduleName As String, _
               ByVal message As String, _
               Optional ByVal errorCode As ErrorCode)
End Sub

''' <summary>
''' ログレベルを設定します
''' </summary>
''' <param name="level">ログレベル</param>
''' <remarks>
''' エラー処理要件：
''' - 無効なログレベルの検証
''' - 設定変更時のログ出力
''' - 現在処理中のログへの影響考慮
''' </remarks>
Public Property Let LogLevel(ByVal level As LogLevel)
End Property

''' <summary>
''' 現在のログレベルを取得します
''' </summary>
''' <returns>ログレベル</returns>
''' <remarks>
''' エラー処理要件：
''' - 初期化状態の確認
''' - デフォルト値の適切な提供
''' </remarks>
Public Property Get LogLevel() As LogLevel
End Property

''' <summary>
''' ログの出力先を設定します
''' </summary>
''' <param name="destination">ログの出力先</param>
''' <remarks>
''' エラー処理要件：
''' - 出力先の有効性確認
''' - 出力先の状態確認
''' - 切り替え時の未処理ログの扱い
''' - 設定変更のトランザクション管理
''' </remarks>
Public Property Let LogDestination(ByVal destination As LogDestination)
End Property

''' <summary>
''' 現在のログ出力先を取得します
''' </summary>
''' <returns>ログの出力先</returns>
''' <remarks>
''' エラー処理要件：
''' - 初期化状態の確認
''' - 出力先の状態確認
''' - デフォルト値の適切な提供
''' </remarks>
Public Property Get LogDestination() As LogDestination
End Property