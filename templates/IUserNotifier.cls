VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IUserNotifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' ユーザー通知インターフェース
' ======================

''' <remarks>
''' このインターフェースを実装するクラスは、以下のエラー処理を実装する必要があります：
'''
''' 1. UI表示に関するエラー処理
'''    - ウィンドウハンドルの検証
'''    - モーダルダイアログのデッドロック防止
'''    - 画面解像度やDPIの考慮
'''    - 多重表示の防止
'''
''' 2. メッセージ内容の処理
'''    - 長文メッセージの適切な表示
'''    - 特殊文字のエスケープ
'''    - 多言語対応
'''    - 機密情報のマスク
'''
''' 3. ユーザー応答の処理
'''    - タイムアウト処理
'''    - 無効な応答の検出
'''    - デフォルト値の提供
'''    - キャンセル操作の適切な処理
'''
''' 4. システム状態への配慮
'''    - バッチ処理中の考慮
'''    - 非対話モードの検出
'''    - システムリソースの確認
'''    - 他のダイアログとの競合回避
''' </remarks>

''' <summary>
''' エラー情報をユーザーに通知します
''' </summary>
''' <param name="errorDetail">エラー詳細情報</param>
''' <param name="buttons">表示するボタンの種類（オプション）</param>
''' <param name="title">ダイアログのタイトル（オプション）</param>
''' <returns>ユーザーの選択結果</returns>
''' <remarks>
''' エラー処理要件：
''' - エラー詳細のNULLチェック
''' - メッセージ長の制限
''' - UI表示失敗時の代替通知手段
''' - ユーザー応答のタイムアウト処理
''' </remarks>
Public Function Notify(ByRef errorDetail As typErrorDetail, _
                      Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
                      Optional ByVal title As String = "エラー") As VbMsgBoxResult
End Function

''' <summary>
''' 通知の表示方法を設定します
''' </summary>
''' <param name="style">通知スタイル</param>
''' <remarks>
''' エラー処理要件：
''' - 無効なスタイル値の検出
''' - スタイル変更時の整合性確認
''' - 現在表示中の通知への影響考慮
''' </remarks>
Public Property Let NotificationStyle(ByVal style As VbMsgBoxStyle)
End Property

''' <summary>
''' 現在の通知表示方法を取得します
''' </summary>
''' <returns>通知スタイル</returns>
''' <remarks>
''' エラー処理要件：
''' - 未初期化状態の検出
''' - デフォルトスタイルの提供
''' - システム状態との整合性確認
''' </remarks>
Public Property Get NotificationStyle() As VbMsgBoxStyle
End Property

''' <summary>
''' デフォルトのダイアログタイトルを設定します
''' </summary>
''' <param name="title">タイトル</param>
''' <remarks>
''' エラー処理要件：
''' - タイトル文字列の検証
''' - 長さ制限の確認
''' - 不正な文字の検出
''' - 設定変更の永続化
''' </remarks>
Public Property Let DefaultTitle(ByVal title As String)
End Property

''' <summary>
''' 現在のデフォルトダイアログタイトルを取得します
''' </summary>
''' <returns>タイトル</returns>
''' <remarks>
''' エラー処理要件：
''' - 未設定時のデフォルト値提供
''' - 文字エンコーディングの考慮
''' - キャッシュの整合性確認
''' </remarks>
Public Property Get DefaultTitle() As String
End Property