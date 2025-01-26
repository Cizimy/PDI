VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IConnectionPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' コネクションプールインターフェース
' ======================

''' <summary>
''' データベース接続を取得します
''' </summary>
''' <returns>データベース接続オブジェクト</returns>
''' <remarks>
''' エラー処理要件：
''' - 接続の有効性確認
''' - プール上限チェック
''' - タイムアウト処理
''' - 再接続ロジック
''' </remarks>
Public Function AcquireConnection() As Object
End Function

''' <summary>
''' データベース接続をプールに返却します
''' </summary>
''' <param name="connection">返却する接続オブジェクト</param>
''' <remarks>
''' エラー処理要件：
''' - 接続の状態確認
''' - トランザクション状態の確認
''' - リソースの適切な解放
''' - プールの状態管理
''' </remarks>
Public Sub ReleaseConnection(ByVal connection As Object)
End Sub

''' <summary>
''' 現在のアクティブな接続数を取得します
''' </summary>
''' <returns>アクティブな接続数</returns>
Public Property Get ActiveConnections() As Long
End Property

''' <summary>
''' プールの最大接続数を取得または設定します
''' </summary>
''' <returns>最大接続数</returns>
Public Property Get MaxConnections() As Long
End Property
Public Property Let MaxConnections(ByVal value As Long)
End Property

''' <summary>
''' 接続タイムアウト時間を取得または設定します（ミリ秒）
''' </summary>
''' <returns>タイムアウト時間</returns>
Public Property Get ConnectionTimeout() As Long
End Property
Public Property Let ConnectionTimeout(ByVal value As Long)
End Property