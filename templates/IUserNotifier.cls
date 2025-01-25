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

''' <summary>
''' エラー情報をユーザーに通知します
''' </summary>
''' <param name="errorDetail">エラー詳細情報</param>
''' <param name="buttons">表示するボタンの種類（オプション）</param>
''' <param name="title">ダイアログのタイトル（オプション）</param>
''' <returns>ユーザーの選択結果</returns>
Public Function Notify(ByRef errorDetail As typErrorDetail, _
                      Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
                      Optional ByVal title As String = "エラー") As VbMsgBoxResult
End Function

''' <summary>
''' 通知の表示方法を設定します
''' </summary>
''' <param name="style">通知スタイル</param>
Public Property Let NotificationStyle(ByVal style As VbMsgBoxStyle)
End Property

''' <summary>
''' 現在の通知表示方法を取得します
''' </summary>
''' <returns>通知スタイル</returns>
Public Property Get NotificationStyle() As VbMsgBoxStyle
End Property

''' <summary>
''' デフォルトのダイアログタイトルを設定します
''' </summary>
''' <param name="title">タイトル</param>
Public Property Let DefaultTitle(ByVal title As String)
End Property

''' <summary>
''' 現在のデフォルトダイアログタイトルを取得します
''' </summary>
''' <returns>タイトル</returns>
Public Property Get DefaultTitle() As String
End Property