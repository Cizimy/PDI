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

''' <summary>
''' ログを記録します
''' </summary>
''' <param name="moduleName">モジュール名</param>
''' <param name="message">ログメッセージ</param>
''' <param name="errorCode">エラーコード（オプション）</param>
Public Sub Log(ByVal moduleName As String, _
               ByVal message As String, _
               Optional ByVal errorCode As ErrorCode)
End Sub

''' <summary>
''' ログレベルを設定します
''' </summary>
''' <param name="level">ログレベル</param>
Public Property Let LogLevel(ByVal level As LogLevel)
End Property

''' <summary>
''' 現在のログレベルを取得します
''' </summary>
''' <returns>ログレベル</returns>
Public Property Get LogLevel() As LogLevel
End Property

''' <summary>
''' ログの出力先を設定します
''' </summary>
''' <param name="destination">ログの出力先</param>
Public Property Let LogDestination(ByVal destination As LogDestination)
End Property

''' <summary>
''' 現在のログ出力先を取得します
''' </summary>
''' <returns>ログの出力先</returns>
Public Property Get LogDestination() As LogDestination
End Property