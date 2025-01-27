VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEmergencyLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("緊急時のエラーログ出力を定義するインターフェース")

'==========================
' インターフェースメソッド
'==========================

''' <summary>
''' 緊急時のエラーをログに記録します
''' </summary>
''' <param name="message">エラーメッセージ</param>
''' <param name="errorInfo">エラー情報（オプション）</param>
Public Sub LogEmergencyError(ByVal message As String, Optional ByRef errorInfo As ErrorInfo)
End Sub