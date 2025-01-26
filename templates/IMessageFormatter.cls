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

''' <summary>
''' エラーメッセージのフォーマットを担当するインターフェース
''' </summary>
''' <remarks>
''' このインターフェースは、エラー情報を基にユーザー向けメッセージを
''' フォーマットする責務を持ちます。
''' </remarks>

''' <summary>
''' エラー情報を基にフォーマットされたメッセージを生成します。
''' </summary>
''' <param name="message">基本メッセージ</param>
''' <param name="errorInfo">エラー情報</param>
''' <returns>フォーマットされたエラーメッセージ</returns>
Public Function FormatMessage(ByVal message As String, ByRef errorInfo As ErrorInfo) As String
End Function