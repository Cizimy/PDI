VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("エラー処理の基本機能を定義するインターフェース")

'==========================
' インターフェースメソッド
'==========================

''' <summary>
''' エラーを処理します
''' </summary>
''' <param name="errorInfo">エラー詳細情報</param>
''' <returns>処理を継続する場合はTrue、中断する場合はFalse</returns>
Public Function HandleError(ByRef errorInfo As ErrorInfo) As Boolean
End Function

''' <summary>
''' エラーコードに対応するエラーハンドラーを登録します
''' </summary>
''' <param name="errorCode">エラーコード</param>
''' <param name="handler">エラーハンドラー</param>
Public Sub RegisterHandler(ByVal errorCode As ErrorCode, ByVal handler As IErrorHandler)
End Sub

''' <summary>
''' エラーコードに対応するエラーハンドラーの登録を解除します
''' </summary>
''' <param name="errorCode">エラーコード</param>
Public Sub UnregisterHandler(ByVal errorCode As ErrorCode)
End Sub

''' <summary>
''' エラーコードごとの発生回数をカウントアップします
''' </summary>
''' <param name="errorCode">エラーコード</param>
Public Sub IncrementErrorCount(ByVal errorCode As ErrorCode)
End Sub

''' <summary>
''' エラーコードの発生回数が閾値を超えているかチェックします
''' </summary>
''' <param name="errorCode">エラーコード</param>
''' <param name="threshold">閾値</param>
''' <returns>閾値を超えている場合はTrue</returns>
Public Function CheckErrorThreshold(ByVal errorCode As ErrorCode, ByVal threshold As Long) As Boolean
End Function