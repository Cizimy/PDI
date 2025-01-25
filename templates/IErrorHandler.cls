VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' エラーハンドラーインターフェース
' ======================

''' <summary>
''' エラーを処理します
''' </summary>
''' <param name="errorDetail">エラー詳細情報</param>
''' <returns>処理を継続する場合はTrue、中断する場合はFalse</returns>
Public Function HandleError(ByRef errorDetail As typErrorDetail) As Boolean
End Function