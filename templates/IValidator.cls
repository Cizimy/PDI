VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' バリデーターインターフェース
' ======================

''' <summary>
''' 値の妥当性を検証します
''' </summary>
''' <param name="value">検証する値</param>
''' <returns>妥当な場合True</returns>
Public Function Validate(ByVal value As Variant) As Boolean
End Function

''' <summary>
''' 検証エラーメッセージを取得します
''' </summary>
''' <returns>エラーメッセージ</returns>
Public Property Get ErrorMessage() As String
End Property