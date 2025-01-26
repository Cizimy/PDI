VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("スタックの基本操作を定義するインターフェース")

'==========================
' インターフェースメソッド
'==========================

''' <summary>
''' スタックに要素を追加
''' </summary>
''' <param name="Item">追加する要素</param>
''' <returns>追加成功時True、スタックが最大数を超えた場合False</returns>
Public Function Push(ByVal Item As String) As Boolean
End Function

''' <summary>
''' スタックから最後の要素を削除して返す
''' </summary>
''' <returns>削除した要素、スタックが空の場合は空文字列</returns>
Public Function Pop() As String
End Function

''' <summary>
''' スタックの最後の要素を削除せずに取得
''' </summary>
''' <returns>最後の要素、スタックが空の場合は空文字列</returns>
Public Function Peek() As String
End Function

''' <summary>
''' スタックが空かどうかを確認
''' </summary>
''' <returns>スタックが空の場合True</returns>
Public Function IsEmpty() As Boolean
End Function

''' <summary>
''' スタック内の要素数を取得
''' </summary>
''' <returns>現在のスタックの要素数</returns>
Public Property Get Count() As Long
End Property

''' <summary>
''' スタックを空にする
''' </summary>
Public Sub Clear()
End Sub