VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' キューインターフェース
' ======================

''' <summary>
''' キューにアイテムを追加します
''' </summary>
''' <param name="item">追加するアイテム</param>
Public Sub Enqueue(ByVal item As Variant)
End Sub

''' <summary>
''' キューからアイテムを取り出します
''' </summary>
''' <returns>取り出したアイテム</returns>
Public Function Dequeue() As Variant
End Function

''' <summary>
''' キューが空かどうかを確認します
''' </summary>
''' <returns>キューが空の場合True</returns>
Public Function IsEmpty() As Boolean
End Function

''' <summary>
''' キュー内のアイテム数を取得します
''' </summary>
''' <returns>アイテム数</returns>
Public Property Get Count() As Long
End Property

''' <summary>
''' キューの先頭アイテムを参照します（取り出さない）
''' </summary>
''' <returns>先頭アイテム</returns>
Public Function Peek() As Variant
End Function

''' <summary>
''' キューをクリアします
''' </summary>
Public Sub Clear()
End Sub