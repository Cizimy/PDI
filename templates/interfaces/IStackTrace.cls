VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStackTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("スタックトレース機能を提供するインターフェース")

'==========================
' インターフェースメソッド
'==========================

''' <summary>
''' スタックにエントリを追加
''' </summary>
''' <param name="ClassName">クラス名</param>
''' <param name="MethodName">メソッド名</param>
''' <param name="Parameters">パラメータ（オプション）</param>
''' <returns>追加成功時True、スタックが最大数を超えた場合False</returns>
Public Function PushStackEntry(ByVal ClassName As String, ByVal MethodName As String, Optional ByVal Parameters As Variant) As Boolean
End Function

''' <summary>
''' スタックから最後のエントリを削除して返す
''' </summary>
''' <returns>最後のエントリ、スタックが空の場合は空文字列</returns>
Public Function PopStackEntry() As String
End Function

''' <summary>
''' 現在のスタックトレースを文字列として取得
''' </summary>
''' <returns>スタックトレースの文字列表現</returns>
Public Function GetStackTrace() As String
End Function

''' <summary>
''' 現在のスタックの深さを取得
''' </summary>
''' <returns>スタック内のエントリ数</returns>
Public Function StackDepth() As Long
End Function

''' <summary>
''' スタックを空にする
''' </summary>
Public Sub ClearStack()
End Sub

''' <summary>
''' スタックが空かどうかを確認
''' </summary>
''' <returns>スタックが空の場合True</returns>
Public Function IsEmpty() As Boolean
End Function

''' <summary>
''' スタックの最後のエントリを削除せずに取得
''' </summary>
''' <returns>最後のエントリ、スタックが空の場合は空文字列</returns>
Public Function Peek() As String
End Function