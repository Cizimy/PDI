VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDatabaseConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

''' <summary>
''' データベース接続設定へのアクセスを提供するインターフェース
''' </summary>

''' <summary>
''' データベース接続文字列を取得します
''' </summary>
''' <returns>データベース接続文字列</returns>
Public Function GetConnectionString() As String
End Function