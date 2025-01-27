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

''' <summary>
''' データベース固有の設定値を取得します
''' </summary>
''' <param name="settingName">設定名</param>
''' <returns>設定値</returns>
Public Function GetDatabaseSetting(ByVal settingName As String) As Variant
End Function

' ======================
' プロパティ
' ======================

''' <summary>
''' データベース接続文字列
''' </summary>
Public Property Get ConnectionString() As String
End Property

Public Property Let ConnectionString(ByVal value As String)
End Property

''' <summary>
''' 接続タイムアウト時間
''' </summary>
Public Property Get ConnectionTimeout() As Long
End Property

Public Property Let ConnectionTimeout(ByVal value As Long)
End Property

''' <summary>
''' コマンドタイムアウト時間
''' </summary>
Public Property Get CommandTimeout() As Long
End Property

Public Property Let CommandTimeout(ByVal value As Long)
End Property

''' <summary>
''' データベースの種類
''' </summary>
Public Property Get DatabaseType() As String
End Property

Public Property Let DatabaseType(ByVal value As String)
End Property

''' <summary>
''' ログテーブル名
''' </summary>
Public Property Get LogTableName() As String
End Property

Public Property Let LogTableName(ByVal value As String)
End Property