VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IConnectionStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

''' <summary>
''' データベース接続文字列を生成するインターフェース
''' </summary>

''' <summary>
''' データベース接続文字列を生成します
''' </summary>
''' <returns>生成された接続文字列</returns>
Public Function BuildConnectionString() As String
End Function

' ======================
' プロパティ
' ======================

''' <summary>
''' サーバー名
''' </summary>
Public Property Get Server() As String
End Property

Public Property Let Server(ByVal value As String)
End Property

''' <summary>
''' データベース名
''' </summary>
Public Property Get Database() As String
End Property

Public Property Let Database(ByVal value As String)
End Property

''' <summary>
''' ユーザー名
''' </summary>
Public Property Get UserName() As String
End Property

Public Property Let UserName(ByVal value As String)
End Property

''' <summary>
''' パスワード
''' </summary>
Public Property Get Password() As String
End Property

Public Property Let Password(ByVal value As String)
End Property

''' <summary>
''' 追加のパラメータ
''' </summary>
Public Property Get AdditionalParameters() As String
End Property

Public Property Let AdditionalParameters(ByVal value As String)
End Property