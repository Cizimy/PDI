VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OLEDBConnectionStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements IConnectionStringBuilder

' ======================
' メンバ変数
' ======================
Private mServer As String
Private mDatabase As String
Private mUserName As String
Private mPassword As String
Private mAdditionalParameters As String

' ======================
' IConnectionStringBuilder実装
' ======================
Private Function IConnectionStringBuilder_BuildConnectionString() As String
    Dim connectionString As String
    
    connectionString = "Provider=SQLOLEDB;"
    
    ' サーバー名
    If Len(mServer) > 0 Then
        connectionString = connectionString & "Data Source=" & mServer & ";"
    End If
    
    ' データベース名
    If Len(mDatabase) > 0 Then
        connectionString = connectionString & "Initial Catalog=" & mDatabase & ";"
    End If
    
    ' ユーザー名とパスワード
    If Len(mUserName) > 0 Then
        connectionString = connectionString & "User ID=" & mUserName & ";"
        If Len(mPassword) > 0 Then
            connectionString = connectionString & "Password=" & mPassword & ";"
        End If
    Else
        ' 統合認証を使用
        connectionString = connectionString & "Integrated Security=SSPI;"
    End If
    
    ' 追加のパラメータ
    If Len(mAdditionalParameters) > 0 Then
        connectionString = connectionString & mAdditionalParameters
    End If
    
    IConnectionStringBuilder_BuildConnectionString = connectionString
End Function

Private Property Get IConnectionStringBuilder_Server() As String
    IConnectionStringBuilder_Server = mServer
End Property

Private Property Let IConnectionStringBuilder_Server(ByVal value As String)
    mServer = value
End Property

Private Property Get IConnectionStringBuilder_Database() As String
    IConnectionStringBuilder_Database = mDatabase
End Property

Private Property Let IConnectionStringBuilder_Database(ByVal value As String)
    mDatabase = value
End Property

Private Property Get IConnectionStringBuilder_UserName() As String
    IConnectionStringBuilder_UserName = mUserName
End Property

Private Property Let IConnectionStringBuilder_UserName(ByVal value As String)
    mUserName = value
End Property

Private Property Get IConnectionStringBuilder_Password() As String
    IConnectionStringBuilder_Password = mPassword
End Property

Private Property Let IConnectionStringBuilder_Password(ByVal value As String)
    mPassword = value
End Property

Private Property Get IConnectionStringBuilder_AdditionalParameters() As String
    IConnectionStringBuilder_AdditionalParameters = mAdditionalParameters
End Property

Private Property Let IConnectionStringBuilder_AdditionalParameters(ByVal value As String)
    mAdditionalParameters = value
End Property

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    mServer = ""
    mDatabase = ""
    mUserName = ""
    mPassword = ""
    mAdditionalParameters = ""
End Sub