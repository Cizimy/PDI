VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockMessageFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' インターフェース実装
' ======================
Implements IMessageFormatter

' ======================
' プライベート変数
' ======================
Private mWasCalled As Boolean
Private mReturnValue As String
Private mLastMessage As String
Private mLastErrorInfo As ErrorInfo

' ======================
' プロパティ
' ======================
Public Property Get WasCalled() As Boolean
    WasCalled = mWasCalled
End Property

Public Property Get LastMessage() As String
    LastMessage = mLastMessage
End Property

Public Property Get LastErrorInfo() As ErrorInfo
    LastErrorInfo = mLastErrorInfo
End Property

' ======================
' パブリックメソッド
' ======================
Public Sub SetReturnValue(ByVal value As String)
    mReturnValue = value
End Sub

Public Sub Reset()
    mWasCalled = False
    mReturnValue = ""
    mLastMessage = ""
    Set mLastErrorInfo = Nothing
End Sub

' ======================
' IMessageFormatter実装
' ======================
Private Function IMessageFormatter_FormatMessage(ByVal message As String, ByRef errorInfo As ErrorInfo) As String
    mWasCalled = True
    mLastMessage = message
    Set mLastErrorInfo = errorInfo
    IMessageFormatter_FormatMessage = mReturnValue
End Function