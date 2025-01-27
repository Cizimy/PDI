VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements IIniFile

Private mReturnValues As Dictionary
Private mWrittenValues As Dictionary
Private mGetCallCount As Long
Private mWriteCallCount As Long

Private Sub Class_Initialize()
    Set mReturnValues = New Dictionary
    Set mWrittenValues = New Dictionary
    mGetCallCount = 0
    mWriteCallCount = 0
End Sub

Private Sub Class_Terminate()
    Set mReturnValues = Nothing
    Set mWrittenValues = Nothing
End Sub

'@Description("テスト用の戻り値を設定する")
Public Sub SetReturnValue(ByVal section As String, ByVal key As String, ByVal value As String)
    Dim dictKey As String
    dictKey = section & "|" & key
    mReturnValues(dictKey) = value
End Sub

'@Description("GetPrivateProfileStringの呼び出し回数を取得")
Public Property Get GetCallCount() As Long
    GetCallCount = mGetCallCount
End Property

'@Description("WritePrivateProfileStringの呼び出し回数を取得")
Public Property Get WriteCallCount() As Long
    WriteCallCount = mWriteCallCount
End Property

'@Description("書き込まれた値を取得する")
Public Function GetWrittenValue(ByVal section As String, ByVal key As String) As String
    Dim dictKey As String
    dictKey = section & "|" & key
    If mWrittenValues.Exists(dictKey) Then
        GetWrittenValue = mWrittenValues(dictKey)
    Else
        GetWrittenValue = ""
    End If
End Function

'@Description("IIniFile.GetPrivateProfileStringの実装")
Private Function IIniFile_GetPrivateProfileString(ByVal section As String, ByVal key As String, ByVal defaultValue As String, ByVal filePath As String) As String
    mGetCallCount = mGetCallCount + 1
    
    Dim dictKey As String
    dictKey = section & "|" & key
    
    If mReturnValues.Exists(dictKey) Then
        IIniFile_GetPrivateProfileString = mReturnValues(dictKey)
    Else
        IIniFile_GetPrivateProfileString = defaultValue
    End If
End Function

'@Description("IIniFile.WritePrivateProfileStringの実装")
Private Function IIniFile_WritePrivateProfileString(ByVal section As String, ByVal key As String, ByVal value As String, ByVal filePath As String) As Boolean
    mWriteCallCount = mWriteCallCount + 1
    
    Dim dictKey As String
    dictKey = section & "|" & key
    mWrittenValues(dictKey) = value
    
    IIniFile_WritePrivateProfileString = True
End Function

'@Description("すべてのカウンターをリセットする")
Public Sub ResetCounters()
    mGetCallCount = 0
    mWriteCallCount = 0
    Set mWrittenValues = New Dictionary
End Sub