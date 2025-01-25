VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFileOperations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' ファイル操作インターフェース
' ======================

' テキストファイル操作
Public Function ReadTextFile(ByVal filePath As String, _
                           Optional ByVal encoding As String) As String
End Function

Public Function WriteTextFile(ByVal filePath As String, _
                            ByVal content As String, _
                            Optional ByVal append As Boolean = False, _
                            Optional ByVal encoding As String) As Boolean
End Function

' バイナリファイル操作
Public Function ReadBinaryFile(ByVal filePath As String) As Byte()
End Function

Public Function WriteBinaryFile(ByVal filePath As String, _
                              ByRef data() As Byte) As Boolean
End Function

' ファイルシステム操作
Public Function FileExists(ByVal filePath As String) As Boolean
End Function

Public Function FolderExists(ByVal folderPath As String) As Boolean
End Function

Public Function CreateFolder(ByVal folderPath As String) As Boolean
End Function

Public Function DeleteFile(ByVal filePath As String) As Boolean
End Function

Public Function DeleteFolder(ByVal folderPath As String) As Boolean
End Function

' パス操作
Public Function GetAbsolutePath(ByVal relativePath As String, _
                              Optional ByVal basePath As String) As String
End Function