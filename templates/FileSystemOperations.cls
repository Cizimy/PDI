VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystemOperations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "FileSystemOperations"

' ======================
' インターフェース実装
' ======================
Implements IFileOperations

' ======================
' プライベート変数
' ======================
Private mLock As clsLock

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mLock = New clsLock
    modFileIO.InitializeModule
End Sub

Private Sub Class_Terminate()
    Set mLock = Nothing
End Sub

' ======================
' IFileOperations実装
' ======================
Private Function IFileOperations_ReadTextFile(ByVal filePath As String, _
                                            Optional ByVal encoding As String) As String
    IFileOperations_ReadTextFile = modFileIO.ReadTextFile(filePath, encoding)
End Function

Private Function IFileOperations_WriteTextFile(ByVal filePath As String, _
                                             ByVal content As String, _
                                             Optional ByVal append As Boolean = False, _
                                             Optional ByVal encoding As String) As Boolean
    IFileOperations_WriteTextFile = modFileIO.WriteTextFile(filePath, content, append, encoding)
End Function

Private Function IFileOperations_ReadBinaryFile(ByVal filePath As String) As Byte()
    IFileOperations_ReadBinaryFile = modFileIO.ReadBinaryFile(filePath)
End Function

Private Function IFileOperations_WriteBinaryFile(ByVal filePath As String, _
                                               ByRef data() As Byte) As Boolean
    IFileOperations_WriteBinaryFile = modFileIO.WriteBinaryFile(filePath, data)
End Function

Private Function IFileOperations_FileExists(ByVal filePath As String) As Boolean
    IFileOperations_FileExists = modFileIO.FileExists(filePath)
End Function

Private Function IFileOperations_FolderExists(ByVal folderPath As String) As Boolean
    IFileOperations_FolderExists = modFileIO.FolderExists(folderPath)
End Function

Private Function IFileOperations_CreateFolder(ByVal folderPath As String) As Boolean
    IFileOperations_CreateFolder = modFileIO.CreateFolder(folderPath)
End Function

Private Function IFileOperations_DeleteFile(ByVal filePath As String) As Boolean
    IFileOperations_DeleteFile = modFileIO.DeleteFile(filePath)
End Function

Private Function IFileOperations_DeleteFolder(ByVal folderPath As String) As Boolean
    IFileOperations_DeleteFolder = modFileIO.DeleteFolder(folderPath)
End Function

Private Function IFileOperations_GetAbsolutePath(ByVal relativePath As String, _
                                               Optional ByVal basePath As String) As String
    IFileOperations_GetAbsolutePath = modFileIO.GetAbsolutePath(relativePath, basePath)
End Function

' ======================
' パブリックメソッド
' ======================
Public Function CreateInstance() As IFileOperations
    Set CreateInstance = New FileSystemOperations
End Function