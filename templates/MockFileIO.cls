VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockFileIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' インターフェース実装
' ======================
Implements IFileIO

' ======================
' プライベート変数
' ======================
Private Type FileExistsEntry
    FilePath As String
    Exists As Boolean
End Type

Private Type CreateFolderEntry
    FolderPath As String
    Success As Boolean
End Type

Private Type OpenFileEntry
    FilePath As String
    Result As Object
End Type

Private Type BinaryFileEntry
    FilePath As String
    Content() As Byte
End Type

Private mFileExistsEntries As Collection
Private mFolderExistsEntries As Collection
Private mCreateFolderEntries As Collection
Private mDeleteFileEntries As Collection
Private mDeleteFolderEntries As Collection
Private mOpenFileEntries As Collection
Private mBinaryFileEntries As Collection

Private mMethodCalls As Collection

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mFileExistsEntries = New Collection
    Set mFolderExistsEntries = New Collection
    Set mCreateFolderEntries = New Collection
    Set mDeleteFileEntries = New Collection
    Set mDeleteFolderEntries = New Collection
    Set mOpenFileEntries = New Collection
    Set mBinaryFileEntries = New Collection
    Set mMethodCalls = New Collection
End Sub

' ======================
' モック設定メソッド
' ======================
Public Sub SetFileExists(ByVal filePath As String, ByVal exists As Boolean)
    Dim entry As FileExistsEntry
    entry.FilePath = filePath
    entry.Exists = exists
    AddToCollection mFileExistsEntries, entry, filePath
End Sub

Public Sub SetFolderExists(ByVal folderPath As String, ByVal exists As Boolean)
    Dim entry As FileExistsEntry
    entry.FilePath = folderPath
    entry.Exists = exists
    AddToCollection mFolderExistsEntries, entry, folderPath
End Sub

Public Sub SetCreateFolderResult(ByVal folderPath As String, ByVal success As Boolean)
    Dim entry As CreateFolderEntry
    entry.FolderPath = folderPath
    entry.Success = success
    AddToCollection mCreateFolderEntries, entry, folderPath
End Sub

Public Sub SetOpenFileResult(ByVal filePath As String, ByVal result As Object)
    Dim entry As OpenFileEntry
    entry.FilePath = filePath
    Set entry.Result = result
    AddToCollection mOpenFileEntries, entry, filePath
End Sub

Public Sub SetBinaryFileContent(ByVal filePath As String, ByRef content() As Byte)
    Dim entry As BinaryFileEntry
    entry.FilePath = filePath
    ReDim entry.Content(LBound(content) To UBound(content))
    Dim i As Long
    For i = LBound(content) To UBound(content)
        entry.Content(i) = content(i)
    Next i
    AddToCollection mBinaryFileEntries, entry, filePath
End Sub

' ======================
' 検証メソッド
' ======================
Public Function WasMethodCalled(ByVal methodName As String) As Boolean
    Dim call As Variant
    For Each call In mMethodCalls
        If call = methodName Then
            WasMethodCalled = True
            Exit Function
        End If
    Next call
    WasMethodCalled = False
End Function

Public Function GetMethodCallCount(ByVal methodName As String) As Long
    Dim count As Long
    Dim call As Variant
    For Each call In mMethodCalls
        If call = methodName Then count = count + 1
    Next call
    GetMethodCallCount = count
End Function

' ======================
' IFileIO実装
' ======================
Private Function IFileIO_FileExists(ByVal filePath As String) As Boolean
    mMethodCalls.Add "FileExists"
    
    Dim entry As FileExistsEntry
    Dim var As Variant
    For Each var In mFileExistsEntries
        entry = var
        If entry.FilePath = filePath Then
            IFileIO_FileExists = entry.Exists
            Exit Function
        End If
    Next var
    IFileIO_FileExists = False
End Function

Private Function IFileIO_FolderExists(ByVal folderPath As String) As Boolean
    mMethodCalls.Add "FolderExists"
    
    Dim entry As FileExistsEntry
    Dim var As Variant
    For Each var In mFolderExistsEntries
        entry = var
        If entry.FilePath = folderPath Then
            IFileIO_FolderExists = entry.Exists
            Exit Function
        End If
    Next var
    IFileIO_FolderExists = False
End Function

Private Function IFileIO_CreateFolder(ByVal folderPath As String) As Boolean
    mMethodCalls.Add "CreateFolder"
    
    Dim entry As CreateFolderEntry
    Dim var As Variant
    For Each var In mCreateFolderEntries
        entry = var
        If entry.FolderPath = folderPath Then
            IFileIO_CreateFolder = entry.Success
            Exit Function
        End If
    Next var
    IFileIO_CreateFolder = False
End Function

Private Function IFileIO_DeleteFile(ByVal filePath As String) As Boolean
    mMethodCalls.Add "DeleteFile"
    IFileIO_DeleteFile = True ' デフォルトで成功を返す
End Function

Private Function IFileIO_CopyFile(ByVal sourceFilePath As String, ByVal destinationFilePath As String) As Boolean
    mMethodCalls.Add "CopyFile"
    IFileIO_CopyFile = True ' デフォルトで成功を返す
End Function

Private Function IFileIO_MoveFile(ByVal sourceFilePath As String, ByVal destinationFilePath As String) As Boolean
    mMethodCalls.Add "MoveFile"
    IFileIO_MoveFile = True ' デフォルトで成功を返す
End Function

Private Function IFileIO_OpenFile(ByVal filePath As String, ByVal mode As String) As Object
    mMethodCalls.Add "OpenFile"
    
    Dim entry As OpenFileEntry
    Dim var As Variant
    For Each var In mOpenFileEntries
        entry = var
        If entry.FilePath = filePath Then
            Set IFileIO_OpenFile = entry.Result
            Exit Function
        End If
    Next var
    Set IFileIO_OpenFile = Nothing
End Function

Private Function IFileIO_CloseFile(ByVal file As Object) As Boolean
    mMethodCalls.Add "CloseFile"
    IFileIO_CloseFile = True ' デフォルトで成功を返す
End Function

' ======================
' プライベートヘルパーメソッド
' ======================
Private Sub AddToCollection(ByVal col As Collection, ByRef entry As Variant, ByVal key As String)
    On Error Resume Next
    col.Remove key
    On Error GoTo 0
    col.Add entry, key
End Sub