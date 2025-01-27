VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockFileOperations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements IFileOperations

Private mFileExists As Boolean
Private mFileAccessError As Boolean
Private mCopyFileWasCalled As Boolean
Private mReadFileWasCalled As Boolean
Private mWriteFileWasCalled As Boolean
Private mFileContents As Dictionary
Private mLastSourcePath As String
Private mLastDestPath As String

Private Sub Class_Initialize()
    Set mFileContents = New Dictionary
    ResetFlags
End Sub

Private Sub Class_Terminate()
    Set mFileContents = Nothing
End Sub

'@Description("ファイルの存在有無を設定")
Public Sub SetupFileExists(ByVal exists As Boolean)
    mFileExists = exists
End Sub

'@Description("ファイルアクセスエラーを設定")
Public Sub SetupFileAccessError(ByVal hasError As Boolean)
    mFileAccessError = hasError
End Sub

'@Description("ファイルの内容を設定")
Public Sub SetupFileContent(ByVal path As String, ByVal content As String)
    mFileContents(path) = content
End Sub

'@Description("CopyFileメソッドが呼び出されたかどうか")
Public Property Get CopyFileWasCalled() As Boolean
    CopyFileWasCalled = mCopyFileWasCalled
End Property

'@Description("ReadTextFileメソッドが呼び出されたかどうか")
Public Property Get ReadFileWasCalled() As Boolean
    ReadFileWasCalled = mReadFileWasCalled
End Property

'@Description("WriteTextFileメソッドが呼び出されたかどうか")
Public Property Get WriteFileWasCalled() As Boolean
    WriteFileWasCalled = mWriteFileWasCalled
End Property

'@Description("最後にコピーされたソースパス")
Public Property Get LastSourcePath() As String
    LastSourcePath = mLastSourcePath
End Property

'@Description("最後にコピーされた宛先パス")
Public Property Get LastDestPath() As String
    LastDestPath = mLastDestPath
End Property

'@Description("フラグをリセット")
Public Sub ResetFlags()
    mFileExists = True
    mFileAccessError = False
    mCopyFileWasCalled = False
    mReadFileWasCalled = False
    mWriteFileWasCalled = False
    mLastSourcePath = ""
    mLastDestPath = ""
    Set mFileContents = New Dictionary
End Sub

' IFileOperations実装

Private Function IFileOperations_FileExists(ByVal path As String) As Boolean
    If mFileAccessError Then
        Err.Raise 75, "MockFileOperations", "ファイルアクセスエラー"
    End If
    IFileOperations_FileExists = mFileExists
End Function

Private Function IFileOperations_FolderExists(ByVal path As String) As Boolean
    If mFileAccessError Then
        Err.Raise 75, "MockFileOperations", "ファイルアクセスエラー"
    End If
    IFileOperations_FolderExists = mFileExists
End Function

Private Sub IFileOperations_CreateFolder(ByVal path As String)
    If mFileAccessError Then
        Err.Raise 75, "MockFileOperations", "ファイルアクセスエラー"
    End If
End Sub

Private Function IFileOperations_CopyFile(ByVal sourcePath As String, ByVal destPath As String) As Boolean
    If mFileAccessError Then
        Err.Raise 75, "MockFileOperations", "ファイルアクセスエラー"
    End If
    
    mCopyFileWasCalled = True
    mLastSourcePath = sourcePath
    mLastDestPath = destPath
    
    If mFileContents.Exists(sourcePath) Then
        mFileContents(destPath) = mFileContents(sourcePath)
    End If
    
    IFileOperations_CopyFile = True
End Function

Private Function IFileOperations_DeleteFile(ByVal path As String) As Boolean
    If mFileAccessError Then
        Err.Raise 75, "MockFileOperations", "ファイルアクセスエラー"
    End If
    
    If mFileContents.Exists(path) Then
        mFileContents.Remove path
    End If
    
    IFileOperations_DeleteFile = True
End Function

Private Function IFileOperations_MoveFile(ByVal sourcePath As String, ByVal destPath As String) As Boolean
    If mFileAccessError Then
        Err.Raise 75, "MockFileOperations", "ファイルアクセスエラー"
    End If
    
    If mFileContents.Exists(sourcePath) Then
        mFileContents(destPath) = mFileContents(sourcePath)
        mFileContents.Remove sourcePath
    End If
    
    IFileOperations_MoveFile = True
End Function

Private Function IFileOperations_ReadTextFile(ByVal path As String) As String
    If mFileAccessError Then
        Err.Raise 75, "MockFileOperations", "ファイルアクセスエラー"
    End If
    
    mReadFileWasCalled = True
    
    If mFileContents.Exists(path) Then
        IFileOperations_ReadTextFile = mFileContents(path)
    Else
        IFileOperations_ReadTextFile = ""
    End If
End Function

Private Sub IFileOperations_WriteTextFile(ByVal path As String, ByVal content As String)
    If mFileAccessError Then
        Err.Raise 75, "MockFileOperations", "ファイルアクセスエラー"
    End If
    
    mWriteFileWasCalled = True
    mFileContents(path) = content
End Sub

Private Function IFileOperations_ReadBinaryFile(ByVal path As String) As Byte()
    If mFileAccessError Then
        Err.Raise 75, "MockFileOperations", "ファイルアクセスエラー"
    End If
    
    ' バイナリファイルの読み込みは未実装
    Dim result() As Byte
    IFileOperations_ReadBinaryFile = result
End Function

Private Sub IFileOperations_WriteBinaryFile(ByVal path As String, ByVal content() As Byte)
    If mFileAccessError Then
        Err.Raise 75, "MockFileOperations", "ファイルアクセスエラー"
    End If
    
    ' バイナリファイルの書き込みは未実装
End Sub

Private Function IFileOperations_GetAbsolutePath(ByVal path As String) As String
    IFileOperations_GetAbsolutePath = path ' 単純に入力パスを返す
End Function