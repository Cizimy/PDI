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
Private mIsInitialized As Boolean

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    ' 必要なモジュールの初期化を確認
    If Not modFileIO.IsInitialized Then
        modFileIO.InitializeModule
    End If
    
    Set mLock = New clsLock
    mIsInitialized = True
End Sub

Private Sub Class_Terminate()
    Set mLock = Nothing
    mIsInitialized = False
End Sub

Private Sub InitializeIfNeeded()
    If Not mIsInitialized Then
        Class_Initialize
    End If
End Sub

' ======================
' IFileOperations実装
' ======================
Private Function IFileOperations_ReadTextFile(ByVal filePath As String, _
                                            Optional ByVal encoding As String) As String
    InitializeIfNeeded
    IFileOperations_ReadTextFile = modFileIO.ReadTextFile(filePath, encoding)
End Function

Private Function IFileOperations_WriteTextFile(ByVal filePath As String, _
                                             ByVal content As String, _
                                             Optional ByVal append As Boolean = False, _
                                             Optional ByVal encoding As String) As Boolean
    InitializeIfNeeded
    IFileOperations_WriteTextFile = modFileIO.WriteTextFile(filePath, content, append, encoding)
End Function

Private Function IFileOperations_ReadBinaryFile(ByVal filePath As String) As Byte()
    InitializeIfNeeded
    IFileOperations_ReadBinaryFile = modFileIO.ReadBinaryFile(filePath)
End Function

Private Function IFileOperations_WriteBinaryFile(ByVal filePath As String, _
                                               ByRef data() As Byte) As Boolean
    InitializeIfNeeded
    IFileOperations_WriteBinaryFile = modFileIO.WriteBinaryFile(filePath, data)
End Function

Private Function IFileOperations_FileExists(ByVal filePath As String) As Boolean
    InitializeIfNeeded
    IFileOperations_FileExists = modFileIO.FileExists(filePath)
End Function

Private Function IFileOperations_FolderExists(ByVal folderPath As String) As Boolean
    InitializeIfNeeded
    IFileOperations_FolderExists = modFileIO.FolderExists(folderPath)
End Function

Private Function IFileOperations_CreateFolder(ByVal folderPath As String) As Boolean
    InitializeIfNeeded
    IFileOperations_CreateFolder = modFileIO.CreateFolder(folderPath)
End Function

Private Function IFileOperations_DeleteFile(ByVal filePath As String) As Boolean
    InitializeIfNeeded
    IFileOperations_DeleteFile = modFileIO.DeleteFile(filePath)
End Function

Private Function IFileOperations_DeleteFolder(ByVal folderPath As String) As Boolean
    InitializeIfNeeded
    IFileOperations_DeleteFolder = modFileIO.DeleteFolder(folderPath)
End Function

Private Function IFileOperations_GetAbsolutePath(ByVal relativePath As String, _
                                               Optional ByVal basePath As String) As String
    InitializeIfNeeded
    IFileOperations_GetAbsolutePath = modFileIO.GetAbsolutePath(relativePath, basePath)
End Function

' ======================
' パブリックメソッド
' ======================
Public Function CreateInstance() As IFileOperations
    Set CreateInstance = New FileSystemOperations
End Function

' ======================
' テストサポート機能（開発環境専用）
' 警告: これらのメソッドは開発時のテスト目的でのみ使用し、
' 本番環境では使用しないでください。
' ======================
#If DEBUG Then
    ''' <summary>
    ''' モジュールの状態を初期化（テスト用）
    ''' </summary>
    Private Sub ResetModule()
        Class_Terminate
        Class_Initialize
    End Sub
    
    ''' <summary>
    ''' 初期化状態を取得（テスト用）
    ''' </summary>
    Private Property Get IsInitialized() As Boolean
        IsInitialized = mIsInitialized
    End Property
#End If