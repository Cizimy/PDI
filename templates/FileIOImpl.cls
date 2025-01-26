VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileIOImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "FileIOImpl"

' ======================
' インターフェース実装
' ======================
Implements IFileIO

' ======================
' プライベート変数
' ======================
Private mFSO As Object ' Scripting.FileSystemObject
Private mIsInitialized As Boolean

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mFSO = CreateObject("Scripting.FileSystemObject")
    mIsInitialized = True
End Sub

Private Sub Class_Terminate()
    Set mFSO = Nothing
    mIsInitialized = False
End Sub

Private Sub InitializeIfNeeded()
    If Not mIsInitialized Then
        Class_Initialize
    End If
End Sub

' ======================
' IFileIO実装
' ======================
Private Function IFileIO_FileExists(ByVal filePath As String) As Boolean
    InitializeIfNeeded
    On Error Resume Next
    IFileIO_FileExists = mFSO.FileExists(filePath)
    On Error GoTo 0
End Function

Private Function IFileIO_FolderExists(ByVal folderPath As String) As Boolean
    InitializeIfNeeded
    On Error Resume Next
    IFileIO_FolderExists = mFSO.FolderExists(folderPath)
    On Error GoTo 0
End Function

Private Function IFileIO_CreateFolder(ByVal folderPath As String) As Boolean
    InitializeIfNeeded
    On Error GoTo ErrorHandler
    
    If IFileIO_FolderExists(folderPath) Then
        IFileIO_CreateFolder = True
        Exit Function
    End If
    
    mFSO.CreateFolder folderPath
    IFileIO_CreateFolder = True
    Exit Function
    
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrFileAccessDenied
        .Description = "フォルダの作成に失敗しました: " & folderPath
        .Source = MODULE_NAME & ".IFileIO_CreateFolder"
    End With
    Call modError.HandleError(errInfo)
    IFileIO_CreateFolder = False
End Function

Private Function IFileIO_DeleteFile(ByVal filePath As String) As Boolean
    InitializeIfNeeded
    On Error GoTo ErrorHandler
    
    If Not IFileIO_FileExists(filePath) Then
        IFileIO_DeleteFile = True
        Exit Function
    End If
    
    mFSO.DeleteFile filePath, True ' Force = True
    IFileIO_DeleteFile = True
    Exit Function
    
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrFileAccessDenied
        .Description = "ファイルの削除に失敗しました: " & filePath
        .Source = MODULE_NAME & ".IFileIO_DeleteFile"
    End With
    Call modError.HandleError(errInfo)
    IFileIO_DeleteFile = False
End Function

Private Function IFileIO_CopyFile(ByVal sourceFilePath As String, ByVal destinationFilePath As String) As Boolean
    InitializeIfNeeded
    On Error GoTo ErrorHandler
    
    If Not IFileIO_FileExists(sourceFilePath) Then
        Err.Raise ErrFileNotFound, MODULE_NAME & ".IFileIO_CopyFile", "コピー元ファイルが存在しません: " & sourceFilePath
    End If
    
    mFSO.CopyFile sourceFilePath, destinationFilePath, True ' OverWrite = True
    IFileIO_CopyFile = True
    Exit Function
    
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = Err.Number
        .Description = "ファイルのコピーに失敗しました: " & Err.Description
        .Source = MODULE_NAME & ".IFileIO_CopyFile"
    End With
    Call modError.HandleError(errInfo)
    IFileIO_CopyFile = False
End Function

Private Function IFileIO_MoveFile(ByVal sourceFilePath As String, ByVal destinationFilePath As String) As Boolean
    InitializeIfNeeded
    On Error GoTo ErrorHandler
    
    If Not IFileIO_FileExists(sourceFilePath) Then
        Err.Raise ErrFileNotFound, MODULE_NAME & ".IFileIO_MoveFile", "移動元ファイルが存在しません: " & sourceFilePath
    End If
    
    mFSO.MoveFile sourceFilePath, destinationFilePath
    IFileIO_MoveFile = True
    Exit Function
    
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = Err.Number
        .Description = "ファイルの移動に失敗しました: " & Err.Description
        .Source = MODULE_NAME & ".IFileIO_MoveFile"
    End With
    Call modError.HandleError(errInfo)
    IFileIO_MoveFile = False
End Function

Private Function IFileIO_OpenFile(ByVal filePath As String, ByVal mode As String) As Object
    InitializeIfNeeded
    On Error GoTo ErrorHandler
    
    Dim fileMode As Integer
    Select Case UCase$(mode)
        Case "READ"
            fileMode = 1 ' ForReading
        Case "WRITE"
            fileMode = 2 ' ForWriting
        Case "APPEND"
            fileMode = 8 ' ForAppending
        Case Else
            Err.Raise vbObjectError + 1001, MODULE_NAME & ".IFileIO_OpenFile", "無効なファイルモード: " & mode
    End Select
    
    Set IFileIO_OpenFile = mFSO.OpenTextFile(filePath, fileMode, True) ' Create = True
    Exit Function
    
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = Err.Number
        .Description = "ファイルを開けませんでした: " & Err.Description
        .Source = MODULE_NAME & ".IFileIO_OpenFile"
    End With
    Call modError.HandleError(errInfo)
    Set IFileIO_OpenFile = Nothing
End Function

Private Function IFileIO_CloseFile(ByVal file As Object) As Boolean
    On Error GoTo ErrorHandler
    
    If file Is Nothing Then
        IFileIO_CloseFile = True
        Exit Function
    End If
    
    file.Close
    IFileIO_CloseFile = True
    Exit Function
    
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = Err.Number
        .Description = "ファイルを閉じる際にエラーが発生しました: " & Err.Description
        .Source = MODULE_NAME & ".IFileIO_CloseFile"
    End With
    Call modError.HandleError(errInfo)
    IFileIO_CloseFile = False
End Function

' ======================
' パブリックメソッド
' ======================
Public Function CreateInstance() As IFileIO
    Set CreateInstance = New FileIOImpl
End Function

' ======================
' テストサポート機能（開発環境専用）
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