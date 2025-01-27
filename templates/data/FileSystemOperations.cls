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
Private mLock As ILock
Private mPerformanceMonitor As IPerformanceMonitor
Private mFileIO As IFileIO
Private mIsInitialized As Boolean

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    ' 依存性は外部から注入されるため、ここでは初期化しない
    mIsInitialized = False
End Sub

''' <summary>
''' 依存性を注入して初期化します
''' </summary>
''' <param name="lock">ロック機構</param>
''' <param name="performanceMonitor">パフォーマンスモニター</param>
''' <param name="fileIO">ファイルI/O操作</param>
''' <remarks>
''' このメソッドは必ずインスタンス作成後に呼び出す必要があります
''' </remarks>
Public Sub Initialize(ByVal lock As ILock, _
                     ByVal performanceMonitor As IPerformanceMonitor, _
                     ByVal fileIO As IFileIO)
    Set mLock = lock
    Set mPerformanceMonitor = performanceMonitor
    Set mFileIO = fileIO
    mIsInitialized = True
End Sub

Private Sub Class_Terminate()
    Set mLock = Nothing
    Set mPerformanceMonitor = Nothing
    Set mFileIO = Nothing
    mIsInitialized = False
End Sub

Private Sub InitializeIfNeeded()
    If Not mIsInitialized Then
        Err.Raise vbObjectError + 1000, MODULE_NAME, "オブジェクトが初期化されていません。Initialize メソッドを呼び出してください。"
    End If
End Sub

' ======================
' IFileOperations実装
' ======================
Private Function IFileOperations_ReadTextFile(ByVal filePath As String, _
                                           Optional ByVal encoding As String) As String
    InitializeIfNeeded
    mPerformanceMonitor.StartMeasurement "Read Text File"
    Dim content As String
    On Error GoTo ErrorHandler
    content = ReadTextFileUsingFileIO(filePath, encoding)
    IFileOperations_ReadTextFile = content
    mPerformanceMonitor.EndMeasurement "Read Text File"
    Exit Function
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrFileNotFound
        .Description = "ファイルの読み込みに失敗しました: " & filePath
        .Source = MODULE_NAME & ".IFileOperations_ReadTextFile"
    End With
    Call modError.HandleError(errInfo)
    mPerformanceMonitor.EndMeasurement "Read Text File"
    IFileOperations_ReadTextFile = vbNullString
End Function

Private Function IFileOperations_WriteTextFile(ByVal filePath As String, _
                                            ByVal content As String, _
                                            Optional ByVal append As Boolean = False, _
                                            Optional ByVal encoding As String) As Boolean
    InitializeIfNeeded
    mPerformanceMonitor.StartMeasurement "Write Text File"
    Dim result As Boolean
    On Error GoTo ErrorHandler
    result = WriteTextFileUsingFileIO(filePath, content, append, encoding)
    IFileOperations_WriteTextFile = result
    mPerformanceMonitor.EndMeasurement "Write Text File"
    Exit Function
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrFileAccessDenied
        .Description = "ファイルの書き込みに失敗しました: " & filePath
        .Source = MODULE_NAME & ".IFileOperations_WriteTextFile"
    End With
    Call modError.HandleError(errInfo)
    mPerformanceMonitor.EndMeasurement "Write Text File"
    IFileOperations_WriteTextFile = False
End Function

Private Function IFileOperations_ReadBinaryFile(ByVal filePath As String) As Byte()
    InitializeIfNeeded
    mPerformanceMonitor.StartMeasurement "Read Binary File"
    On Error GoTo ErrorHandler
    IFileOperations_ReadBinaryFile = ReadBinaryFileUsingFileIO(filePath)
    mPerformanceMonitor.EndMeasurement "Read Binary File"
    Exit Function
ErrorHandler:
    mPerformanceMonitor.EndMeasurement "Read Binary File"
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrFileNotFound
        .Description = "バイナリファイルの読み込みに失敗しました: " & filePath
        .Source = MODULE_NAME & ".IFileOperations_ReadBinaryFile"
    End With
    Call modError.HandleError(errInfo)
    IFileOperations_ReadBinaryFile = Array()
End Function

Private Function IFileOperations_WriteBinaryFile(ByVal filePath As String, _
                                              ByRef data() As Byte) As Boolean
    InitializeIfNeeded
    mPerformanceMonitor.StartMeasurement "Write Binary File"
    On Error GoTo ErrorHandler
    IFileOperations_WriteBinaryFile = WriteBinaryFileUsingFileIO(filePath, data)
    mPerformanceMonitor.EndMeasurement "Write Binary File"
    Exit Function
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrFileAccessDenied
        .Description = "バイナリファイルの書き込みに失敗しました: " & filePath
        .Source = MODULE_NAME & ".IFileOperations_WriteBinaryFile"
    End With
    Call modError.HandleError(errInfo)
    mPerformanceMonitor.EndMeasurement "Write Binary File"
    IFileOperations_WriteBinaryFile = False
End Function

Private Function IFileOperations_FileExists(ByVal filePath As String) As Boolean
    InitializeIfNeeded
    IFileOperations_FileExists = mFileIO.FileExists(filePath)
End Function

Private Function IFileOperations_FolderExists(ByVal folderPath As String) As Boolean
    InitializeIfNeeded
    IFileOperations_FolderExists = mFileIO.FolderExists(folderPath)
End Function

Private Function IFileOperations_CreateFolder(ByVal folderPath As String) As Boolean
    InitializeIfNeeded
    mPerformanceMonitor.StartMeasurement "Create Folder"
    On Error GoTo ErrorHandler
    IFileOperations_CreateFolder = mFileIO.CreateFolder(folderPath)
    mPerformanceMonitor.EndMeasurement "Create Folder"
    Exit Function
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrFileAccessDenied
        .Description = "フォルダの作成に失敗しました: " & folderPath
        .Source = MODULE_NAME & ".IFileOperations_CreateFolder"
    End With
    Call modError.HandleError(errInfo)
    mPerformanceMonitor.EndMeasurement "Create Folder"
    IFileOperations_CreateFolder = False
End Function

Private Function IFileOperations_DeleteFile(ByVal filePath As String) As Boolean
    InitializeIfNeeded
    mPerformanceMonitor.StartMeasurement "Delete File"
    On Error GoTo ErrorHandler
    IFileOperations_DeleteFile = mFileIO.DeleteFile(filePath)
    mPerformanceMonitor.EndMeasurement "Delete File"
    Exit Function
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrFileAccessDenied
        .Description = "ファイルの削除に失敗しました: " & filePath
        .Source = MODULE_NAME & ".IFileOperations_DeleteFile"
    End With
    Call modError.HandleError(errInfo)
    mPerformanceMonitor.EndMeasurement "Delete File"
    IFileOperations_DeleteFile = False
End Function

Private Function IFileOperations_DeleteFolder(ByVal folderPath As String) As Boolean
    InitializeIfNeeded
    mPerformanceMonitor.StartMeasurement "Delete Folder"
    On Error GoTo ErrorHandler
    IFileOperations_DeleteFolder = mFileIO.DeleteFolder(folderPath)
    mPerformanceMonitor.EndMeasurement "Delete Folder"
    Exit Function
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrFileAccessDenied
        .Description = "フォルダの削除に失敗しました: " & folderPath
        .Source = MODULE_NAME & ".IFileOperations_DeleteFolder"
    End With
    Call modError.HandleError(errInfo)
    mPerformanceMonitor.EndMeasurement "Delete Folder"
    IFileOperations_DeleteFolder = False
End Function

Private Function IFileOperations_GetAbsolutePath(ByVal relativePath As String, _
                                              Optional ByVal basePath As String) As String
    InitializeIfNeeded
    ' FileIOImplでは絶対パスの解決は実装していないため、
    ' FSO（FileSystemObject）の機能を直接使用
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If basePath = "" Then basePath = CurDir()
    IFileOperations_GetAbsolutePath = fso.GetAbsolutePathName(fso.BuildPath(basePath, relativePath))
    Set fso = Nothing
    On Error GoTo 0
End Function

' ======================
' パブリックメソッド
' ======================
Public Function CreateInstance(ByVal lock As ILock, _
                             ByVal performanceMonitor As IPerformanceMonitor, _
                             ByVal fileIO As IFileIO) As IFileOperations
    Dim instance As New FileSystemOperations
    instance.Initialize lock, performanceMonitor, fileIO
    Set CreateInstance = instance
End Function

' ======================
' プライベートヘルパーメソッド
' ======================
Private Function ReadTextFileUsingFileIO(ByVal filePath As String, _
                                       Optional ByVal encoding As String) As String
    Dim file As Object
    Set file = mFileIO.OpenFile(filePath, "READ")
    
    If file Is Nothing Then Exit Function
    
    On Error GoTo ErrorHandler
    ReadTextFileUsingFileIO = file.ReadAll
    mFileIO.CloseFile file
    Exit Function
    
ErrorHandler:
    If Not file Is Nothing Then mFileIO.CloseFile file
    ReadTextFileUsingFileIO = vbNullString
End Function

Private Function WriteTextFileUsingFileIO(ByVal filePath As String, _
                                        ByVal content As String, _
                                        ByVal append As Boolean, _
                                        Optional ByVal encoding As String) As Boolean
    Dim file As Object
    Set file = mFileIO.OpenFile(filePath, IIf(append, "APPEND", "WRITE"))
    
    If file Is Nothing Then Exit Function
    
    On Error GoTo ErrorHandler
    file.Write content
    mFileIO.CloseFile file
    WriteTextFileUsingFileIO = True
    Exit Function
    
ErrorHandler:
    If Not file Is Nothing Then mFileIO.CloseFile file
    WriteTextFileUsingFileIO = False
End Function

Private Function ReadBinaryFileUsingFileIO(ByVal filePath As String) As Byte()
    Dim file As Object
    Set file = mFileIO.OpenFile(filePath, "READ")
    
    If file Is Nothing Then
        ReadBinaryFileUsingFileIO = Array()
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    Dim fileContent As String
    fileContent = file.ReadAll
    mFileIO.CloseFile file
    
    ' 文字列をバイト配列に変換
    Dim bytes() As Byte
    bytes = StrConv(fileContent, vbFromUnicode)
    ReadBinaryFileUsingFileIO = bytes
    Exit Function
    
ErrorHandler:
    If Not file Is Nothing Then mFileIO.CloseFile file
    ReadBinaryFileUsingFileIO = Array()
End Function

Private Function WriteBinaryFileUsingFileIO(ByVal filePath As String, _
                                          ByRef data() As Byte) As Boolean
    Dim file As Object
    Set file = mFileIO.OpenFile(filePath, "WRITE")
    
    If file Is Nothing Then Exit Function
    
    On Error GoTo ErrorHandler
    ' バイト配列を文字列に変換
    Dim content As String
    content = StrConv(data, vbUnicode)
    file.Write content
    mFileIO.CloseFile file
    WriteBinaryFileUsingFileIO = True
    Exit Function
    
ErrorHandler:
    If Not file Is Nothing Then mFileIO.CloseFile file
    WriteBinaryFileUsingFileIO = False
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