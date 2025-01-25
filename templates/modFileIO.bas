Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "modFileIO"

' ======================
' 定数定義
' ======================
Private Const MAX_BUFFER_SIZE As Long = 1024
Private Const DEFAULT_ENCODING As String = "UTF-8"

' ======================
' プライベート変数
' ======================
Private mLock As clsLock
Private mPerformanceMonitor As clsPerformanceMonitor
Private mIsInitialized As Boolean

' ======================
' 初期化・終了処理
' ======================
Public Property Get IsInitialized() As Boolean
    IsInitialized = mIsInitialized
End Property

Private Sub InitializeIfNeeded()
    If Not mIsInitialized Then InitializeModule
End Sub

Public Sub InitializeModule()
    If mIsInitialized Then Exit Sub
    
    Set mLock = New clsLock
    Set mPerformanceMonitor = New clsPerformanceMonitor
    
    mIsInitialized = True
End Sub

Public Sub TerminateModule()
    If Not mIsInitialized Then Exit Sub
    
    Set mLock = Nothing
    Set mPerformanceMonitor = Nothing
    
    mIsInitialized = False
End Sub

' ======================
' 公開メソッド
' ======================
Public Function ReadTextFile(ByVal filePath As String, _
                           Optional ByVal encoding As String = DEFAULT_ENCODING) As String
    InitializeIfNeeded
    
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    If Not FileExists(filePath) Then
        RaiseFileError modErrorCodes.ErrFileNotFound, "ファイルが見つかりません: " & filePath
    End If
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Input As #fileNum Encoding encoding
        ReadTextFile = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    
CleanUp:
    mLock.ReleaseLock
    Exit Function
    
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = GetFileErrorCode(Err.Number)
        .Category = modErrorCodes.ECFileIO
        .Description = Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "ReadTextFile"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    
    modCommon.HandleError errInfo
    Resume CleanUp
End Function

Public Function WriteTextFile(ByVal filePath As String, _
                            ByVal content As String, _
                            Optional ByVal append As Boolean = False, _
                            Optional ByVal encoding As String = DEFAULT_ENCODING) As Boolean
    InitializeIfNeeded
    
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    If append Then
        Open filePath For Append As #fileNum Encoding encoding
    Else
        Open filePath For Output As #fileNum Encoding encoding
    End If
    
    Print #fileNum, content
    Close #fileNum
    
    WriteTextFile = True
    
CleanUp:
    mLock.ReleaseLock
    Exit Function
    
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = GetFileErrorCode(Err.Number)
        .Category = modErrorCodes.ECFileIO
        .Description = Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "WriteTextFile"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    
    modCommon.HandleError errInfo
    WriteTextFile = False
    Resume CleanUp
End Function

Public Function ReadBinaryFile(ByVal filePath As String) As Byte()
    InitializeIfNeeded
    
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    If Not FileExists(filePath) Then
        RaiseFileError modErrorCodes.ErrFileNotFound, "ファイルが見つかりません: " & filePath
    End If
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Binary Access Read As #fileNum
        Dim fileData() As Byte
        ReDim fileData(LOF(fileNum) - 1)
        Get #fileNum, , fileData
    Close #fileNum
    
    ReadBinaryFile = fileData
    
CleanUp:
    mLock.ReleaseLock
    Exit Function
    
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = GetFileErrorCode(Err.Number)
        .Category = modErrorCodes.ECFileIO
        .Description = Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "ReadBinaryFile"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    
    modCommon.HandleError errInfo
    Resume CleanUp
End Function

Public Function WriteBinaryFile(ByVal filePath As String, _
                              ByRef data() As Byte) As Boolean
    InitializeIfNeeded
    
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Binary Access Write As #fileNum
        Put #fileNum, , data
    Close #fileNum
    
    WriteBinaryFile = True
    
CleanUp:
    mLock.ReleaseLock
    Exit Function
    
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = GetFileErrorCode(Err.Number)
        .Category = modErrorCodes.ECFileIO
        .Description = Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "WriteBinaryFile"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    
    modCommon.HandleError errInfo
    WriteBinaryFile = False
    Resume CleanUp
End Function

Public Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function

Public Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (Dir(folderPath, vbDirectory) <> "")
    On Error GoTo 0
End Function

Public Function CreateFolder(ByVal folderPath As String) As Boolean
    On Error Resume Next
    MkDir folderPath
    CreateFolder = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function DeleteFile(ByVal filePath As String) As Boolean
    On Error Resume Next
    Kill filePath
    DeleteFile = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function DeleteFolder(ByVal folderPath As String) As Boolean
    On Error Resume Next
    RmDir folderPath
    DeleteFolder = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function GetAbsolutePath(ByVal relativePath As String, _
                              Optional ByVal basePath As String) As String
    If Len(basePath) = 0 Then basePath = CurDir
    GetAbsolutePath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(basePath & "\" & relativePath)
End Function

' ======================
' プライベートメソッド
' ======================
Private Function GetFileErrorCode(ByVal errNumber As Long) As ErrorCode
    Select Case errNumber
        Case 53 ' File not found
            GetFileErrorCode = modErrorCodes.ErrFileNotFound
        Case 70 ' Permission denied
            GetFileErrorCode = modErrorCodes.ErrFileAccessDenied
        Case 75, 76 ' Path/File access error
            GetFileErrorCode = modErrorCodes.ErrFileAccessDenied
        Case Else
            GetFileErrorCode = modErrorCodes.ErrUnexpected
    End Select
End Function

Private Sub RaiseFileError(ByVal errorCode As ErrorCode, ByVal description As String)
    Err.Raise errorCode, MODULE_NAME, description
End Sub

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
        TerminateModule
        InitializeModule
    End Sub
#End If