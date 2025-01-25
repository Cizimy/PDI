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

''' <summary>
''' テキストファイルを読み込みます
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="encoding">文字エンコーディング（オプション）</param>
''' <returns>ファイルの内容、エラー時は空文字列</returns>
''' <remarks>
''' エラー処理要件：
''' - ファイルの存在確認
''' - エンコーディングの検証
''' - ファイルロックの確認
''' - メモリ不足への対応
''' </remarks>
Public Function ReadTextFile(ByVal filePath As String, _
                           Optional ByVal encoding As String = DEFAULT_ENCODING) As String
    InitializeIfNeeded
    
    mLock.AcquireLock
    mPerformanceMonitor.StartMeasurement "Read Text File"
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
    mPerformanceMonitor.EndMeasurement "Read Text File"
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
    
    modError.HandleError errInfo
    ReadTextFile = ""  ' エラー時は空文字列を返す（IFileOperationsの規定に従う）
    Resume CleanUp
End Function

''' <summary>
''' テキストファイルに書き込みます
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="content">書き込む内容</param>
''' <param name="append">追記モード（オプション）</param>
''' <param name="encoding">文字エンコーディング（オプション）</param>
''' <returns>成功時True、失敗時False</returns>
''' <remarks>
''' エラー処理要件：
''' - 書き込み権限の確認
''' - ディスク容量の確認
''' - 既存ファイルのバックアップ
''' - 書き込み失敗時の復旧処理
''' </remarks>
Public Function WriteTextFile(ByVal filePath As String, _
                            ByVal content As String, _
                            Optional ByVal append As Boolean = False, _
                            Optional ByVal encoding As String = DEFAULT_ENCODING) As Boolean
    InitializeIfNeeded
    
    mLock.AcquireLock
    mPerformanceMonitor.StartMeasurement "Write Text File"
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
    mPerformanceMonitor.EndMeasurement "Write Text File"
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
    
    modError.HandleError errInfo
    WriteTextFile = False
    Resume CleanUp
End Function

''' <summary>
''' バイナリファイルを読み込みます
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <returns>ファイルのバイトデータ、エラー時は空配列</returns>
''' <remarks>
''' エラー処理要件：
''' - ファイルサイズの検証
''' - メモリ使用量の監視
''' - 破損ファイルの検出
''' - エラー発生時は空配列を返す
''' </remarks>
Public Function ReadBinaryFile(ByVal filePath As String) As Byte()
    InitializeIfNeeded
    
    mLock.AcquireLock
    mPerformanceMonitor.StartMeasurement "Read Binary File"
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
    mPerformanceMonitor.EndMeasurement "Read Binary File"
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
    
    modError.HandleError errInfo
    ReadBinaryFile = Array()  ' エラー時は空配列を返す
    Resume CleanUp
End Function

''' <summary>
''' バイナリファイルに書き込みます
''' </summary>
''' <param name="filePath">ファイルパス</param>
''' <param name="data">書き込むバイトデータ</param>
''' <returns>成功時True、失敗時False</returns>
''' <remarks>
''' エラー処理要件：
''' - データの整合性チェック
''' - 部分書き込みの防止
''' - 書き込み失敗時のロールバック
''' - エラー発生時はFalseを返す
''' </remarks>
Public Function WriteBinaryFile(ByVal filePath As String, _
                              ByRef data() As Byte) As Boolean
    InitializeIfNeeded
    
    mLock.AcquireLock
    mPerformanceMonitor.StartMeasurement "Write Binary File"
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Binary Access Write As #fileNum
        Put #fileNum, , data
    Close #fileNum
    
    WriteBinaryFile = True
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "Write Binary File"
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
    
    modError.HandleError errInfo
    WriteBinaryFile = False
    Resume CleanUp
End Function

Public Function FileExists(ByVal filePath As String) As Boolean
    InitializeIfNeeded
    
    mLock.AcquireLock
    mPerformanceMonitor.StartMeasurement "FileExists"
    On Error GoTo ErrorHandler
    
    FileExists = (Dir(filePath) <> "")
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "FileExists"
    Exit Function
    
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = GetFileErrorCode(Err.Number)
        .Category = modErrorCodes.ECFileIO
        .Description = "ファイルの存在確認中にエラーが発生しました: " & filePath & vbCrLf & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "FileExists"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    
    modError.HandleError errInfo
    FileExists = False
    Resume CleanUp
End Function
 
Public Function FolderExists(ByVal folderPath As String) As Boolean
    InitializeIfNeeded
    
    mLock.AcquireLock
    mPerformanceMonitor.StartMeasurement "FolderExists"
    On Error GoTo ErrorHandler
    
    FolderExists = (Dir(folderPath, vbDirectory) <> "")
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "FolderExists"
    Exit Function
    
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = GetFileErrorCode(Err.Number)
        .Category = modErrorCodes.ECFileIO
        .Description = "フォルダの存在確認中にエラーが発生しました: " & folderPath & vbCrLf & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "FolderExists"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    
    modError.HandleError errInfo
    FolderExists = False
    Resume CleanUp
End Function
 
Public Function CreateFolder(ByVal folderPath As String) As Boolean
    InitializeIfNeeded
    
    mLock.AcquireLock
    mPerformanceMonitor.StartMeasurement "Create Folder"
    On Error GoTo ErrorHandler
    
    MkDir folderPath
    CreateFolder = True
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "Create Folder"
    Exit Function
    
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = GetFileErrorCode(Err.Number)
        .Category = modErrorCodes.ECFileIO
        .Description = "フォルダの作成中にエラーが発生しました: " & folderPath & vbCrLf & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "CreateFolder"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    
    modError.HandleError errInfo
    CreateFolder = False
    Resume CleanUp
End Function
 
Public Function DeleteFile(ByVal filePath As String) As Boolean
    InitializeIfNeeded
    
    mLock.AcquireLock
    mPerformanceMonitor.StartMeasurement "Delete File"
    On Error GoTo ErrorHandler
    
    Kill filePath
    DeleteFile = True
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "Delete File"
    Exit Function
    
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = GetFileErrorCode(Err.Number)
        .Category = modErrorCodes.ECFileIO
        .Description = "ファイルの削除中にエラーが発生しました: " & filePath & vbCrLf & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "DeleteFile"
        .StackTrace = modStackTrace.GetStackTrace()
    En   .OccurredAt = Now
    End With
    
    modError.HandleError errInfo
    DeleteFile = False
    Resume CleanUp
End Function
 
Public Function DeleteFolder(ByVal folderPath As String) As Boolean
    InitializeIfNeeded
    
    mLock.AcquireLock
    mPerformanceMonitor.StartMeasurement "Delete Folder"
    On Error GoTo ErrorHandler
    
    RmDir folderPath
    DeleteFolder = True
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "Delete Folder"
    Exit Function
    
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = GetFileErrorCode(Err.Number)
        .Category = modErrorCodes.ECFileIO
        .Description = "フォルダの削除中にエラーが発生しました: " & folderPath & vbCrLf & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "DeleteFolder"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    
    modError.HandleError errInfo
    DeleteFolder = False
    Resume CleanUp
End Function

Public Function GetAbsolutePath(ByVal relativePath As String, _
                              Optional ByVal basePath As String) As String
    InitializeIfNeeded
    
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    mPerformanceMonitor.StartMeasurement "GetAbsolutePath"
    
    If Len(basePath) = 0 Then basePath = CurDir
    GetAbsolutePath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(basePath & "\" & relativePath)
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "GetAbsolutePath"
    Exit Function
    
ErrorHandler:
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = GetFileErrorCode(Err.Number)
        .Category = modErrorCodes.ECFileIO
        .Description = "絶対パスの取得中にエラーが発生しました: " & relativePath & vbCrLf & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "GetAbsolutePath"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    
    modError.HandleError errInfo
    GetAbsolutePath = ""
    Resume CleanUp
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