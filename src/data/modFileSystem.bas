Attribute VB_Name = "modFileSystem"
Option Explicit

'*******************************************************************************
' モジュール: modFileSystem
' 目的：     ファイルシステム操作の中央管理
' 作成日：   2025/01/17
'*******************************************************************************

' ファイルシステムオブジェクト
Private mFSO As Object  ' Scripting.FileSystemObject

'*******************************************************************************
' 目的：    モジュールの初期化
' 引数：    なし
' 戻り値：  なし
'*******************************************************************************
Public Sub Initialize()
    On Error GoTo ErrorHandler
    
    ' FileSystemObjectの作成
    Set mFSO = CreateObject("Scripting.FileSystemObject")
    modLogger.Info "ファイルシステムモジュールを初期化しました。", "FileSystem.Initialize"
    Exit Sub
    
ErrorHandler:
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "FileSystem.Initialize", _
                               etFileSystem
End Sub

'*******************************************************************************
' 目的：    ファイルの存在確認
' 引数：    filePath - ファイルパス
' 戻り値：  存在する場合True
'*******************************************************************************
Public Function FileExists(ByVal filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    If mFSO Is Nothing Then Initialize
    FileExists = mFSO.FileExists(filePath)
    Exit Function
    
ErrorHandler:
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "FileSystem.FileExists", _
                               etFileSystem, _
                               "Path: " & filePath
End Function

'*******************************************************************************
' 目的：    ディレクトリの存在確認
' 引数：    dirPath - ディレクトリパス
' 戻り値：  存在する場合True
'*******************************************************************************
Public Function DirectoryExists(ByVal dirPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    If mFSO Is Nothing Then Initialize
    DirectoryExists = mFSO.FolderExists(dirPath)
    Exit Function
    
ErrorHandler:
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "FileSystem.DirectoryExists", _
                               etFileSystem, _
                               "Path: " & dirPath
End Function

'*******************************************************************************
' 目的：    ファイルの読み込み
' 引数：    filePath - ファイルパス
' 戻り値：  ファイルの内容
'*******************************************************************************
Public Function ReadFile(ByVal filePath As String) As String
    On Error GoTo ErrorHandler
    
    If mFSO Is Nothing Then Initialize
    
    Dim textStream As Object
    
    ' ファイルが存在しない場合はエラー
    If Not FileExists(filePath) Then
        Err.Raise vbObjectError + 1000, "FileSystem.ReadFile", _
                  "指定されたファイルが存在しません: " & filePath
    End If
    
    ' ファイルを読み込み
    Set textStream = mFSO.OpenTextFile(filePath, 1) ' ForReading = 1
    ReadFile = textStream.ReadAll
    textStream.Close
    
    modLogger.Debug "ファイルを読み込みました: " & filePath, "FileSystem.ReadFile"
    Exit Function
    
ErrorHandler:
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "FileSystem.ReadFile", _
                               etFileSystem, _
                               "Path: " & filePath
End Function

'*******************************************************************************
' 目的：    ファイルの書き込み
' 引数：    filePath - ファイルパス
'           content - 書き込む内容
'           append - 追記モード（オプション）
' 戻り値：  なし
'*******************************************************************************
Public Sub WriteFile(ByVal filePath As String, _
                    ByVal content As String, _
                    Optional ByVal append As Boolean = False)
                    
    On Error GoTo ErrorHandler
    
    If mFSO Is Nothing Then Initialize
    
    Dim textStream As Object
    Dim mode As Integer
    
    ' 書き込みモードの設定
    If append Then
        mode = 8  ' ForAppending = 8
    Else
        mode = 2  ' ForWriting = 2
    End If
    
    ' ディレクトリが存在しない場合は作成
    Dim parentPath As String
    parentPath = mFSO.GetParentFolderName(filePath)
    If parentPath <> "" And Not DirectoryExists(parentPath) Then
        CreateDirectory parentPath
    End If
    
    ' ファイルに書き込み
    Set textStream = mFSO.OpenTextFile(filePath, mode, True)
    textStream.Write content
    textStream.Close
    
    modLogger.Debug "ファイルを書き込みました: " & filePath, "FileSystem.WriteFile"
    Exit Sub
    
ErrorHandler:
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "FileSystem.WriteFile", _
                               etFileSystem, _
                               "Path: " & filePath
End Sub

'*******************************************************************************
' 目的：    ディレクトリの作成
' 引数：    dirPath - ディレクトリパス
' 戻り値：  なし
'*******************************************************************************
Public Sub CreateDirectory(ByVal dirPath As String)
    On Error GoTo ErrorHandler
    
    If mFSO Is Nothing Then Initialize
    
    ' ディレクトリが存在しない場合のみ作成
    If Not DirectoryExists(dirPath) Then
        mFSO.CreateFolder dirPath
        modLogger.Debug "ディレクトリを作成しました: " & dirPath, "FileSystem.CreateDirectory"
    End If
    
    Exit Sub
    
ErrorHandler:
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "FileSystem.CreateDirectory", _
                               etFileSystem, _
                               "Path: " & dirPath
End Sub

'*******************************************************************************
' 目的：    ファイルのコピー
' 引数：    sourcePath - コピー元パス
'           destPath - コピー先パス
'           overwrite - 上書き（オプション）
' 戻り値：  なし
'*******************************************************************************
Public Sub CopyFile(ByVal sourcePath As String, _
                   ByVal destPath As String, _
                   Optional ByVal overwrite As Boolean = False)
                   
    On Error GoTo ErrorHandler
    
    If mFSO Is Nothing Then Initialize
    
    ' コピー元ファイルの存在確認
    If Not FileExists(sourcePath) Then
        Err.Raise vbObjectError + 1001, "FileSystem.CopyFile", _
                  "コピー元ファイルが存在しません: " & sourcePath
    End If
    
    ' コピー先のディレクトリが存在しない場合は作成
    Dim destDir As String
    destDir = mFSO.GetParentFolderName(destPath)
    If destDir <> "" And Not DirectoryExists(destDir) Then
        CreateDirectory destDir
    End If
    
    ' ファイルをコピー
    mFSO.CopyFile sourcePath, destPath, overwrite
    
    modLogger.Debug "ファイルをコピーしました: " & sourcePath & " -> " & destPath, _
                   "FileSystem.CopyFile"
    Exit Sub
    
ErrorHandler:
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "FileSystem.CopyFile", _
                               etFileSystem, _
                               "Source: " & sourcePath & ", Dest: " & destPath
End Sub

'*******************************************************************************
' 目的：    ファイルの移動
' 引数：    sourcePath - 移動元パス
'           destPath - 移動先パス
' 戻り値：  なし
'*******************************************************************************
Public Sub MoveFile(ByVal sourcePath As String, ByVal destPath As String)
    On Error GoTo ErrorHandler
    
    If mFSO Is Nothing Then Initialize
    
    ' 移動元ファイルの存在確認
    If Not FileExists(sourcePath) Then
        Err.Raise vbObjectError + 1002, "FileSystem.MoveFile", _
                  "移動元ファイルが存在しません: " & sourcePath
    End If
    
    ' 移動先のディレクトリが存在しない場合は作成
    Dim destDir As String
    destDir = mFSO.GetParentFolderName(destPath)
    If destDir <> "" And Not DirectoryExists(destDir) Then
        CreateDirectory destDir
    End If
    
    ' ファイルを移動
    mFSO.MoveFile sourcePath, destPath
    
    modLogger.Debug "ファイルを移動しました: " & sourcePath & " -> " & destPath, _
                   "FileSystem.MoveFile"
    Exit Sub
    
ErrorHandler:
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "FileSystem.MoveFile", _
                               etFileSystem, _
                               "Source: " & sourcePath & ", Dest: " & destPath
End Sub

'*******************************************************************************
' 目的：    ファイルの削除
' 引数：    filePath - ファイルパス
' 戻り値：  なし
'*******************************************************************************
Public Sub DeleteFile(ByVal filePath As String)
    On Error GoTo ErrorHandler
    
    If mFSO Is Nothing Then Initialize
    
    ' ファイルが存在する場合のみ削除
    If FileExists(filePath) Then
        mFSO.DeleteFile filePath, False ' Force = False
        modLogger.Debug "ファイルを削除しました: " & filePath, "FileSystem.DeleteFile"
    End If
    
    Exit Sub
    
ErrorHandler:
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "FileSystem.DeleteFile", _
                               etFileSystem, _
                               "Path: " & filePath
End Sub

'*******************************************************************************
' 目的：    ディレクトリの削除
' 引数：    dirPath - ディレクトリパス
'           recursive - サブディレクトリも含めて削除
' 戻り値：  なし
'*******************************************************************************
Public Sub DeleteDirectory(ByVal dirPath As String, _
                         Optional ByVal recursive As Boolean = False)
                         
    On Error GoTo ErrorHandler
    
    If mFSO Is Nothing Then Initialize
    
    ' ディレクトリが存在する場合のみ削除
    If DirectoryExists(dirPath) Then
        mFSO.DeleteFolder dirPath, False ' Force = False
        modLogger.Debug "ディレクトリを削除しました: " & dirPath, "FileSystem.DeleteDirectory"
    End If
    
    Exit Sub
    
ErrorHandler:
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "FileSystem.DeleteDirectory", _
                               etFileSystem, _
                               "Path: " & dirPath
End Sub

'*******************************************************************************
' 目的：    ファイル一覧の取得
' 引数：    dirPath - ディレクトリパス
'           pattern - 検索パターン（オプション）
' 戻り値：  ファイル名の配列
'*******************************************************************************
Public Function GetFiles(ByVal dirPath As String, _
                        Optional ByVal pattern As String = "*.*") As Variant
                        
    On Error GoTo ErrorHandler
    
    If mFSO Is Nothing Then Initialize
    
    Dim folder As Object
    Dim file As Object
    Dim fileList() As String
    Dim count As Long
    
    ' ディレクトリの存在確認
    If Not DirectoryExists(dirPath) Then
        Err.Raise vbObjectError + 1003, "FileSystem.GetFiles", _
                  "指定されたディレクトリが存在しません: " & dirPath
    End If
    
    ' ディレクトリ内のファイルを列挙
    Set folder = mFSO.GetFolder(dirPath)
    ReDim fileList(0 To folder.Files.Count - 1)
    
    count = 0
    For Each file In folder.Files
        ' パターンに一致するファイルのみ追加
        If file.Name Like pattern Then
            fileList(count) = file.Name
            count = count + 1
        End If
    Next file
    
    ' 配列のサイズを実際のファイル数に調整
    If count > 0 Then
        ReDim Preserve fileList(0 To count - 1)
        GetFiles = fileList
    Else
        GetFiles = Array()
    End If
    
    modLogger.Debug "ファイル一覧を取得しました: " & dirPath & " (" & count & "件)", _
                   "FileSystem.GetFiles"
    Exit Function
    
ErrorHandler:
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "FileSystem.GetFiles", _
                               etFileSystem, _
                               "Path: " & dirPath
End Function

'*******************************************************************************
' 目的：    ファイル情報の取得
' 引数：    filePath - ファイルパス
' 戻り値：  Dictionary（サイズ、作成日時、更新日時など）
'*******************************************************************************
Public Function GetFileInfo(ByVal filePath As String) As Object
    On Error GoTo ErrorHandler
    
    If mFSO Is Nothing Then Initialize
    
    Dim file As Object
    Dim info As Object
    
    ' ファイルの存在確認
    If Not FileExists(filePath) Then
        Err.Raise vbObjectError + 1004, "FileSystem.GetFileInfo", _
                  "指定されたファイルが存在しません: " & filePath
    End If
    
    ' ファイル情報を取得
    Set file = mFSO.GetFile(filePath)
    Set info = CreateObject("Scripting.Dictionary")
    
    With file
        info.Add "Name", .Name
        info.Add "Path", .Path
        info.Add "Size", .Size
        info.Add "DateCreated", .DateCreated
        info.Add "DateLastModified", .DateLastModified
        info.Add "DateLastAccessed", .DateLastAccessed
        info.Add "Type", .Type
    End With
    
    Set GetFileInfo = info
    
    modLogger.Debug "ファイル情報を取得しました: " & filePath, "FileSystem.GetFileInfo"
    Exit Function
    
ErrorHandler:
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "FileSystem.GetFileInfo", _
                               etFileSystem, _
                               "Path: " & filePath
End Function
