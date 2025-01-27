VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestFileSystemOperations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "TestFileSystemOperations"

' モックオブジェクト
Private mockLock As MockLock
Private mockPerformanceMonitor As MockPerformanceMonitor
Private mockFileIO As MockFileIO
Private fileOps As FileSystemOperations

'@TestInitialize
Private Sub TestInitialize()
    ' モックオブジェクトの初期化
    Set mockLock = New MockLock
    Set mockPerformanceMonitor = New MockPerformanceMonitor
    Set mockFileIO = New MockFileIO
    Set fileOps = New FileSystemOperations
    
    ' 依存性の注入
    fileOps.Initialize mockLock, mockPerformanceMonitor, mockFileIO
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set mockLock = Nothing
    Set mockPerformanceMonitor = Nothing
    Set mockFileIO = Nothing
    Set fileOps = Nothing
End Sub

'@TestMethod("初期化")
Private Sub TestInitialization()
    On Error GoTo TestFail
    
    ' 未初期化の状態でメソッドを呼び出すとエラーになることを確認
    Dim uninitializedOps As New FileSystemOperations
    Dim success As Boolean
    
    On Error Resume Next
    success = uninitializedOps.FileExists("test.txt")
    
    Debug.Assert Err.Number <> 0, "未初期化状態でのメソッド呼び出しはエラーになるべき"
    On Error GoTo 0
    
    ' 正しく初期化された状態ではエラーにならないことを確認
    On Error Resume Next
    success = fileOps.FileExists("test.txt")
    Debug.Assert Err.Number = 0, "初期化済み状態でのメソッド呼び出しはエラーにならない"
    
    Exit Sub
    
TestFail:
    Debug.Print "TestInitialization failed: " & Err.Description
    Resume Next
End Sub

'@TestMethod("ファイル操作")
Private Sub TestFileOperations()
    On Error GoTo TestFail
    
    ' FileExistsのテスト
    mockFileIO.SetFileExists "test.txt", True
    Debug.Assert fileOps.FileExists("test.txt"), "FileExists should return True"
    
    ' CreateFolderのテスト
    mockFileIO.SetCreateFolderResult "testFolder", True
    Debug.Assert fileOps.CreateFolder("testFolder"), "CreateFolder should return True"
    
    ' パフォーマンスモニタリングの確認
    Debug.Assert mockPerformanceMonitor.WasMeasurementStarted("Create Folder"), _
                 "パフォーマンスモニタリングが開始されていない"
    Debug.Assert mockPerformanceMonitor.WasMeasurementEnded("Create Folder"), _
                 "パフォーマンスモニタリングが終了されていない"
    
    Exit Sub
    
TestFail:
    Debug.Print "TestFileOperations failed: " & Err.Description
    Resume Next
End Sub

'@TestMethod("エラー処理")
Private Sub TestErrorHandling()
    On Error GoTo TestFail
    
    ' ファイル読み込みエラーのテスト
    mockFileIO.SetOpenFileResult "error.txt", Nothing
    Dim content As String
    content = fileOps.ReadTextFile("error.txt")
    Debug.Assert content = vbNullString, "エラー時は空文字列を返すべき"
    
    ' ファイル書き込みエラーのテスト
    mockFileIO.SetOpenFileResult "readonly.txt", Nothing
    Dim writeSuccess As Boolean
    writeSuccess = fileOps.WriteTextFile("readonly.txt", "test")
    Debug.Assert Not writeSuccess, "エラー時はFalseを返すべき"
    
    Exit Sub
    
TestFail:
    Debug.Print "TestErrorHandling failed: " & Err.Description
    Resume Next
End Sub

'@TestMethod("バイナリファイル操作")
Private Sub TestBinaryOperations()
    On Error GoTo TestFail
    
    ' バイナリファイル読み込みのテスト
    Dim testData(0 To 2) As Byte
    testData(0) = 1: testData(1) = 2: testData(2) = 3
    
    mockFileIO.SetBinaryFileContent "test.bin", testData
    Dim readData() As Byte
    readData = fileOps.ReadBinaryFile("test.bin")
    
    Debug.Assert UBound(readData) = UBound(testData), "バイナリデータのサイズが一致しない"
    Debug.Assert readData(0) = testData(0), "バイナリデータの内容が一致しない"
    
    Exit Sub
    
TestFail:
    Debug.Print "TestBinaryOperations failed: " & Err.Description
    Resume Next
End Sub

'@TestMethod("パス操作")
Private Sub TestPathOperations()
    On Error GoTo TestFail
    
    ' 相対パスから絶対パスへの変換テスト
    Dim absolutePath As String
    absolutePath = fileOps.GetAbsolutePath("test.txt", "C:\temp")
    Debug.Assert absolutePath = "C:\temp\test.txt", "絶対パスの解決が正しくない"
    
    Exit Sub
    
TestFail:
    Debug.Print "TestPathOperations failed: " & Err.Description
    Resume Next
End Sub