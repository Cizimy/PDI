VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestConfigImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

Private config As ConfigImpl
Private mockIniFile As MockIniFile
Private mockCrypto As MockCryptography
Private mockFileOps As MockFileOperations
Private mockPerfMon As MockPerformanceMonitor

'@ModuleInitialize
Private Sub ModuleInitialize()
    'このメソッドは、モジュール内の各テストの前に1回実行されます
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'このメソッドは、モジュール内のすべてのテストの後に1回実行されます
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'このメソッドは、各テストメソッドの前に実行されます
    Set config = New ConfigImpl
    Set mockIniFile = New MockIniFile
    Set mockCrypto = New MockCryptography
    Set mockFileOps = New MockFileOperations
    Set mockPerfMon = New MockPerformanceMonitor
    
    ' モックオブジェクトの注入
    Set config.IniFile = mockIniFile
    Set config.Crypto = mockCrypto
    Set config.FileOperations = mockFileOps
    Set config.PerformanceMonitor = mockPerfMon
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'このメソッドは、各テストメソッドの後に実行されます
    Set config = Nothing
    Set mockIniFile = Nothing
    Set mockCrypto = Nothing
    Set mockFileOps = Nothing
    Set mockPerfMon = Nothing
End Sub

'@TestMethod("基本機能")
Public Sub TestGetSetSetting()
    On Error GoTo TestFail
    
    'Arrange
    Const KEY As String = "TestKey"
    Const VALUE As String = "TestValue"
    
    'Act
    config.SetSetting KEY, VALUE
    Dim result As Variant
    result = config.GetSetting(KEY)
    
    'Assert
    Assert.AreEqual VALUE, result, "設定値が正しく保存・取得されていません"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("基本機能")
Public Sub TestLoadSettings()
    On Error GoTo TestFail
    
    'Arrange
    Const TEST_PATH As String = "C:\test\config.ini"
    mockIniFile.SetReturnValue "Logging", "LogFilePath", "C:\logs\app.log"
    mockIniFile.SetReturnValue "Logging", "LogLevel", "2" ' llInfo
    
    'Act
    config.LoadSettings TEST_PATH
    
    'Assert
    Assert.AreEqual "C:\logs\app.log", config.LogFilePath, "LogFilePathが正しく読み込まれていません"
    Assert.AreEqual llInfo, config.LogLevel, "LogLevelが正しく読み込まれていません"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("暗号化")
Public Sub TestEncryptedSetting()
    On Error GoTo TestFail
    
    'Arrange
    Const KEY As String = "SecretKey"
    Const VALUE As String = "SecretValue"
    mockCrypto.SetupEncryption "EncryptedValue"
    config.SetEncryptionKey "TestKey"
    
    'Act
    config.SetSetting KEY, VALUE, soEncrypt
    Dim result As Variant
    result = config.GetSetting(KEY, soEncrypt)
    
    'Assert
    Assert.AreEqual VALUE, result, "暗号化された設定値が正しく復号化されていません"
    Assert.IsTrue mockCrypto.EncryptWasCalled, "暗号化メソッドが呼び出されていません"
    Assert.IsTrue mockCrypto.DecryptWasCalled, "復号化メソッドが呼び出されていません"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("環境管理")
Public Sub TestEnvironmentSpecificSettings()
    On Error GoTo TestFail
    
    'Arrange
    mockIniFile.SetReturnValue "Environment_development", "LogFilePath", "C:\dev\logs\app.log"
    
    'Act
    config.SetEnvironment "development"
    config.LoadSettings "test.ini", loApplyEnvironment
    
    'Assert
    Assert.AreEqual "C:\dev\logs\app.log", config.LogFilePath, "環境固有の設定が正しく適用されていません"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("バリデーション")
Public Sub TestSettingsValidation()
    On Error GoTo TestFail
    
    'Arrange
    config.MaxRetryCount = -1 ' 無効な値
    
    'Act
    Dim result As ValidationResult
    Set result = config.ValidateSettings(voDataType)
    
    'Assert
    Assert.IsFalse result.IsValid, "無効な設定値がバリデーションを通過しています"
    Assert.IsTrue result.Errors.Count > 0, "バリデーションエラーが記録されていません"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("パフォーマンス")
Public Sub TestPerformanceMetrics()
    On Error GoTo TestFail
    
    'Arrange
    config.PerformanceMonitoringEnabled = True
    
    'Act
    config.SetSetting "TestKey", "TestValue"
    Dim metrics As Collection
    Set metrics = config.GetPerformanceMetrics
    
    'Assert
    Assert.IsTrue metrics.Count > 0, "パフォーマンスメトリクスが記録されていません"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("バックアップ/復元")
Public Sub TestBackupAndRestore()
    On Error GoTo TestFail
    
    'Arrange
    Const BACKUP_PATH As String = "C:\backup\config.bak"
    mockFileOps.SetupFileExists True
    
    'Act
    Dim backupResult As Boolean
    backupResult = config.CreateBackup(BACKUP_PATH)
    
    'Assert
    Assert.IsTrue backupResult, "バックアップの作成に失敗しました"
    Assert.IsTrue mockFileOps.CopyFileWasCalled, "ファイルのコピーが実行されていません"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("履歴管理")
Public Sub TestSettingHistory()
    On Error GoTo TestFail
    
    'Arrange
    Const KEY As String = "TestKey"
    
    'Act
    config.SetSetting KEY, "Value1", soLogHistory
    config.SetSetting KEY, "Value2", soLogHistory
    Dim history As Collection
    Set history = config.GetSettingHistory(KEY)
    
    'Assert
    Assert.AreEqual 2, history.Count, "設定履歴が正しく記録されていません"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("エラー処理")
Public Sub TestInvalidEncryptionKey()
    On Error GoTo TestFail
    
    'Arrange
    Const KEY As String = "SecretKey"
    Const VALUE As String = "SecretValue"
    
    'Act
    On Error Resume Next
    config.SetSetting KEY, VALUE, soEncrypt
    Dim errNumber As Long
    errNumber = Err.Number
    On Error GoTo TestFail
    
    'Assert
    Assert.IsTrue errNumber <> 0, "暗号化キーが設定されていない場合にエラーが発生していません"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub

'@TestMethod("エラー処理")
Public Sub TestFileAccessError()
    On Error GoTo TestFail
    
    'Arrange
    mockFileOps.SetupFileExists False
    mockFileOps.SetupFileAccessError True
    
    'Act
    On Error Resume Next
    config.LoadSettings "nonexistent.ini"
    Dim errNumber As Long
    errNumber = Err.Number
    On Error GoTo TestFail
    
    'Assert
    Assert.IsTrue errNumber <> 0, "存在しないファイルへのアクセスでエラーが発生していません"
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "テストが失敗しました: " & Err.Description
End Sub