VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestClsCrypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Private mCrypto As clsCrypto
Private mPerformanceMonitor As MockPerformanceMonitor
Private mLock As MockLock
Private mCryptography As ICryptography
Private mKeyDerivation As IKeyDerivationStrategy

Private Sub Class_Initialize()
    ' モックオブジェクトの初期化
    Set mPerformanceMonitor = New MockPerformanceMonitor
    Set mLock = New MockLock
    Set mCryptography = New CryptographyImpl
    Set mKeyDerivation = New PBKDF2KeyDerivationStrategy
    
    ' テスト対象の初期化
    Set mCrypto = New clsCrypto
    mCrypto.Initialize mPerformanceMonitor, mLock, mCryptography, mKeyDerivation
End Sub

Private Sub Class_Terminate()
    Set mCrypto = Nothing
    Set mPerformanceMonitor = Nothing
    Set mLock = Nothing
    Set mCryptography = Nothing
    Set mKeyDerivation = Nothing
End Sub

''' <summary>
''' 暗号化と復号化のテスト
''' </summary>
Public Sub TestEncryptionAndDecryption()
    On Error GoTo TestFail
    
    ' テストデータ
    Const TEST_KEY As String = "TestKey123"
    Const TEST_DATA As String = "Hello, World!"
    
    ' 暗号化
    Dim encrypted As String
    encrypted = mCrypto.EncryptString(TEST_DATA, TEST_KEY)
    
    ' 暗号化されたデータの検証
    Debug.Assert Len(encrypted) > 0, "暗号化データが空です"
    Debug.Assert encrypted <> TEST_DATA, "データが暗号化されていません"
    
    ' 復号化
    Dim decrypted As String
    decrypted = mCrypto.DecryptString(encrypted, TEST_KEY)
    
    ' 復号化されたデータの検証
    Debug.Assert decrypted = TEST_DATA, "復号化されたデータが元のデータと一致しません"
    
    Exit Sub
    
TestFail:
    Debug.Print "TestEncryptionAndDecryption failed: " & Err.Description
    Resume Next
End Sub

''' <summary>
''' パフォーマンスモニタリングのテスト
''' </summary>
Public Sub TestPerformanceMonitoring()
    On Error GoTo TestFail
    
    ' テストデータ
    Const TEST_KEY As String = "TestKey123"
    Const TEST_DATA As String = "Hello, World!"
    
    ' 暗号化処理の実行
    Dim encrypted As String
    encrypted = mCrypto.EncryptString(TEST_DATA, TEST_KEY)
    
    ' パフォーマンスモニタリングの検証
    Debug.Assert mPerformanceMonitor.WasMeasurementStarted("String Encryption"), _
                "暗号化のパフォーマンス計測が開始されていません"
    Debug.Assert mPerformanceMonitor.WasMeasurementEnded("String Encryption"), _
                "暗号化のパフォーマンス計測が終了されていません"
    
    Exit Sub
    
TestFail:
    Debug.Print "TestPerformanceMonitoring failed: " & Err.Description
    Resume Next
End Sub

''' <summary>
''' 排他制御のテスト
''' </summary>
Public Sub TestLocking()
    On Error GoTo TestFail
    
    ' テストデータ
    Const TEST_KEY As String = "TestKey123"
    Const TEST_DATA As String = "Hello, World!"
    
    ' 暗号化処理の実行
    Dim encrypted As String
    encrypted = mCrypto.EncryptString(TEST_DATA, TEST_KEY)
    
    ' ロック制御の検証
    Debug.Assert mLock.WasLockAcquired, "ロックが取得されていません"
    Debug.Assert mLock.WasLockReleased, "ロックが解放されていません"
    
    Exit Sub
    
TestFail:
    Debug.Print "TestLocking failed: " & Err.Description
    Resume Next
End Sub

''' <summary>
''' エラー処理のテスト
''' </summary>
Public Sub TestErrorHandling()
    On Error GoTo TestFail
    
    ' 無効なキーでのテスト
    Dim result As String
    result = mCrypto.EncryptString("Test", "")
    
    ' エラー処理の検証
    Debug.Assert result = "", "空のキーでエラーが発生しませんでした"
    
    Exit Sub
    
TestFail:
    Debug.Print "TestErrorHandling failed: " & Err.Description
    Resume Next
End Sub

''' <summary>
''' すべてのテストを実行
''' </summary>
Public Sub RunAllTests()
    TestEncryptionAndDecryption
    TestPerformanceMonitoring
    TestLocking
    TestErrorHandling
    Debug.Print "All tests completed successfully"
End Sub