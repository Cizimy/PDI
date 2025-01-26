Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "modTestRunner"

' ======================
' テスト実行モジュール
' ======================
Public Sub RunAllTests()
    On Error GoTo ErrorHandler
    
    modTestUtility.InitializeTestModule

    ' エラーハンドリング関連のテスト
    RunErrorHandlingTests
    
    ' ロギング関連のテスト
    RunLoggingTests
    
    ' ファイル操作関連のテスト
    RunFileOperationsTests
    
    ' バリデーション関連のテスト
    RunValidationTests
    
    ' ユーティリティ関連のテスト
    RunUtilityTests
    
    ' データベース関連のテスト
    RunDatabaseTests
    
    ' セキュリティ関連のテスト
    RunSecurityTests
    
    ' パフォーマンス関連のテスト
    RunPerformanceTests
    
    ' テストレポートの出力
    Debug.Print modTestUtility.GenerateTestReport
    modTestUtility.CleanupTestModule
    Exit Sub

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "テスト実行中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "RunAllTests"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
End Sub

' ======================
' エラーハンドリングテスト
' ======================
Private Sub RunErrorHandlingTests()
    On Error GoTo ErrorHandler
    
    modTestUtility.StartTest "Test_ErrorHandlers", "エラーハンドラーのテスト"
    Test_ErrorHandlers
    modTestUtility.EndTest modTestUtility.ResultPass
    
    Exit Sub

ErrorHandler:
    modTestUtility.EndTest modTestUtility.ResultFail, "エラーハンドリングテストでエラー発生: " & Err.Description
End Sub

' ======================
' ロギングテスト
' ======================
Private Sub RunLoggingTests()
    On Error GoTo ErrorHandler
    
    modTestUtility.StartTest "Test_Loggers", "ロガーのテスト"
    Test_Loggers
    modTestUtility.EndTest modTestUtility.ResultPass
    
    Exit Sub

ErrorHandler:
    modTestUtility.EndTest modTestUtility.ResultFail, "ロギングテストでエラー発生: " & Err.Description
End Sub

' ======================
' ファイル操作テスト
' ======================
Private Sub RunFileOperationsTests()
    On Error GoTo ErrorHandler
    
    modTestUtility.StartTest "Test_FileOperations", "ファイル操作のテスト"
    Test_FileOperations
    modTestUtility.EndTest modTestUtility.ResultPass
    
    Exit Sub

ErrorHandler:
    modTestUtility.EndTest modTestUtility.ResultFail, "ファイル操作テストでエラー発生: " & Err.Description
End Sub

' ======================
' バリデーションテスト
' ======================
Private Sub RunValidationTests()
    On Error GoTo ErrorHandler
    
    modTestUtility.StartTest "Test_Validators", "バリデーターのテスト"
    Test_Validators
    modTestUtility.EndTest modTestUtility.ResultPass
    
    Exit Sub

ErrorHandler:
    modTestUtility.EndTest modTestUtility.ResultFail, "バリデーションテストでエラー発生: " & Err.Description
End Sub

' ======================
' ユーティリティテスト
' ======================
Private Sub RunUtilityTests()
    On Error GoTo ErrorHandler
    
    modTestUtility.StartTest "Test_MathUtils", "数学ユーティリティのテスト"
    Test_MathUtils
    modTestUtility.EndTest modTestUtility.ResultPass
    
    modTestUtility.StartTest "Test_StringUtils", "文字列ユーティリティのテスト"
    Test_StringUtils
    modTestUtility.EndTest modTestUtility.ResultPass
    
    modTestUtility.StartTest "Test_DateUtils", "日付ユーティリティのテスト"
    Test_DateUtils
    modTestUtility.EndTest modTestUtility.ResultPass
    
    Exit Sub

ErrorHandler:
    modTestUtility.EndTest modTestUtility.ResultFail, "ユーティリティテストでエラー発生: " & Err.Description
End Sub

' ======================
' データベーステスト
' ======================
Private Sub RunDatabaseTests()
    On Error GoTo ErrorHandler
    
    modTestUtility.StartTest "Test_DatabaseUtils", "データベースユーティリティのテスト"
    Test_DatabaseUtils
    modTestUtility.EndTest modTestUtility.ResultPass
    
    modTestUtility.StartTest "Test_ConnectionPool", "接続プールのテスト"
    Test_ConnectionPool
    modTestUtility.EndTest modTestUtility.ResultPass
    
    Exit Sub

ErrorHandler:
    modTestUtility.EndTest modTestUtility.ResultFail, "データベーステストでエラー発生: " & Err.Description
End Sub

' ======================
' セキュリティテスト
' ======================
Private Sub RunSecurityTests()
    On Error GoTo ErrorHandler
    
    modTestUtility.StartTest "Test_Crypto", "暗号化のテスト"
    Test_Crypto
    modTestUtility.EndTest modTestUtility.ResultPass
    
    Exit Sub

ErrorHandler:
    modTestUtility.EndTest modTestUtility.ResultFail, "セキュリティテストでエラー発生: " & Err.Description
End Sub

' ======================
' パフォーマンステスト
' ======================
Private Sub RunPerformanceTests()
    On Error GoTo ErrorHandler
    
    modTestUtility.StartTest "Test_PerformanceMonitor", "パフォーマンスモニターのテスト"
    Test_PerformanceMonitor
    modTestUtility.EndTest modTestUtility.ResultPass
    
    modTestUtility.StartTest "Test_Lock", "ロック機能のテスト"
    Test_Lock
    modTestUtility.EndTest modTestUtility.ResultPass
    
    modTestUtility.StartTest "Test_CallStack", "コールスタックのテスト"
    Test_CallStack
    modTestUtility.EndTest modTestUtility.ResultPass
    
    Exit Sub

ErrorHandler:
    modTestUtility.EndTest modTestUtility.ResultFail, "パフォーマンステストでエラー発生: " & Err.Description
End Sub

#If DEBUG Then
    ' ======================
    ' エラーハンドリングテストケース
    ' ======================
    Private Sub Test_ErrorHandlers()
        ' DatabaseConnectionErrorHandlerのテスト
        Dim dbErrorHandler As New DatabaseConnectionErrorHandler
        modTestUtility.AssertTrue TypeOf dbErrorHandler Is IErrorHandler, "DatabaseConnectionErrorHandlerがIErrorHandlerを実装していることを確認"
        
        ' FileNotFoundErrorHandlerのテスト
        Dim fileNotFoundHandler As New FileNotFoundErrorHandler
        modTestUtility.AssertTrue TypeOf fileNotFoundHandler Is IErrorHandler, "FileNotFoundErrorHandlerがIErrorHandlerを実装していることを確認"
        
        ' InvalidInputErrorHandlerのテスト
        Dim invalidInputHandler As New InvalidInputErrorHandler
        modTestUtility.AssertTrue TypeOf invalidInputHandler Is IErrorHandler, "InvalidInputErrorHandlerがIErrorHandlerを実装していることを確認"
        
        ' エラーコードの検証
        Dim errInfo As New ErrorInfo
        modTestUtility.AssertTrue dbErrorHandler.HandleError(errInfo) <> 0, "エラーハンドリングの結果確認"
    End Sub
    
    ' ======================
    ' ロギングテストケース
    ' ======================
    Private Sub Test_Loggers()
        ' FileLoggerのテスト
        Dim fileLogger As New FileLogger
        modTestUtility.AssertTrue TypeOf fileLogger Is ILogger, "FileLoggerがILoggerを実装していることを確認"
        
        ' MockLoggerのテスト
        Dim mockLogger As New MockLogger
        modTestUtility.AssertTrue TypeOf mockLogger Is ILogger, "MockLoggerがILoggerを実装していることを確認"
        
        ' clsLoggerのテスト
        Dim logger As New clsLogger
        modTestUtility.AssertTrue TypeOf logger Is ILogger, "clsLoggerがILoggerを実装していることを確認"
        
        ' DefaultLoggerSettingsのテスト
        Dim settings As New DefaultLoggerSettings
        modTestUtility.AssertTrue TypeOf settings Is ILoggerSettings, "DefaultLoggerSettingsがILoggerSettingsを実装していることを確認"
        
        ' ログ出力のテスト
        fileLogger.LogMessage "テストメッセージ", LogLevel.Info
        mockLogger.LogMessage "テストメッセージ", LogLevel.Info
        logger.LogMessage "テストメッセージ", LogLevel.Info
    End Sub
    
    ' ======================
    ' ファイル操作テストケース
    ' ======================
    Private Sub Test_FileOperations()
        ' FileSystemOperationsのテスト
        Dim fileOps As New FileSystemOperations
        modTestUtility.AssertTrue TypeOf fileOps Is IFileOperations, "FileSystemOperationsがIFileOperationsを実装していることを確認"
        
        ' modFileIOのテスト
        Dim testPath As String
        testPath = "test.txt"
        
        modFileIO.WriteTextFile testPath, "テストデータ"
        modTestUtility.AssertTrue modFileIO.FileExists(testPath), "ファイル作成の確認"
        
        Dim content As String
        content = modFileIO.ReadTextFile(testPath)
        modTestUtility.AssertEqual "テストデータ", content, "ファイル内容の確認"
        
        modFileIO.DeleteFile testPath
        modTestUtility.AssertFalse modFileIO.FileExists(testPath), "ファイル削除の確認"
    End Sub
    
    ' ======================
    ' バリデーションテストケース
    ' ======================
    Private Sub Test_Validators()
        ' StringValidatorのテスト
        Dim strValidator As New StringValidator
        modTestUtility.AssertTrue TypeOf strValidator Is IValidator, "StringValidatorがIValidatorを実装していることを確認"
        modTestUtility.AssertTrue strValidator.Validate("テスト"), "有効な文字列の検証"
        modTestUtility.AssertFalse strValidator.Validate(""), "空文字列の検証"
        
        ' DateValidatorのテスト
        Dim dateValidator As New DateValidator
        modTestUtility.AssertTrue TypeOf dateValidator Is IValidator, "DateValidatorがIValidatorを実装していることを確認"
        modTestUtility.AssertTrue dateValidator.Validate(Date), "有効な日付の検証"
        modTestUtility.AssertFalse dateValidator.Validate(Empty), "無効な日付の検証"
    End Sub
    
    ' ======================
    ' ユーティリティテストケース
    ' ======================
    Private Sub Test_MathUtils()
        modTestUtility.AssertEqual 10, modMathUtils.Add(7, 3), "加算のテスト"
        modTestUtility.AssertEqual 4, modMathUtils.Subtract(7, 3), "減算のテスト"
        modTestUtility.AssertEqual 21, modMathUtils.Multiply(7, 3), "乗算のテスト"
    End Sub
    
    Private Sub Test_StringUtils()
        modTestUtility.AssertEqual "HELLO", modStringUtils.ToUpper("hello"), "大文字変換のテスト"
        modTestUtility.AssertEqual "hello", modStringUtils.ToLower("HELLO"), "小文字変換のテスト"
        modTestUtility.AssertTrue modStringUtils.IsEmpty(""), "空文字チェックのテスト"
    End Sub
    
    Private Sub Test_DateUtils()
        Dim testDate As Date
        testDate = DateSerial(2025, 1, 1)
        
        modTestUtility.AssertEqual 2025, modDateUtils.GetYear(testDate), "年の取得テスト"
        modTestUtility.AssertEqual 1, modDateUtils.GetMonth(testDate), "月の取得テスト"
        modTestUtility.AssertEqual 1, modDateUtils.GetDay(testDate), "日の取得テスト"
    End Sub
    
    ' ======================
    ' データベーステストケース
    ' ======================
    Private Sub Test_DatabaseUtils()
        ' 接続文字列の生成テスト
        Dim connStr As String
        connStr = modDatabaseUtils.BuildConnectionString("Server", "Database", "User", "Pass")
        modTestUtility.AssertTrue Len(connStr) > 0, "接続文字列生成のテスト"
    End Sub
    
    Private Sub Test_ConnectionPool()
        Dim pool As New ConnectionPool
        
        ' プール設定のテスト
        pool.MaxPoolSize = 10
        modTestUtility.AssertEqual 10, pool.MaxPoolSize, "最大プールサイズの設定テスト"
        
        ' 接続管理のテスト
        modTestUtility.AssertEqual 0, pool.ActiveConnections, "初期接続数のテスト"
    End Sub
    
    ' ======================
    ' セキュリティテストケース
    ' ======================
    Private Sub Test_Crypto()
        Dim crypto As New clsCrypto
        
        ' プロバイダーの検証
        modTestUtility.AssertTrue crypto.ValidateProvider(), "ValidateProviderのテスト"
        
        ' 暗号化/復号化のテスト
        Const testString As String = "テスト文字列"
        Const testKey As String = "テストキー"
        
        Dim encrypted As String
        encrypted = crypto.EncryptString(testString, testKey)
        modTestUtility.AssertTrue Len(encrypted) > 0, "暗号化テスト"
        
        Dim decrypted As String
        decrypted = crypto.DecryptString(encrypted, testKey)
        modTestUtility.AssertEqual testString, decrypted, "復号化テスト"
    End Sub
    
    ' ======================
    ' パフォーマンステストケース
    ' ======================
    Private Sub Test_PerformanceMonitor()
        Dim monitor As New clsPerformanceMonitor
        
        monitor.StartMeasurement "TestOperation"
        ' 何らかの処理
        monitor.EndMeasurement "TestOperation"
        
        Dim result As String
        result = monitor.GetMeasurement("TestOperation")
        modTestUtility.AssertTrue Len(result) > 0, "パフォーマンス計測結果の確認"
    End Sub
    
    Private Sub Test_Lock()
        Dim lock As New clsLock
        
        modTestUtility.AssertTrue lock.TryAcquire(), "ロック取得のテスト"
        lock.Release
        modTestUtility.AssertTrue lock.TryAcquire(), "ロック解放後の再取得テスト"
    End Sub
    
    Private Sub Test_CallStack()
        Dim callStack As New clsCallStack
        
        ' Push/Popのテスト
        callStack.Push "Module1", "Proc1"
        callStack.Push "Module2", "Proc2"
        
        modTestUtility.AssertEqual "Module2.Proc2", callStack.Pop(), "Pop()のテスト1"
        modTestUtility.AssertEqual "Module1.Proc1", callStack.Pop(), "Pop()のテスト2"
        
        ' スタックの状態検証
        modTestUtility.AssertTrue callStack.ValidateStackState(), "ValidateStackStateのテスト"
    End Sub
#End If