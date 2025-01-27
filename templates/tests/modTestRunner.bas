Attribute VB_Name = "modTestRunner"
Option Explicit

'@Description("テストの実行を管理するモジュール")

' テスト結果を格納する構造体
Private Type TestResult
    TestName As String
    ClassName As String
    Success As Boolean
    ErrorMessage As String
    ExecutionTime As Double
End Type

' グローバル変数
Private mTestResults As Collection
Private mTotalTests As Long
Private mPassedTests As Long
Private mFailedTests As Long
Private mStartTime As Date
Private mEndTime As Date

'@Description("すべてのテストを実行する")
Public Sub RunAllTests()
    ' テスト結果の初期化
    Set mTestResults = New Collection
    mTotalTests = 0
    mPassedTests = 0
    mFailedTests = 0
    mStartTime = Now
    
    On Error GoTo ErrorHandler
    
    ' テストの実行
    Debug.Print "テストの実行を開始します..."
    Debug.Print String(50, "-")
    
    ' 各テストクラスの実行
    RunConfigTests
    RunLoggerTests
    RunFileOperationsTests
    RunValidatorTests
    RunDatabaseTests
    RunSecurityTests
    
    ' 結果の出力
    mEndTime = Now
    OutputTestResults
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "テスト実行中にエラーが発生しました: " & Err.Description
    Resume Next
End Sub

'@Description("設定関連のテストを実行")
Private Sub RunConfigTests()
    Dim testConfig As New TestConfigImpl
    
    ' テストメソッドの実行
    RunTestMethod testConfig, "TestGetSetSetting"
    RunTestMethod testConfig, "TestLoadSettings"
    RunTestMethod testConfig, "TestEncryptedSetting"
    RunTestMethod testConfig, "TestEnvironmentSpecificSettings"
    RunTestMethod testConfig, "TestSettingsValidation"
    RunTestMethod testConfig, "TestPerformanceMetrics"
    RunTestMethod testConfig, "TestBackupAndRestore"
    RunTestMethod testConfig, "TestSettingHistory"
    RunTestMethod testConfig, "TestInvalidEncryptionKey"
    RunTestMethod testConfig, "TestFileAccessError"
    
    Set testConfig = Nothing
End Sub

'@Description("ロガー関連のテストを実行")
Private Sub RunLoggerTests()
    ' 既存のロガーテスト
    ' ...（既存のコード）
End Sub

'@Description("ファイル操作関連のテストを実行")
Private Sub RunFileOperationsTests()
    ' 既存のファイル操作テスト
    ' ...（既存のコード）
End Sub

'@Description("バリデーション関連のテストを実行")
Private Sub RunValidatorTests()
    ' 既存のバリデーションテスト
    ' ...（既存のコード）
End Sub

'@Description("データベース関連のテストを実行")
Private Sub RunDatabaseTests()
    ' 既存のデータベーステスト
    ' ...（既存のコード）
End Sub

'@Description("セキュリティ関連のテストを実行")
Private Sub RunSecurityTests()
    ' 既存のセキュリティテスト
    ' ...（既存のコード）
End Sub

'@Description("テストメソッドを実行")
Private Sub RunTestMethod(ByVal testClass As Object, ByVal methodName As String)
    Dim result As TestResult
    result.TestName = methodName
    result.ClassName = TypeName(testClass)
    
    On Error Resume Next
    
    Dim startTime As Date
    startTime = Now
    
    ' テストの初期化
    CallByName testClass, "TestInitialize", VbMethod
    
    ' テストメソッドの実行
    CallByName testClass, methodName, VbMethod
    
    ' エラーチェック
    If Err.Number = 0 Then
        result.Success = True
    Else
        result.Success = False
        result.ErrorMessage = Err.Description
    End If
    
    ' テストのクリーンアップ
    CallByName testClass, "TestCleanup", VbMethod
    
    result.ExecutionTime = DateDiff("s", startTime, Now)
    
    ' 結果の集計
    mTestResults.Add result
    mTotalTests = mTotalTests + 1
    If result.Success Then
        mPassedTests = mPassedTests + 1
    Else
        mFailedTests = mFailedTests + 1
    End If
    
    ' 進捗表示
    Debug.Print result.ClassName & "." & result.TestName & ": " & _
                IIf(result.Success, "成功", "失敗 - " & result.ErrorMessage)
    
    On Error GoTo 0
End Sub

'@Description("テスト結果を出力")
Private Sub OutputTestResults()
    Debug.Print String(50, "-")
    Debug.Print "テスト実行結果"
    Debug.Print String(50, "-")
    Debug.Print "合計テスト数: " & mTotalTests
    Debug.Print "成功: " & mPassedTests
    Debug.Print "失敗: " & mFailedTests
    Debug.Print "実行時間: " & Format$(DateDiff("s", mStartTime, mEndTime), "#,##0") & " 秒"
    Debug.Print String(50, "-")
    
    ' 失敗したテストの詳細を出力
    If mFailedTests > 0 Then
        Debug.Print
        Debug.Print "失敗したテストの詳細:"
        Debug.Print String(50, "-")
        
        Dim result As TestResult
        Dim i As Long
        For i = 1 To mTestResults.Count
            result = mTestResults(i)
            If Not result.Success Then
                Debug.Print result.ClassName & "." & result.TestName
                Debug.Print "エラー: " & result.ErrorMessage
                Debug.Print String(50, "-")
            End If
        Next i
    End If
End Sub