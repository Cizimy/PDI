Option Explicit

' ======================
' テスト関連の定数
' ======================
Private Const MODULE_NAME As String = "modTestUtility"
Private Const TEST_RESULTS_FILE As String = "TestResults.log"

' テスト結果の状態
Public Enum TestResult
    ResultPass = 1
    ResultFail = 2
    ResultSkip = 3
    ResultError = 4
End Enum

' テストケース情報
Private Type TestCase
    Name As String
    Description As String
    Category As String
    Priority As Integer
    Result As TestResult
    ErrorMessage As String
    ExecutionTime As Double
End Type

' ======================
' プライベート変数
' ======================
Private mTestCases As Collection
Private mPerformanceMonitor As clsPerformanceMonitor
Private mCurrentTestCase As TestCase
Private mIsInitialized As Boolean

' ======================
' 初期化処理
' ======================
Public Sub InitializeTestModule()
    If mIsInitialized Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Set mTestCases = New Collection
    Set mPerformanceMonitor = New clsPerformanceMonitor
    mIsInitialized = True
    Exit Sub

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "テストモジュールの初期化中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "InitializeTestModule"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    
    ' 初期化失敗時は例外を再スロー
    Err.Raise errDetail.Code, errDetail.Source, errDetail.Description
End Sub

' ======================
' テスト実行関連
' ======================
Public Sub StartTest(ByVal testName As String, ByVal description As String, _
                    Optional ByVal category As String = "General", _
                    Optional ByVal priority As Integer = 1)
                    
    If Not mIsInitialized Then InitializeTestModule
    On Error GoTo ErrorHandler
    
    ' 新しいテストケースを初期化
    With mCurrentTestCase
        .Name = testName
        .Description = description
        .Category = category
        .Priority = priority
        .Result = ResultSkip
        .ErrorMessage = ""
    End With
    
    ' パフォーマンス計測開始
    mPerformanceMonitor.StartMeasurement testName
    
    ' ログにテスト開始を記録
    LogTestEvent "テスト開始: " & testName & " (" & description & ")"
    Exit Sub

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "テスト開始処理中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "StartTest"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    
    ' エラー発生時は現在のテストのパフォーマンス計測を終了
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement testName
    End If
    ' エラー発生時はテストをエラー状態で終了
    EndTest ResultError, "テスト開始処理中にエラーが発生: " & Err.Description
End Sub

Public Sub EndTest(ByVal result As TestResult, Optional ByVal errorMessage As String = "")
    If Not mIsInitialized Then Exit Sub
    On Error GoTo ErrorHandler
    
    Dim originalResult As TestResult
    originalResult = result
    
    ' パフォーマンス計測終了
    mPerformanceMonitor.EndMeasurement mCurrentTestCase.Name
    
    ' テスト結果を設定
    With mCurrentTestCase
        .Result = result
        .ErrorMessage = errorMessage
        .ExecutionTime = GetTestExecutionTime(.Name)
    End With
    
    ' テストケースをコレクションに追加
    mTestCases.Add mCurrentTestCase, mCurrentTestCase.Name
    
    ' ログにテスト終了を記録
    LogTestEvent "テスト終了: " & mCurrentTestCase.Name & " - " & GetResultText(result)
    If errorMessage <> "" Then
        LogTestEvent "エラー詳細: " & errorMessage
    End If
    Exit Sub

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "テスト終了処理中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "EndTest"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    
    ' エラー発生時は元のテスト結果を保持しつつ、エラーメッセージを追加
    With mCurrentTestCase
        .Result = originalResult
        .ErrorMessage = .ErrorMessage & vbCrLf & "テスト終了処理中にエラーが発生: " & Err.Description
    End With
End Sub

' ======================
' アサーション関数
' ======================
Public Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, _
                      Optional ByVal message As String = "")
    If Not mIsInitialized Then InitializeTestModule
    On Error GoTo ErrorHandler
    
    If expected <> actual Then
        Dim errorMsg As String
        errorMsg = "AssertEqual失敗: " & vbCrLf & _
                  "期待値: " & CStr(expected) & vbCrLf & _
                  "実際値: " & CStr(actual)
        If message <> "" Then
            errorMsg = errorMsg & vbCrLf & "メッセージ: " & message
        End If
        
        EndTest ResultFail, errorMsg
        Exit Sub
    End If
    Exit Sub

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "AssertEqual実行中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "AssertEqual"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    
    ' アサーション実行中のエラーはテストを失敗として扱う
    EndTest ResultFail, "アサーション実行中にエラーが発生: " & Err.Description
End Sub

Public Sub AssertTrue(ByVal condition As Boolean, Optional ByVal message As String = "")
    If Not mIsInitialized Then InitializeTestModule
    On Error GoTo ErrorHandler
    
    If Not condition Then
        Dim errorMsg As String
        errorMsg = "AssertTrue失敗"
        If message <> "" Then
            errorMsg = errorMsg & ": " & message
        End If
        
        EndTest ResultFail, errorMsg
        Exit Sub
    End If
    Exit Sub

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "AssertTrue実行中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "AssertTrue"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    
    ' アサーション実行中のエラーはテストを失敗として扱う
    EndTest ResultFail, "アサーション実行中にエラーが発生: " & Err.Description
End Sub

Public Sub AssertFalse(ByVal condition As Boolean, Optional ByVal message As String = "")
    If Not mIsInitialized Then InitializeTestModule
    On Error GoTo ErrorHandler
    
    If condition Then
        Dim errorMsg As String
        errorMsg = "AssertFalse失敗"
        If message <> "" Then
            errorMsg = errorMsg & ": " & message
        End If
        
        EndTest ResultFail, errorMsg
        Exit Sub
    End If
    Exit Sub

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "AssertFalse実行中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "AssertFalse"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    
    ' アサーション実行中のエラーはテストを失敗として扱う
    EndTest ResultFail, "アサーション実行中にエラーが発生: " & Err.Description
End Sub

' ======================
' テスト結果レポート
' ======================
Public Function GenerateTestReport() As String
    If Not mIsInitialized Then InitializeTestModule
    On Error GoTo ErrorHandler
    
    Dim report As String
    Dim testCase As TestCase
    Dim i As Long
    Dim totalTests As Long
    Dim passedTests As Long
    Dim failedTests As Long
    Dim skippedTests As Long
    Dim errorTests As Long
    
    report = "テスト実行レポート" & vbCrLf & _
            "=================" & vbCrLf & _
            "実行日時: " & Now & vbCrLf & vbCrLf
    
    ' カテゴリ別の結果集計
    Dim categories As Collection
    Set categories = New Collection
    
    For i = 1 To mTestCases.Count
        testCase = mTestCases(i)
        
        ' カテゴリの追加
        On Error Resume Next
        categories.Add testCase.Category, testCase.Category
        On Error GoTo ErrorHandler
        
        ' 全体の集計
        totalTests = totalTests + 1
        Select Case testCase.Result
            Case ResultPass: passedTests = passedTests + 1
            Case ResultFail: failedTests = failedTests + 1
            Case ResultSkip: skippedTests = skippedTests + 1
            Case ResultError: errorTests = errorTests + 1
        End Select
    Next i
    
    ' 概要の追加
    report = report & "概要:" & vbCrLf & _
            "- 総テスト数: " & totalTests & vbCrLf & _
            "- 成功: " & passedTests & vbCrLf & _
            "- 失敗: " & failedTests & vbCrLf & _
            "- スキップ: " & skippedTests & vbCrLf & _
            "- エラー: " & errorTests & vbCrLf & vbCrLf
    
    ' カテゴリ別の詳細
    report = report & "カテゴリ別詳細:" & vbCrLf & _
            "=================" & vbCrLf
    
    Dim category As Variant
    For Each category In categories
        report = report & vbCrLf & "カテゴリ: " & category & vbCrLf
        
        For i = 1 To mTestCases.Count
            testCase = mTestCases(i)
            If testCase.Category = category Then
                report = report & _
                        "  - " & testCase.Name & vbCrLf & _
                        "    結果: " & GetResultText(testCase.Result) & vbCrLf & _
                        "    実行時間: " & Format$(testCase.ExecutionTime, "0.000") & " ms" & vbCrLf
                If testCase.ErrorMessage <> "" Then
                    report = report & "    エラー: " & testCase.ErrorMessage & vbCrLf
                End If
            End If
        Next i
    Next category
    
    GenerateTestReport = report
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "テストレポート生成中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "GenerateTestReport"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    
    ' レポート生成エラー時は基本情報のみ返す
    GenerateTestReport = "テストレポート生成中にエラーが発生しました。" & vbCrLf & _
                        "エラー詳細: " & Err.Description & vbCrLf & _
                        "実行日時: " & Now
End Function

' ======================
' ユーティリティ関数
' ======================
Private Function GetResultText(ByVal result As TestResult) As String
    Select Case result
        Case ResultPass: GetResultText = "成功"
        Case ResultFail: GetResultText = "失敗"
        Case ResultSkip: GetResultText = "スキップ"
        Case ResultError: GetResultText = "エラー"
        Case Else: GetResultText = "不明"
    End Select
End Function

Private Function GetTestExecutionTime(ByVal testName As String) As Double
    On Error GoTo ErrorHandler
    
    Dim perfData As String
    perfData = mPerformanceMonitor.GetMeasurement(testName)
    
    ' 実行時間を抽出（パフォーマンスモニターの出力形式に依存）
    Dim pos As Long
    pos = InStr(perfData, "Elapsed Time: ")
    If pos > 0 Then
        GetTestExecutionTime = Val(Mid$(perfData, pos + 14))
    End If
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "テスト実行時間の取得中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "GetTestExecutionTime"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    
    ' エラー時は0を返す
    GetTestExecutionTime = 0
End Function

Private Sub LogTestEvent(ByVal message As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    
    On Error Resume Next
    
    ' ファイルが開けない場合は作成を試みる
    If Dir(TEST_RESULTS_FILE) = "" Then
        Open TEST_RESULTS_FILE For Output As #fileNum
        Close #fileNum
    End If
    
    Open TEST_RESULTS_FILE For Append As #fileNum
    Print #fileNum, Format$(Now, "yyyy/mm/dd hh:nn:ss") & " - " & message
    Close #fileNum
    
    If Err.Number <> 0 Then
        Debug.Print "ログ出力エラー: " & Err.Description
        Err.Clear
    End If
    
    On Error GoTo 0
End Sub

' ======================
' クリーンアップ
' ======================
Public Sub CleanupTestModule()
    If Not mIsInitialized Then Exit Sub
    
    On Error Resume Next
    Set mTestCases = Nothing
    Set mPerformanceMonitor = Nothing
    mIsInitialized = False
    
    If Err.Number <> 0 Then
        Debug.Print "クリーンアップ中にエラーが発生: " & Err.Description
        Err.Clear
    End If
End Sub