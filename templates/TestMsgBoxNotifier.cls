VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestMsgBoxNotifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "TestMsgBoxNotifier"

' テスト対象
Private mNotifier As MsgBoxNotifier
Private mMockLock As MockLock
Private mMockMonitor As MockPerformanceMonitor
Private mMockConfig As MockAppConfig
Private mMockFormatter As MockMessageFormatter

' ======================
' セットアップ/クリーンアップ
' ======================
Private Sub Class_Initialize()
    ' モックオブジェクトの初期化
    Set mMockLock = New MockLock
    Set mMockMonitor = New MockPerformanceMonitor
    Set mMockConfig = New MockAppConfig
    Set mMockFormatter = New MockMessageFormatter
    
    ' テスト対象の初期化
    Set mNotifier = New MsgBoxNotifier
    mNotifier.Initialize mMockLock, mMockMonitor, mMockConfig, mMockFormatter
End Sub

Private Sub Class_Terminate()
    Set mNotifier = Nothing
    Set mMockLock = Nothing
    Set mMockMonitor = Nothing
    Set mMockConfig = Nothing
    Set mMockFormatter = Nothing
End Sub

' ======================
' テストケース
' ======================

''' <summary>
''' 初期化テスト - 正常系
''' </summary>
Public Sub TestInitialize()
    On Error GoTo TestFail
    
    ' 新しいインスタンスで初期化テスト
    Dim notifier As New MsgBoxNotifier
    notifier.Initialize mMockLock, mMockMonitor, mMockConfig, mMockFormatter
    
    Debug.Print "TestInitialize: OK"
    Exit Sub
    
TestFail:
    Debug.Print "TestInitialize: Failed - " & Err.Description
End Sub

''' <summary>
''' 初期化テスト - 異常系（Nothingを渡す）
''' </summary>
Public Sub TestInitializeWithNothing()
    On Error Resume Next
    
    Dim notifier As New MsgBoxNotifier
    notifier.Initialize Nothing, Nothing, Nothing, Nothing
    
    ' エラーが発生することを確認
    If Err.Number = 5 Then
        Debug.Print "TestInitializeWithNothing: OK"
    Else
        Debug.Print "TestInitializeWithNothing: Failed - Expected error 5, got " & Err.Number
    End If
    
    On Error GoTo 0
End Sub

''' <summary>
''' 通知スタイル設定テスト
''' </summary>
Public Sub TestNotificationStyle()
    On Error GoTo TestFail
    
    ' スタイル設定
    Dim testStyle As VbMsgBoxStyle
    testStyle = vbOKOnly Or vbInformation
    mNotifier.NotificationStyle = testStyle
    
    ' 設定値の取得と検証
    If mNotifier.NotificationStyle = testStyle Then
        Debug.Print "TestNotificationStyle: OK"
    Else
        Debug.Print "TestNotificationStyle: Failed - Style mismatch"
    End If
    Exit Sub
    
TestFail:
    Debug.Print "TestNotificationStyle: Failed - " & Err.Description
End Sub

''' <summary>
''' メッセージフォーマット処理テスト
''' </summary>
Public Sub TestMessageFormatting()
    On Error GoTo TestFail
    
    ' テスト用エラー情報
    Dim testError As ErrorInfo
    With testError
        .Code = ErrUnexpected
        .Description = "テストエラー"
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "TestMessageFormatting"
    End With
    
    ' モックフォーマッタの戻り値を設定
    mMockFormatter.SetReturnValue "テストフォーマット済メッセージ"
    
    ' 通知実行
    mNotifier.Notify testError
    
    ' フォーマッタが呼び出されたことを確認
    If mMockFormatter.WasCalled Then
        Debug.Print "TestMessageFormatting: OK"
    Else
        Debug.Print "TestMessageFormatting: Failed - Formatter not called"
    End If
    Exit Sub
    
TestFail:
    Debug.Print "TestMessageFormatting: Failed - " & Err.Description
End Sub

''' <summary>
''' ロック処理テスト
''' </summary>
Public Sub TestLocking()
    On Error GoTo TestFail
    
    ' テスト用エラー情報
    Dim testError As ErrorInfo
    With testError
        .Code = ErrUnexpected
        .Description = "テストエラー"
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "TestLocking"
    End With
    
    ' 通知実行
    mNotifier.Notify testError
    
    ' ロックの取得と解放が行われたことを確認
    If mMockLock.AcquireCalled And mMockLock.ReleaseCalled Then
        Debug.Print "TestLocking: OK"
    Else
        Debug.Print "TestLocking: Failed - Lock operations not called"
    End If
    Exit Sub
    
TestFail:
    Debug.Print "TestLocking: Failed - " & Err.Description
End Sub

' ======================
' テスト実行
' ======================
Public Sub RunAllTests()
    TestInitialize
    TestInitializeWithNothing
    TestNotificationStyle
    TestMessageFormatting
    TestLocking
End Sub