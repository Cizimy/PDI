VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventLogImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements IEventLog

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "EventLogImpl"

' ======================
' メンバ変数
' ======================
Private mIsInitialized As Boolean
Private mLock As clsLock
Private mPerformanceMonitor As clsPerformanceMonitor
Private mLastError As String

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mLock = New clsLock
    Set mPerformanceMonitor = New clsPerformanceMonitor
    mIsInitialized = True
End Sub

Private Sub Class_Terminate()
    Set mLock = Nothing
    Set mPerformanceMonitor = Nothing
    mIsInitialized = False
End Sub

' ======================
' IEventLog実装
' ======================
Private Function IEventLog_WriteToEventLog(ByVal source As String, _
                                         ByVal message As String, _
                                         ByVal eventType As EventLogType) As Boolean
    If Not mIsInitialized Then Exit Function
    
    mPerformanceMonitor.StartMeasurement "WriteToEventLog"
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    ' ソースの存在確認
    If Not IEventLog_EventSourceExists(source) Then
        LogError "イベントソースが存在しません: " & source
        GoTo CleanUp
    End If
    
    ' メッセージの長さチェック
    If Len(message) > 32766 Then
        LogError "メッセージが長すぎます。32,766文字以内にしてください。"
        GoTo CleanUp
    End If
    
    ' Shell経由でイベントログに書き込み
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    
    ' イベントログ書き込み
    shell.LogEvent CInt(eventType), message, source
    
    Set shell = Nothing
    IEventLog_WriteToEventLog = True
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "WriteToEventLog"
    Exit Function

ErrorHandler:
    mLastError = "イベントログへの書き込み中にエラーが発生しました: " & Err.Description
    LogError mLastError
    IEventLog_WriteToEventLog = False
    Resume CleanUp
End Function

Private Function IEventLog_EventSourceExists(ByVal source As String) As Boolean
    If Not mIsInitialized Then Exit Function
    
    mPerformanceMonitor.StartMeasurement "EventSourceExists"
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    ' レジストリキーの存在確認
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    
    Const HKLM As String = "HKLM\SYSTEM\CurrentControlSet\Services\Eventlog\Application\"
    
    ' レジストリキーの存在確認を試行
    Dim regValue
    regValue = shell.RegRead(HKLM & source & "\EventMessageFile")
    
    Set shell = Nothing
    IEventLog_EventSourceExists = True
    
CleanUp:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "EventSourceExists"
    Exit Function

ErrorHandler:
    ' レジストリキーが存在しない場合も含む
    mLastError = "イベントソースの確認中にエラーが発生しました: " & Err.Description
    IEventLog_EventSourceExists = False
    Resume CleanUp
End Function

' ======================
' プライベートメソッド
' ======================
Private Sub LogError(ByVal message As String)
    Dim errInfo As ErrorInfo
    With errInfo
        .Code = ErrUnexpected
        .Description = message
        .Category = ECSystem
        .Source = MODULE_NAME
        .ProcedureName = "Unknown"  ' スタックトレースから取得可能
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errInfo
End Sub

' ======================
' パブリックプロパティ
' ======================
Public Property Get LastError() As String
    LastError = mLastError
End Property

' ======================
' テストサポート機能
' ======================
#If DEBUG Then
    Public Function ValidateState() As Boolean
        ValidateState = mIsInitialized And _
                       Not (mLock Is Nothing) And _
                       Not (mPerformanceMonitor Is Nothing)
    End Function
    
    Public Function GetPerformanceMonitor() As clsPerformanceMonitor
        Set GetPerformanceMonitor = mPerformanceMonitor
    End Function
    
    Public Function TestEventLogAccess(ByVal testSource As String) As Boolean
        ' イベントログへのアクセス権限をテスト
        Dim testMessage As String
        testMessage = "EventLogImpl Test Message @ " & Now
        
        TestEventLogAccess = IEventLog_WriteToEventLog(testSource, _
                                                      testMessage, _
                                                      EVENTLOG_INFORMATION_TYPE)
    End Function
#End If