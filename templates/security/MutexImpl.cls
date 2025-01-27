VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MutexImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements IMutex

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "MutexImpl"

' ======================
' プライベート変数
' ======================
Private mMutexHandle As LongPtr
Private mIsInitialized As Boolean
Private mPerformanceMonitor As clsPerformanceMonitor

' ======================
' Win32 APIの再定義
' ======================
Private Const WAIT_OBJECT_0 As Long = 0
Private Const WAIT_FAILED As Long = &HFFFFFFFF
Private Const WAIT_ABANDONED As Long = &H80
Private Const ERROR_ALREADY_EXISTS As Long = 183&

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    mMutexHandle = 0
    Set mPerformanceMonitor = New clsPerformanceMonitor
    mIsInitialized = True
End Sub

Private Sub Class_Terminate()
    If mMutexHandle <> 0 Then
        modWindowsAPI.CloseHandle mMutexHandle
        mMutexHandle = 0
    End If
    Set mPerformanceMonitor = Nothing
    mIsInitialized = False
End Sub

' ======================
' IMutex実装
' ======================
Private Function IMutex_CreateMutex(ByVal initialOwner As Boolean, ByVal name As String) As Boolean
    If Not mIsInitialized Then Exit Function
    
    mPerformanceMonitor.StartMeasurement "CreateMutex"
    On Error GoTo ErrorHandler
    
    ' 既存のハンドルをクリーンアップ
    If mMutexHandle <> 0 Then
        modWindowsAPI.CloseHandle mMutexHandle
        mMutexHandle = 0
    End If
    
    ' ミューテックスの作成
    mMutexHandle = modWindowsAPI.CreateMutex(0&, IIf(initialOwner, 1&, 0&), name)
    
    If mMutexHandle = 0 Then
        LogError "ミューテックスの作成に失敗しました。"
        IMutex_CreateMutex = False
    Else
        IMutex_CreateMutex = True
    End If
    
    mPerformanceMonitor.EndMeasurement "CreateMutex"
    Exit Function

ErrorHandler:
    LogError "ミューテックスの作成中にエラーが発生しました: " & Err.Description
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "CreateMutex"
    End If
    IMutex_CreateMutex = False
End Function

Private Function IMutex_ReleaseMutex() As Boolean
    If Not mIsInitialized Then Exit Function
    If mMutexHandle = 0 Then Exit Function
    
    mPerformanceMonitor.StartMeasurement "ReleaseMutex"
    On Error GoTo ErrorHandler
    
    IMutex_ReleaseMutex = (modWindowsAPI.ReleaseMutex(mMutexHandle) <> 0)
    
    mPerformanceMonitor.EndMeasurement "ReleaseMutex"
    Exit Function

ErrorHandler:
    LogError "ミューテックスの解放中にエラーが発生しました: " & Err.Description
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "ReleaseMutex"
    End If
    IMutex_ReleaseMutex = False
End Function

Private Function IMutex_WaitForSingleObject(ByVal timeoutMilliseconds As Long) As Boolean
    If Not mIsInitialized Then Exit Function
    If mMutexHandle = 0 Then Exit Function
    
    mPerformanceMonitor.StartMeasurement "WaitForSingleObject"
    On Error GoTo ErrorHandler
    
    Dim result As Long
    result = modWindowsAPI.WaitForSingleObject(mMutexHandle, timeoutMilliseconds)
    
    Select Case result
        Case WAIT_OBJECT_0, WAIT_ABANDONED
            IMutex_WaitForSingleObject = True
        Case Else
            IMutex_WaitForSingleObject = False
    End Select
    
    mPerformanceMonitor.EndMeasurement "WaitForSingleObject"
    Exit Function

ErrorHandler:
    LogError "ミューテックスの待機中にエラーが発生しました: " & Err.Description
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "WaitForSingleObject"
    End If
    IMutex_WaitForSingleObject = False
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
' テストサポート機能
' ======================
#If DEBUG Then
    Public Function GetMutexHandle() As LongPtr
        GetMutexHandle = mMutexHandle
    End Function
    
    Public Function IsValid() As Boolean
        IsValid = (mMutexHandle <> 0)
    End Function
    
    Public Sub ForceRelease()
        If mMutexHandle <> 0 Then
            modWindowsAPI.CloseHandle mMutexHandle
            mMutexHandle = 0
        End If
    End Sub
    
    Public Function GetPerformanceMonitor() As clsPerformanceMonitor
        Set GetPerformanceMonitor = mPerformanceMonitor
    End Function
#End If