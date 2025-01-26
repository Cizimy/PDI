VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PerformanceCounterImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements IPerformanceCounter

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "PerformanceCounterImpl"

' ======================
' メンバ変数
' ======================
Private mIsInitialized As Boolean
Private mFrequency As Currency
Private mLock As clsLock
Private mIsHighResolutionSupported As Boolean

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mLock = New clsLock
    
    ' 高分解能タイマーのサポートを確認
    mIsHighResolutionSupported = CheckHighResolutionSupport()
    If mIsHighResolutionSupported Then
        ' 周波数の取得
        If Not modWindowsAPI.QueryPerformanceFrequency(mFrequency) Then
            LogError "パフォーマンスカウンターの周波数取得に失敗しました。"
            mIsHighResolutionSupported = False
        End If
    End If
    
    mIsInitialized = True
End Sub

Private Sub Class_Terminate()
    Set mLock = Nothing
    mIsInitialized = False
End Sub

' ======================
' IPerformanceCounter実装
' ======================
Private Function IPerformanceCounter_QueryPerformanceCounter(ByRef performanceCount As Currency) As Boolean
    If Not mIsInitialized Then Exit Function
    If Not mIsHighResolutionSupported Then Exit Function
    
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    IPerformanceCounter_QueryPerformanceCounter = modWindowsAPI.QueryPerformanceCounter(performanceCount)
    
CleanUp:
    mLock.ReleaseLock
    Exit Function

ErrorHandler:
    LogError "パフォーマンスカウンターの値取得中にエラーが発生しました: " & Err.Description
    IPerformanceCounter_QueryPerformanceCounter = False
    Resume CleanUp
End Function

Private Function IPerformanceCounter_QueryPerformanceFrequency(ByRef frequency As Currency) As Boolean
    If Not mIsInitialized Then Exit Function
    If Not mIsHighResolutionSupported Then Exit Function
    
    mLock.AcquireLock
    
    On Error GoTo ErrorHandler
    
    ' キャッシュされた周波数を返す
    frequency = mFrequency
    IPerformanceCounter_QueryPerformanceFrequency = True
    
CleanUp:
    mLock.ReleaseLock
    Exit Function

ErrorHandler:
    LogError "パフォーマンスカウンターの周波数取得中にエラーが発生しました: " & Err.Description
    IPerformanceCounter_QueryPerformanceFrequency = False
    Resume CleanUp
End Function

' ======================
' プライベートメソッド
' ======================
Private Function CheckHighResolutionSupport() As Boolean
    On Error GoTo ErrorHandler
    
    Dim counter As Currency
    Dim frequency As Currency
    
    ' カウンターと周波数の取得を試みる
    CheckHighResolutionSupport = (modWindowsAPI.QueryPerformanceCounter(counter) <> 0) And _
                                (modWindowsAPI.QueryPerformanceFrequency(frequency) <> 0)
    Exit Function

ErrorHandler:
    CheckHighResolutionSupport = False
End Function

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
' パブリックメソッド
' ======================
Public Function GetResolution() As Double
    ' 分解能を秒単位で返す
    If mIsHighResolutionSupported And mFrequency <> 0 Then
        GetResolution = 1 / CDbl(mFrequency)
    Else
        GetResolution = 0
    End If
End Function

' ======================
' テストサポート機能
' ======================
#If DEBUG Then
    Public Function ValidateState() As Boolean
        ValidateState = mIsInitialized And _
                       Not (mLock Is Nothing) And _
                       mIsHighResolutionSupported And _
                       mFrequency > 0
    End Function
    
    Public Function GetFrequency() As Currency
        GetFrequency = mFrequency
    End Function
    
    Public Function IsHighResolutionSupported() As Boolean
        IsHighResolutionSupported = mIsHighResolutionSupported
    End Function
    
    Public Function TestTimerConsistency() As Boolean
        ' タイマーの一貫性をテスト
        Dim counter1 As Currency
        Dim counter2 As Currency
        
        If Not IPerformanceCounter_QueryPerformanceCounter(counter1) Then
            TestTimerConsistency = False
            Exit Function
        End If
        
        If Not IPerformanceCounter_QueryPerformanceCounter(counter2) Then
            TestTimerConsistency = False
            Exit Function
        End If
        
        ' 2回目のカウンターが1回目より大きいことを確認
        TestTimerConsistency = (counter2 >= counter1)
    End Function
#End If