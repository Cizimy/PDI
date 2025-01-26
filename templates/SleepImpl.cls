VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SleepImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISleep

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "SleepImpl"

' ======================
' メンバ変数
' ======================
Private mIsInitialized As Boolean
Private mPerformanceCounter As IPerformanceCounter
Private mIsHighPrecision As Boolean

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    ' 高精度タイマーが利用可能か確認
    Set mPerformanceCounter = New PerformanceCounterImpl
    mIsHighPrecision = (TypeOf mPerformanceCounter Is PerformanceCounterImpl)
    mIsInitialized = True
End Sub

Private Sub Class_Terminate()
    Set mPerformanceCounter = Nothing
    mIsInitialized = False
End Sub

' ======================
' ISleep実装
' ======================
Private Sub ISleep_Sleep(ByVal milliseconds As Long)
    If Not mIsInitialized Then Exit Sub
    If milliseconds <= 0 Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    If mIsHighPrecision Then
        HighPrecisionSleep milliseconds
    Else
        StandardSleep milliseconds
    End If
    
    Exit Sub

ErrorHandler:
    LogError "スリープ処理中にエラーが発生しました: " & Err.Description
    StandardSleep milliseconds  ' エラー時は標準のSleepにフォールバック
End Sub

' ======================
' プライベートメソッド
' ======================
Private Sub HighPrecisionSleep(ByVal milliseconds As Long)
    Dim startCount As Currency
    Dim currentCount As Currency
    Dim frequency As Currency
    Dim targetCount As Currency
    
    ' 周波数とカウンターの取得
    If Not mPerformanceCounter.QueryPerformanceFrequency(frequency) Then
        StandardSleep milliseconds
        Exit Sub
    End If
    
    If Not mPerformanceCounter.QueryPerformanceCounter(startCount) Then
        StandardSleep milliseconds
        Exit Sub
    End If
    
    ' 目標カウントの計算
    targetCount = startCount + (frequency * milliseconds) / 1000
    
    ' 高精度ウェイト
    Do
        If Not mPerformanceCounter.QueryPerformanceCounter(currentCount) Then
            StandardSleep milliseconds
            Exit Sub
        End If
    Loop While currentCount < targetCount
End Sub

Private Sub StandardSleep(ByVal milliseconds As Long)
    modWindowsAPI.Sleep milliseconds
End Sub

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
    Public Function ValidateState() As Boolean
        ValidateState = mIsInitialized And Not (mPerformanceCounter Is Nothing)
    End Function
    
    Public Function IsHighPrecision() As Boolean
        IsHighPrecision = mIsHighPrecision
    End Function
    
    Public Function TestSleepAccuracy(ByVal milliseconds As Long) As Double
        Dim startTime As Currency
        Dim endTime As Currency
        
        ' 開始時刻を取得
        If Not mPerformanceCounter.QueryPerformanceCounter(startTime) Then
            TestSleepAccuracy = -1
            Exit Function
        End If
        
        ' スリープ実行
        ISleep_Sleep milliseconds
        
        ' 終了時刻を取得
        If Not mPerformanceCounter.QueryPerformanceCounter(endTime) Then
            TestSleepAccuracy = -1
            Exit Function
        End If
        
        ' 実際の経過時間を計算（ミリ秒）
        Dim frequency As Currency
        If Not mPerformanceCounter.QueryPerformanceFrequency(frequency) Then
            TestSleepAccuracy = -1
            Exit Function
        End If
        
        TestSleepAccuracy = (endTime - startTime) * 1000 / frequency
    End Function
#End If