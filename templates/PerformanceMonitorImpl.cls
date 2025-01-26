VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PerformanceMonitorImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements IPerformanceMonitor

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "PerformanceMonitorImpl"

' ======================
' 型定義
' ======================
Private Type PerformanceData
    StartTime As Currency
    EndTime As Currency
    ElapsedTime As Double
    MemoryUsage As Currency
    PeakMemoryUsage As Currency
    Description As String
End Type

Private Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Currency
    WorkingSetSize As Currency
    QuotaPeakPagedPoolUsage As Currency
    QuotaPagedPoolUsage As Currency
    QuotaPeakNonPagedPoolUsage As Currency
    QuotaNonPagedPoolUsage As Currency
    PagefileUsage As Currency
    PeakPagefileUsage As Currency
End Type

' ======================
' メンバ変数
' ======================
Private mFrequency As Currency
Private mMeasurements As Collection
Private mIsEnabled As Boolean
Private mCurrentOperation As String
Private mLock As clsLock
Private mConfig As IAppConfig

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    modWindowsAPI.QueryPerformanceFrequency mFrequency
    Set mMeasurements = New Collection
    Set mLock = New clsLock
    Set mConfig = modConfig.GetAppConfig()
    mIsEnabled = mConfig.GetSetting("PerformanceMonitorEnabled", True)
End Sub

Private Sub Class_Terminate()
    Set mMeasurements = Nothing
    Set mLock = Nothing
    Set mConfig = Nothing
End Sub

' ======================
' IPerformanceMonitor実装
' ======================
Private Sub IPerformanceMonitor_Start(ByVal measurementName As String)
    If Not mIsEnabled Then Exit Sub
    
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    Dim perfData As PerformanceData
    Dim memInfo As PROCESS_MEMORY_COUNTERS
    
    ' 現在の時刻を取得
    modWindowsAPI.QueryPerformanceCounter perfData.StartTime
    
    ' メモリ使用量を取得
    memInfo.cb = LenB(memInfo)
    modWindowsAPI.GetProcessMemoryInfo modWindowsAPI.GetCurrentProcess(), memInfo, LenB(memInfo)
    perfData.MemoryUsage = memInfo.WorkingSetSize
    
    perfData.Description = measurementName
    mCurrentOperation = measurementName
    
    ' コレクションに追加
    mMeasurements.Add perfData, measurementName
    
    mLock.ReleaseLock
    Exit Sub

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "パフォーマンス測定の開始中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "IPerformanceMonitor_Start"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    mLock.ReleaseLock
End Sub

Private Sub IPerformanceMonitor_Stop(ByVal measurementName As String)
    If Not mIsEnabled Then Exit Sub
    
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    Dim perfData As PerformanceData
    Dim memInfo As PROCESS_MEMORY_COUNTERS
    Dim i As Long
    
    ' 対応する測定データを検索
    For i = mMeasurements.Count To 1 Step -1
        If TypeName(mMeasurements(i)) = "PerformanceData" Then
            perfData = mMeasurements(i)
            If perfData.Description = measurementName Then
                ' 終了時刻を記録
                modWindowsAPI.QueryPerformanceCounter perfData.EndTime
                
                ' メモリ情報を更新
                memInfo.cb = LenB(memInfo)
                modWindowsAPI.GetProcessMemoryInfo modWindowsAPI.GetCurrentProcess(), memInfo, LenB(memInfo)
                perfData.PeakMemoryUsage = memInfo.PeakWorkingSetSize
                
                ' 経過時間を計算（ミリ秒単位）
                perfData.ElapsedTime = (perfData.EndTime - perfData.StartTime) * 1000 / mFrequency
                
                ' コレクションを更新
                mMeasurements.Remove i
                mMeasurements.Add perfData, measurementName
                Exit For
            End If
        End If
    Next i
    
    mLock.ReleaseLock
    Exit Sub

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "パフォーマンス測定の終了中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "IPerformanceMonitor_Stop"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    mLock.ReleaseLock
End Sub

Private Function IPerformanceMonitor_GetMeasurement(ByVal measurementName As String) As Double
    If Not mIsEnabled Then Exit Function
    
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    Dim perfData As PerformanceData
    perfData = mMeasurements(measurementName)
    IPerformanceMonitor_GetMeasurement = perfData.ElapsedTime
    
    mLock.ReleaseLock
    Exit Function

ErrorHandler:
    If Err.Number = 5 Then ' Item not found in collection
        IPerformanceMonitor_GetMeasurement = 0
    Else
        Dim errDetail As ErrorInfo
        With errDetail
            .Code = ErrUnexpected
            .Description = "パフォーマンス測定データの取得中にエラーが発生しました: " & Err.Description
            .Category = ECGeneral
            .Source = MODULE_NAME
            .ProcedureName = "IPerformanceMonitor_GetMeasurement"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        modError.HandleError errDetail
    End If
    mLock.ReleaseLock
End Function

Private Function IPerformanceMonitor_GetAllMeasurements() As Variant
    If Not mIsEnabled Then Exit Function
    
    mLock.AcquireLock
    On Error GoTo ErrorHandler
    
    Dim results() As Double
    ReDim results(0 To mMeasurements.Count - 1)
    
    Dim i As Long
    Dim perfData As PerformanceData
    For i = 1 To mMeasurements.Count
        perfData = mMeasurements(i)
        results(i - 1) = perfData.ElapsedTime
    Next i
    
    IPerformanceMonitor_GetAllMeasurements = results
    
    mLock.ReleaseLock
    Exit Function

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "パフォーマンス測定データの一括取得中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "IPerformanceMonitor_GetAllMeasurements"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    mLock.ReleaseLock
End Function

Private Sub IPerformanceMonitor_Clear()
    mLock.AcquireLock
    Set mMeasurements = New Collection
    mLock.ReleaseLock
End Sub

Private Property Get IPerformanceMonitor_IsEnabled() As Boolean
    IPerformanceMonitor_IsEnabled = mIsEnabled
End Property

Private Sub IPerformanceMonitor_Enable()
    mIsEnabled = True
    mConfig.SetSetting "PerformanceMonitorEnabled", True
End Sub

Private Sub IPerformanceMonitor_Disable()
    mIsEnabled = False
    mConfig.SetSetting "PerformanceMonitorEnabled", False
End Sub

' ======================
' パブリックメソッド
' ======================
Public Function GetDetailedMeasurement(ByVal measurementName As String) As String
    If Not mIsEnabled Then Exit Function
    
    mLock.AcquireLock
    
    Dim perfData As PerformanceData
    Dim result As String
    
    On Error Resume Next
    perfData = mMeasurements(measurementName)
    If Err.Number = 0 Then
        result = "Performance Report for: " & measurementName & vbCrLf & _
                "Elapsed Time: " & Format$(perfData.ElapsedTime, "0.000") & " ms" & vbCrLf & _
                "Initial Memory: " & Format$(perfData.MemoryUsage / 1024 / 1024, "0.00") & " MB" & vbCrLf & _
                "Peak Memory: " & Format$(perfData.PeakMemoryUsage / 1024 / 1024, "0.00") & " MB"
    End If
    On Error GoTo 0
    
    mLock.ReleaseLock
    GetDetailedMeasurement = result
End Function

' ======================
' テストサポート機能（開発環境専用）
' ======================
#If DEBUG Then
    ''' <summary>
    ''' パフォーマンスカウンターの有効性を検証（テスト用）
    ''' </summary>
    Private Function ValidatePerformanceCounter() As Boolean
        Dim counter As Currency
        ValidatePerformanceCounter = (modWindowsAPI.QueryPerformanceCounter(counter) <> 0) And (mFrequency <> 0)
    End Function
    
    ''' <summary>
    ''' 測定データの数を取得（テスト用）
    ''' </summary>
    Private Function GetMeasurementCount() As Long
        mLock.AcquireLock
        GetMeasurementCount = mMeasurements.Count
        mLock.ReleaseLock
    End Function
    
    ''' <summary>
    ''' テスト用の測定データを追加（テスト用）
    ''' </summary>
    Private Sub SimulateMeasurement(ByVal description As String, ByVal elapsedTime As Double)
        mLock.AcquireLock
        Dim perfData As PerformanceData
        perfData.Description = description
        perfData.ElapsedTime = elapsedTime
        perfData.MemoryUsage = 1000000 ' 1MB
        perfData.PeakMemoryUsage = 2000000 ' 2MB
        mMeasurements.Add perfData, description
        mLock.ReleaseLock
    End Sub
    
    ''' <summary>
    ''' 現在の操作名を取得します
    ''' </summary>
    Public Property Get CurrentOperation() As String
        CurrentOperation = mCurrentOperation
    End Property
#End If