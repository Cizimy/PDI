VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPerformanceMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' メモリ情報構造体
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

' パフォーマンス測定データ型
Private Type PerformanceData
    StartTime As Currency
    EndTime As Currency
    ElapsedTime As Double
    MemoryUsage As Currency
    PeakMemoryUsage As Currency
    Description As String
End Type

' メンバ変数
Private mFrequency As Currency
Private mMeasurements As Collection
Private mIsEnabled As Boolean
Private mCurrentOperation As String

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    modWindowsAPI.QueryPerformanceFrequency mFrequency
    Set mMeasurements = New Collection
    mIsEnabled = True
End Sub

Private Sub Class_Terminate()
    Set mMeasurements = Nothing
End Sub

' ======================
' パブリックメソッド
' ======================
Public Sub StartMeasurement(ByVal description As String)
    If Not mIsEnabled Then Exit Sub
    
    Dim perfData As PerformanceData
    Dim memInfo As PROCESS_MEMORY_COUNTERS
    
    ' 現在の時刻を取得
    modWindowsAPI.QueryPerformanceCounter perfData.StartTime
    
    ' メモリ使用量を取得
    memInfo.cb = LenB(memInfo)
    modWindowsAPI.GetProcessMemoryInfo modWindowsAPI.GetCurrentProcess(), memInfo, LenB(memInfo)
    perfData.MemoryUsage = memInfo.WorkingSetSize
    
    perfData.Description = description
    mCurrentOperation = description
    
    ' コレクションに追加
    mMeasurements.Add perfData, description
End Sub

Public Sub EndMeasurement(ByVal description As String)
    If Not mIsEnabled Then Exit Sub
    
    Dim perfData As PerformanceData
    Dim memInfo As PROCESS_MEMORY_COUNTERS
    Dim i As Long
    
    ' 対応する測定データを検索
    For i = mMeasurements.Count To 1 Step -1
        If TypeName(mMeasurements(i)) = "PerformanceData" Then
            perfData = mMeasurements(i)
            If perfData.Description = description Then
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
                mMeasurements.Add perfData, description
                Exit For
            End If
        End If
    Next i
End Sub

Public Function GetMeasurement(ByVal description As String) As String
    If Not mIsEnabled Then Exit Function
    
    Dim perfData As PerformanceData
    Dim result As String
    
    On Error Resume Next
    perfData = mMeasurements(description)
    If Err.Number = 0 Then
        result = "Performance Report for: " & description & vbCrLf & _
                "Elapsed Time: " & Format$(perfData.ElapsedTime, "0.000") & " ms" & vbCrLf & _
                "Initial Memory: " & Format$(perfData.MemoryUsage / 1024 / 1024, "0.00") & " MB" & vbCrLf & _
                "Peak Memory: " & Format$(perfData.PeakMemoryUsage / 1024 / 1024, "0.00") & " MB"
    End If
    On Error GoTo 0
    
    GetMeasurement = result
End Function

Public Function GetAllMeasurements() As String
    If Not mIsEnabled Then Exit Function
    
    Dim result As String
    Dim perfData As PerformanceData
    Dim i As Long
    
    result = "Performance Monitoring Report" & vbCrLf & _
            "=========================" & vbCrLf
    
    For i = 1 To mMeasurements.Count
        perfData = mMeasurements(i)
        result = result & vbCrLf & GetMeasurement(perfData.Description)
    Next i
    
    GetAllMeasurements = result
End Function

Public Sub ClearMeasurements()
    Set mMeasurements = New Collection
End Sub

' ======================
' プロパティ
' ======================
Public Property Get IsEnabled() As Boolean
    IsEnabled = mIsEnabled
End Property

Public Property Let IsEnabled(ByVal Value As Boolean)
    mIsEnabled = Value
End Property

Public Property Get CurrentOperation() As String
    CurrentOperation = mCurrentOperation
End Property

' ======================
' テストサポート機能（開発環境専用）
' 警告: これらのメソッドは開発時のテスト目的でのみ使用し、
' 本番環境では使用しないでください。
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
        GetMeasurementCount = mMeasurements.Count
    End Function
    
    ''' <summary>
    ''' テスト用の測定データを追加（テスト用）
    ''' </summary>
    Private Sub SimulateMeasurement(ByVal description As String, ByVal elapsedTime As Double)
        Dim perfData As PerformanceData
        perfData.Description = description
        perfData.ElapsedTime = elapsedTime
        perfData.MemoryUsage = 1000000 ' 1MB
        perfData.PeakMemoryUsage = 2000000 ' 2MB
        mMeasurements.Add perfData, description
    End Sub
#End If