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

'@Description("パフォーマンス監視機能を提供するレガシークラス（新規実装はIPerformanceMonitorを使用してください）")

' ======================
' メンバ変数
' ======================
Private mImpl As IPerformanceMonitor

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mImpl = New PerformanceMonitorImpl
End Sub

Private Sub Class_Terminate()
    Set mImpl = Nothing
End Sub

' ======================
' パブリックメソッド
' （レガシーサポート）
' ======================
Public Sub StartMeasurement(ByVal description As String)
    mImpl.Start description
End Sub

Public Sub EndMeasurement(ByVal description As String)
    mImpl.Stop description
End Sub

Public Function GetMeasurement(ByVal description As String) As String
    GetMeasurement = PerformanceMonitorImpl(mImpl).GetDetailedMeasurement(description)
End Function

Public Function GetAllMeasurements() As String
    Dim measurements As Variant
    measurements = mImpl.GetAllMeasurements()
    
    Dim result As String
    result = "Performance Monitoring Report" & vbCrLf & _
            "=========================" & vbCrLf
    
    Dim i As Long
    For i = LBound(measurements) To UBound(measurements)
        result = result & vbCrLf & "Measurement " & (i + 1) & ": " & measurements(i) & " ms"
    Next i
    
    GetAllMeasurements = result
End Function

Public Sub ClearMeasurements()
    mImpl.Clear
End Sub

' ======================
' プロパティ
' ======================
Public Property Get IsEnabled() As Boolean
    IsEnabled = mImpl.IsEnabled
End Property

Public Property Let IsEnabled(ByVal Value As Boolean)
    If Value Then
        mImpl.Enable
    Else
        mImpl.Disable
    End If
End Property

Public Property Get CurrentOperation() As String
    CurrentOperation = PerformanceMonitorImpl(mImpl).CurrentOperation
End Property