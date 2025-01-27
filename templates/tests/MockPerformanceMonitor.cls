VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockPerformanceMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements IPerformanceMonitor

' モック用の内部状態
Private Type MeasurementInfo
    StartTime As Date
    EndTime As Date
    Category As String
End Type

Private mMeasurements() As MeasurementInfo
Private mMeasurementCount As Long
Private mMaxMeasurements As Long

'*******************************************************************************
' 初期化
'*******************************************************************************
Private Sub Class_Initialize()
    mMaxMeasurements = 100
    ReDim mMeasurements(1 To mMaxMeasurements)
    mMeasurementCount = 0
End Sub

'*******************************************************************************
' IPerformanceMonitor インターフェースの実装
'*******************************************************************************
Private Sub IPerformanceMonitor_StartMeasurement(ByVal category As String)
    If mMeasurementCount >= mMaxMeasurements Then Exit Sub
    
    mMeasurementCount = mMeasurementCount + 1
    With mMeasurements(mMeasurementCount)
        .StartTime = Now
        .Category = category
    End With
End Sub

Private Sub IPerformanceMonitor_EndMeasurement(ByVal category As String)
    Dim i As Long
    For i = mMeasurementCount To 1 Step -1
        If mMeasurements(i).Category = category Then
            mMeasurements(i).EndTime = Now
            Exit Sub
        End If
    Next i
End Sub

'*******************************************************************************
' モック用のヘルパーメソッド
'*******************************************************************************
Public Function GetMeasurementCount() As Long
    GetMeasurementCount = mMeasurementCount
End Function

Public Function GetMeasurementDuration(ByVal index As Long) As Double
    If index < 1 Or index > mMeasurementCount Then
        GetMeasurementDuration = 0
        Exit Function
    End If
    
    With mMeasurements(index)
        If .EndTime = #12:00:00 AM# Then
            GetMeasurementDuration = 0
        Else
            GetMeasurementDuration = DateDiff("s", .StartTime, .EndTime)
        End If
    End With
End Function

Public Function GetMeasurementCategory(ByVal index As Long) As String
    If index < 1 Or index > mMeasurementCount Then
        GetMeasurementCategory = ""
        Exit Function
    End If
    
    GetMeasurementCategory = mMeasurements(index).Category
End Function

Public Sub ResetMeasurements()
    mMeasurementCount = 0
End Sub

'*******************************************************************************
' テスト用のヘルパーメソッド
'*******************************************************************************
Public Function HasMeasurement(ByVal category As String) As Boolean
    Dim i As Long
    For i = 1 To mMeasurementCount
        If mMeasurements(i).Category = category Then
            HasMeasurement = True
            Exit Function
        End If
    Next i
    HasMeasurement = False
End Function

Public Function GetCompletedMeasurementCount() As Long
    Dim count As Long
    Dim i As Long
    For i = 1 To mMeasurementCount
        If mMeasurements(i).EndTime <> #12:00:00 AM# Then
            count = count + 1
        End If
    Next i
    GetCompletedMeasurementCount = count
End Function