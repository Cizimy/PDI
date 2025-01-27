VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExecutionContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

' プライベート変数
Private mStartTime As Date
Private mEndTime As Date
Private mChainId As String
Private mErrorInfo As ErrorInfo
Private mCurrentStrategy As String
Private mCurrentPriority As Long
Private mUsingFallback As Boolean
Private mDuration As Long

' プロパティ
Public Property Get StartTime() As Date
    StartTime = mStartTime
End Property

Public Property Let StartTime(ByVal value As Date)
    mStartTime = value
End Property

Public Property Get EndTime() As Date
    EndTime = mEndTime
End Property

Public Property Let EndTime(ByVal value As Date)
    mEndTime = value
End Property

Public Property Get ChainId() As String
    ChainId = mChainId
End Property

Public Property Let ChainId(ByVal value As String)
    mChainId = value
End Property

Public Property Get ErrorInfo() As ErrorInfo
    Set ErrorInfo = mErrorInfo
End Property

Public Property Set ErrorInfo(ByVal value As ErrorInfo)
    Set mErrorInfo = value
End Property

Public Property Get CurrentStrategy() As String
    CurrentStrategy = mCurrentStrategy
End Property

Public Property Let CurrentStrategy(ByVal value As String)
    mCurrentStrategy = value
End Property

Public Property Get CurrentPriority() As Long
    CurrentPriority = mCurrentPriority
End Property

Public Property Let CurrentPriority(ByVal value As Long)
    mCurrentPriority = value
End Property

Public Property Get UsingFallback() As Boolean
    UsingFallback = mUsingFallback
End Property

Public Property Let UsingFallback(ByVal value As Boolean)
    mUsingFallback = value
End Property

Public Property Get Duration() As Long
    Duration = mDuration
End Property

Public Property Let Duration(ByVal value As Long)
    mDuration = value
End Property

' メソッド
Public Sub Initialize(ByVal chainId As String, ByRef errorInfo As ErrorInfo)
    mStartTime = Now
    mChainId = chainId
    Set mErrorInfo = errorInfo
    mUsingFallback = False
End Sub

Public Sub UpdateDuration()
    mEndTime = Now
    mDuration = DateDiff("s", mStartTime, mEndTime)
End Sub

Public Function Clone() As ExecutionContext
    Dim result As New ExecutionContext
    
    With result
        .StartTime = Me.StartTime
        .EndTime = Me.EndTime
        .ChainId = Me.ChainId
        Set .ErrorInfo = Me.ErrorInfo
        .CurrentStrategy = Me.CurrentStrategy
        .CurrentPriority = Me.CurrentPriority
        .UsingFallback = Me.UsingFallback
        .Duration = Me.Duration
    End With
    
    Set Clone = result
End Function

Public Function ToDictionary() As Dictionary
    Dim result As New Dictionary
    
    With result
        .Add "StartTime", mStartTime
        .Add "EndTime", mEndTime
        .Add "ChainId", mChainId
        .Add "ErrorInfo", mErrorInfo
        .Add "CurrentStrategy", mCurrentStrategy
        .Add "CurrentPriority", mCurrentPriority
        .Add "UsingFallback", mUsingFallback
        .Add "Duration", mDuration
    End With
    
    Set ToDictionary = result
End Function