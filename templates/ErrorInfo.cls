VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' プロパティ
' ======================
Private mCode As Long
Private mDescription As String
Private mCategory As ErrorCodeCategory
Private mSource As String
Private mProcedureName As String
Private mStackTrace As String
Private mOccurredAt As Date
Private mSeverity As ErrorSeverity
Private mAdditionalInfo As Collection
Private mInnerError As ErrorInfo
Private mRecoveryAttempted As Boolean
Private mRecoverySuccessful As Boolean

' ======================
' 列挙型
' ======================
Public Enum ErrorSeverity
    ESInfo = 1
    ESWarning = 2
    ESError = 3
    ESCritical = 4
End Enum

' ======================
' プロパティ
' ======================
Public Property Get Code() As Long
    Code = mCode
End Property
Public Property Let Code(ByVal value As Long)
    mCode = value
End Property

Public Property Get Description() As String
    Description = mDescription
End Property
Public Property Let Description(ByVal value As String)
    mDescription = value
End Property

Public Property Get Category() As ErrorCodeCategory
    Category = mCategory
End Property
Public Property Let Category(ByVal value As ErrorCodeCategory)
    mCategory = value
End Property

Public Property Get Source() As String
    Source = mSource
End Property
Public Property Let Source(ByVal value As String)
    mSource = value
End Property

Public Property Get ProcedureName() As String
    ProcedureName = mProcedureName
End Property
Public Property Let ProcedureName(ByVal value As String)
    mProcedureName = value
End Property

Public Property Get StackTrace() As String
    StackTrace = mStackTrace
End Property
Public Property Let StackTrace(ByVal value As String)
    mStackTrace = value
End Property

Public Property Get OccurredAt() As Date
    OccurredAt = mOccurredAt
End Property
Public Property Let OccurredAt(ByVal value As Date)
    mOccurredAt = value
End Property

Public Property Get Severity() As ErrorSeverity
    Severity = mSeverity
End Property
Public Property Let Severity(ByVal value As ErrorSeverity)
    mSeverity = value
End Property

Public Property Get InnerError() As ErrorInfo
    Set InnerError = mInnerError
End Property
Public Property Set InnerError(ByVal value As ErrorInfo)
    Set mInnerError = value
End Property

Public Property Get RecoveryAttempted() As Boolean
    RecoveryAttempted = mRecoveryAttempted
End Property
Public Property Let RecoveryAttempted(ByVal value As Boolean)
    mRecoveryAttempted = value
End Property

Public Property Get RecoverySuccessful() As Boolean
    RecoverySuccessful = mRecoverySuccessful
End Property
Public Property Let RecoverySuccessful(ByVal value As Boolean)
    mRecoverySuccessful = value
End Property

' ======================
' 初期化
' ======================
Private Sub Class_Initialize()
    Set mAdditionalInfo = New Collection
End Sub

' ======================
' パブリックメソッド
' ======================
Public Sub AddAdditionalInfo(ByVal key As String, ByVal value As Variant)
    On Error Resume Next
    mAdditionalInfo.Remove key
    On Error GoTo 0
    mAdditionalInfo.Add value, key
End Sub

Public Function GetAdditionalInfo(ByVal key As String) As Variant
    On Error Resume Next
    GetAdditionalInfo = mAdditionalInfo(key)
    On Error GoTo 0
End Function

Public Function HasAdditionalInfo(ByVal key As String) As Boolean
    On Error Resume Next
    Dim temp As Variant
    temp = mAdditionalInfo(key)
    HasAdditionalInfo = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function Clone() As ErrorInfo
    Dim result As New ErrorInfo
    With result
        .Code = mCode
        .Description = mDescription
        .Category = mCategory
        .Source = mSource
        .ProcedureName = mProcedureName
        .StackTrace = mStackTrace
        .OccurredAt = mOccurredAt
        .Severity = mSeverity
        .RecoveryAttempted = mRecoveryAttempted
        .RecoverySuccessful = mRecoverySuccessful
        If Not mInnerError Is Nothing Then
            Set .InnerError = mInnerError.Clone
        End If
        
        ' 追加情報のコピー
        Dim key As Variant
        For Each key In mAdditionalInfo
            .AddAdditionalInfo CStr(key), mAdditionalInfo(key)
        Next key
    End With
    Set Clone = result
End Function

Public Function ToString() As String
    Dim result As String
    result = "Error Details:" & vbCrLf & _
            "Code: " & mCode & vbCrLf & _
            "Description: " & mDescription & vbCrLf & _
            "Category: " & mCategory & vbCrLf & _
            "Source: " & mSource & vbCrLf & _
            "Procedure: " & mProcedureName & vbCrLf & _
            "Occurred At: " & Format(mOccurredAt, "yyyy/mm/dd hh:nn:ss") & vbCrLf & _
            "Severity: " & mSeverity & vbCrLf & _
            "Recovery Attempted: " & mRecoveryAttempted & vbCrLf & _
            "Recovery Successful: " & mRecoverySuccessful & vbCrLf & _
            "Stack Trace: " & vbCrLf & mStackTrace
    
    If Not mInnerError Is Nothing Then
        result = result & vbCrLf & vbCrLf & "Inner Error:" & vbCrLf & _
                mInnerError.ToString
    End If
    
    ToString = result
End Function