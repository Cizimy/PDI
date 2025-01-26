VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InvalidInputErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "InvalidInputErrorHandler"

' ======================
' プライベート変数
' ======================
Private mLock As ILock
Private mLogger As ILogger
Private mEmergencyLogger As IEmergencyLogger
Private mUserNotifier As IUserNotifier
Private mValidator As IValidator

Implements IErrorHandler

' =====================
' コンストラクタ
' =====================
Public Function Create( _
    ByVal lock As ILock, _
    ByVal logger As ILogger, _
    ByVal emergencyLogger As IEmergencyLogger, _
    ByVal userNotifier As IUserNotifier, _
    ByVal validator As IValidator _
) As InvalidInputErrorHandler
    
    Dim instance As New InvalidInputErrorHandler
    
    ' 依存性の注入
    With instance
        If lock Is Nothing Then
            Err.Raise 5, "InvalidInputErrorHandler.Create", "ILock is required"
        End If
        Set .Lock = lock
        
        If logger Is Nothing Then
            Err.Raise 5, "InvalidInputErrorHandler.Create", "ILogger is required"
        End If
        Set .Logger = logger
        
        If emergencyLogger Is Nothing Then
            Err.Raise 5, "InvalidInputErrorHandler.Create", "IEmergencyLogger is required"
        End If
        Set .EmergencyLogger = emergencyLogger
        
        If userNotifier Is Nothing Then
            Err.Raise 5, "InvalidInputErrorHandler.Create", "IUserNotifier is required"
        End If
        Set .UserNotifier = userNotifier
        
        If validator Is Nothing Then
            Err.Raise 5, "InvalidInputErrorHandler.Create", "IValidator is required"
        End If
        Set .Validator = validator
    End With
    
    Set Create = instance
End Function

' プロパティ
Public Property Set Lock(ByVal value As ILock)
    Set mLock = value
End Property

Public Property Set Logger(ByVal value As ILogger)
    Set mLogger = value
End Property

Public Property Set EmergencyLogger(ByVal value As IEmergencyLogger)
    Set mEmergencyLogger = value
End Property

Public Property Set UserNotifier(ByVal value As IUserNotifier)
    Set mUserNotifier = value
End Property

Public Property Set Validator(ByVal value As IValidator)
    Set mValidator = value
End Property

' ======================
' IErrorHandler実装
' ======================
Private Function IErrorHandler_HandleError(ByRef errorDetail As ErrorInfo) As Boolean
    Dim proceed As Boolean
    proceed = False  ' デフォルトでは処理を中断
    
    mLock.AcquireLock
    
    ' エラー情報をログに記録
    LogError errorDetail
    
    ' 入力値の自動補正を試みる
    If errorDetail.HasAdditionalInfo("InputValue") And _
       errorDetail.HasAdditionalInfo("InputType") Then
        
        Dim inputValue As Variant
        Dim inputType As String
        inputValue = errorDetail.GetAdditionalInfo("InputValue")
        inputType = errorDetail.GetAdditionalInfo("InputType")
        
        ' 入力値の自動補正を試みる
        Dim correctedValue As Variant
        If TryCorrectInput(inputValue, inputType, correctedValue) Then
            errorDetail.AddAdditionalInfo "CorrectedValue", correctedValue
            errorDetail.RecoveryAttempted = True
            errorDetail.RecoverySuccessful = True
            proceed = True
        End If
    End If
    
    mLock.ReleaseLock
    
    ' ユーザーに通知
    NotifyUser errorDetail, IIf(proceed, vbInformation, vbExclamation)
    
    IErrorHandler_HandleError = proceed
End Function

' ======================
' プライベートメソッド
' ======================
Private Function TryCorrectInput(ByVal value As Variant, ByVal inputType As String, ByRef correctedValue As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    Select Case LCase(inputType)
        Case "number"
            If TryCorrectNumber(value, correctedValue) Then
                TryCorrectInput = True
                Exit Function
            End If
            
        Case "date"
            If TryCorrectDate(value, correctedValue) Then
                TryCorrectInput = True
                Exit Function
            End If
            
        Case "string"
            If TryCorrectString(value, correctedValue) Then
                TryCorrectInput = True
                Exit Function
            End If
    End Select
    
    TryCorrectInput = False
    Exit Function

ErrorHandler:
    TryCorrectInput = False
End Function

Private Function TryCorrectNumber(ByVal value As Variant, ByRef correctedValue As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' 数値への変換を試みる
    If IsNumeric(value) Then
        correctedValue = CDbl(value)
        TryCorrectNumber = True
        Exit Function
    End If
    
    ' カンマや空白を除去して再試行
    Dim cleanValue As String
    cleanValue = Replace(Replace(CStr(value), ",", ""), " ", "")
    
    If IsNumeric(cleanValue) Then
        correctedValue = CDbl(cleanValue)
        TryCorrectNumber = True
        Exit Function
    End If
    
    TryCorrectNumber = False
    Exit Function

ErrorHandler:
    TryCorrectNumber = False
End Function

Private Function TryCorrectDate(ByVal value As Variant, ByRef correctedValue As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' 日付への変換を試みる
    If IsDate(value) Then
        correctedValue = CDate(value)
        TryCorrectDate = True
        Exit Function
    End If
    
    ' 一般的な日付形式を試行
    Dim formats() As String
    formats = Array("yyyy/mm/dd", "dd/mm/yyyy", "mm/dd/yyyy", "yyyy-mm-dd", "dd-mm-yyyy", "mm-dd-yyyy")
    
    Dim i As Long
    For i = LBound(formats) To UBound(formats)
        If mValidator.IsValidDate(CStr(value), formats(i)) Then
            correctedValue = CDate(value)
            TryCorrectDate = True
            Exit Function
        End If
    Next i
    
    TryCorrectDate = False
    Exit Function

ErrorHandler:
    TryCorrectDate = False
End Function

Private Function TryCorrectString(ByVal value As Variant, ByRef correctedValue As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    ' 文字列の基本的なクリーニング
    Dim cleanValue As String
    cleanValue = Trim(CStr(value))
    
    ' 制御文字の除去
    Dim i As Long
    Dim result As String
    For i = 1 To Len(cleanValue)
        Dim ch As String
        ch = Mid(cleanValue, i, 1)
        If Asc(ch) >= 32 Or ch = vbTab Then
            result = result & ch
        End If
    Next i
    
    If Len(result) > 0 Then
        correctedValue = result
        TryCorrectString = True
    Else
        TryCorrectString = False
    End If
    Exit Function

ErrorHandler:
    TryCorrectString = False
End Function

Private Sub LogError(ByRef errorDetail As ErrorInfo)
    On Error GoTo ErrorHandler
    
    ' 通常のログ出力
    mLogger.Log MODULE_NAME, "無効な入力: " & errorDetail.Description, errorDetail.Code
    
    ' 重要なエラーは緊急ログにも記録
    If errorDetail.Severity >= ESError Then
        mEmergencyLogger.LogEmergency MODULE_NAME, errorDetail.ToString
    End If
    
    Exit Sub

ErrorHandler:
    Dim errDetail As ErrorInfo
    Set errDetail = New ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "入力エラーのログ出力中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "LogError"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
        .Severity = ESError
        Set .InnerError = errorDetail.Clone
    End With
    modError.HandleError errDetail
End Sub

Private Sub NotifyUser(ByRef errorDetail As ErrorInfo, ByVal style As VbMsgBoxStyle)
    On Error GoTo ErrorHandler
    
    Dim message As String
    message = "無効な入力: " & errorDetail.Description
    
    If errorDetail.HasAdditionalInfo("CorrectedValue") Then
        message = message & vbCrLf & vbCrLf & _
                 "推奨される値: " & errorDetail.GetAdditionalInfo("CorrectedValue")
    End If
    
    mUserNotifier.Notify errorDetail, style, "入力エラー"
    Exit Sub

ErrorHandler:
    Dim errDetail As ErrorInfo
    Set errDetail = New ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "入力エラーの通知中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "NotifyUser"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
        .Severity = ESError
        Set .InnerError = errorDetail.Clone
    End With
    modError.HandleError errDetail
End Sub