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

Implements IErrorHandler

' ======================
' シングルトンインスタンス
' ======================
Private Shared instance As InvalidInputErrorHandler

' ======================
' ファクトリメソッド
' ======================
Public Function Create() As InvalidInputErrorHandler
    If instance Is Nothing Then
        Set instance = New InvalidInputErrorHandler
    End If
    Set Create = instance
End Function

' ======================
' IErrorHandler実装
' ======================
Private Function IErrorHandler_HandleError(ByRef errorDetail As typErrorDetail) As Boolean
    Dim proceed As Boolean
    proceed = False  ' 入力エラーは基本的に処理を中断
    
    ' エラー情報をログに記録
    LogError errorDetail
    
    ' ユーザーに通知
    NotifyUser errorDetail, vbExclamation
    
    ' 入力値の自動補正や、デフォルト値の適用など、
    ' 入力エラー時の回復処理をここに実装可能
    
    IErrorHandler_HandleError = proceed
End Function

' ======================
' プライベートメソッド
' ======================
Private Sub LogError(ByRef errorDetail As typErrorDetail)
    ' エラーログの出力
    Dim logger As ILogger
    Set logger = New FileLogger  ' または適切なログ実装を使用
    
    logger.Log MODULE_NAME, _
              "無効な入力: " & errorDetail.Description, _
              errorDetail.ErrorCode
End Sub

Private Sub NotifyUser(ByRef errorDetail As typErrorDetail, ByVal style As VbMsgBoxStyle)
    ' ユーザーへの通知
    Dim notifier As IUserNotifier
    Set notifier = New MsgBoxNotifier  ' または適切な通知実装を使用
    
    notifier.Notify errorDetail, style, "入力エラー"
End Sub

' ======================
' 入力値検証用のヘルパーメソッド
' ======================
Public Function ValidateNumericRange(ByVal value As Variant, _
                                   ByVal minValue As Double, _
                                   ByVal maxValue As Double) As Boolean
    If Not IsNumeric(value) Then
        ValidateNumericRange = False
        Exit Function
    End If
    
    Dim numValue As Double
    numValue = CDbl(value)
    
    ValidateNumericRange = (numValue >= minValue And numValue <= maxValue)
End Function

Public Function ValidateStringLength(ByVal value As String, _
                                   ByVal minLength As Long, _
                                   ByVal maxLength As Long) As Boolean
    Dim length As Long
    length = Len(value)
    
    ValidateStringLength = (length >= minLength And length <= maxLength)
End Function

Public Function ValidateRequired(ByVal value As Variant) As Boolean
    If IsNull(value) Then
        ValidateRequired = False
        Exit Function
    End If
    
    If VarType(value) = vbString Then
        ValidateRequired = (Len(Trim(CStr(value))) > 0)
    Else
        ValidateRequired = True
    End If
End Function