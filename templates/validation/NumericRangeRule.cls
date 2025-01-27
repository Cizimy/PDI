VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NumericRangeRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("数値の範囲のバリデーションルール")

Private WithEvents BaseRule As ValidationRuleBase
Private mFieldName As String           ' フィールド名
Private mMinValue As Double           ' 最小値
Private mMaxValue As Double           ' 最大値
Private mIncludeMin As Boolean        ' 最小値を含むかどうか
Private mIncludeMax As Boolean        ' 最大値を含むかどうか
Private mAllowNull As Boolean         ' Null値を許可するかどうか
Private mPrecision As Long            ' 小数点以下の桁数

Private Sub Class_Initialize()
    Set BaseRule = New ValidationRuleBase
    BaseRule.Name = "NumericRangeRule"
    BaseRule.Description = "数値が指定された範囲内であることを検証します"
    mMinValue = 0
    mMaxValue = 0
    mIncludeMin = True
    mIncludeMax = True
    mAllowNull = False
    mPrecision = -1 ' デフォルトは制限なし
End Sub

'@Description("フィールド名を取得または設定します")
Public Property Get FieldName() As String
    FieldName = mFieldName
End Property

Public Property Let FieldName(ByVal value As String)
    mFieldName = value
    UpdateErrorMessage
End Property

'@Description("最小値を取得または設定します")
Public Property Get MinValue() As Double
    MinValue = mMinValue
End Property

Public Property Let MinValue(ByVal value As Double)
    mMinValue = value
    UpdateErrorMessage
End Property

'@Description("最大値を取得または設定します")
Public Property Get MaxValue() As Double
    MaxValue = mMaxValue
End Property

Public Property Let MaxValue(ByVal value As Double)
    If value < mMinValue Then
        Err.Raise 5, "NumericRangeRule", "最大値は最小値以上である必要があります"
    End If
    mMaxValue = value
    UpdateErrorMessage
End Property

'@Description("最小値を含むかどうかを取得または設定します")
Public Property Get IncludeMin() As Boolean
    IncludeMin = mIncludeMin
End Property

Public Property Let IncludeMin(ByVal value As Boolean)
    mIncludeMin = value
    UpdateErrorMessage
End Property

'@Description("最大値を含むかどうかを取得または設定します")
Public Property Get IncludeMax() As Boolean
    IncludeMax = mIncludeMax
End Property

Public Property Let IncludeMax(ByVal value As Boolean)
    mIncludeMax = value
    UpdateErrorMessage
End Property

'@Description("Null値を許可するかどうかを取得または設定します")
Public Property Get AllowNull() As Boolean
    AllowNull = mAllowNull
End Property

Public Property Let AllowNull(ByVal value As Boolean)
    mAllowNull = value
End Property

'@Description("小数点以下の桁数を取得または設定します")
Public Property Get Precision() As Long
    Precision = mPrecision
End Property

Public Property Let Precision(ByVal value As Long)
    If value < -1 Then
        Err.Raise 5, "NumericRangeRule", "精度は-1以上である必要があります"
    End If
    mPrecision = value
End Property

'@Description("バリデーションを実行します")
Public Function Validate(ByVal value As Variant) As Boolean
    Validate = BaseRule.Validate(value)
End Function

'@Description("具体的なバリデーション処理を実行します")
Private Function ValidateCore(ByVal value As Variant) As Boolean
    ' Null値のチェック
    If IsNull(value) Then
        If mAllowNull Then
            ValidateCore = True
        Else
            BaseRule.ErrorMessage = GetErrorMessage("NULL値は許可されていません")
            ValidateCore = False
        End If
        Exit Function
    End If
    
    ' 数値型のチェック
    If Not IsNumeric(value) Then
        BaseRule.ErrorMessage = GetErrorMessage("数値型である必要があります")
        ValidateCore = False
        Exit Function
    End If
    
    Dim numValue As Double
    numValue = CDbl(value)
    
    ' 精度のチェック
    If mPrecision >= 0 Then
        Dim decimalPlaces As Long
        decimalPlaces = GetDecimalPlaces(numValue)
        If decimalPlaces > mPrecision Then
            BaseRule.ErrorMessage = GetErrorMessage("小数点以下" & mPrecision & "桁までしか許可されていません")
            ValidateCore = False
            Exit Function
        End If
    End If
    
    ' 範囲チェック
    If mMinValue <> 0 Or mMaxValue <> 0 Then
        ' 最小値チェック
        If mIncludeMin Then
            If numValue < mMinValue Then
                BaseRule.ErrorMessage = GetErrorMessage(mMinValue & "以上である必要があります")
                ValidateCore = False
                Exit Function
            End If
        Else
            If numValue <= mMinValue Then
                BaseRule.ErrorMessage = GetErrorMessage(mMinValue & "より大きい必要があります")
                ValidateCore = False
                Exit Function
            End If
        End If
        
        ' 最大値チェック
        If mMaxValue > 0 Then
            If mIncludeMax Then
                If numValue > mMaxValue Then
                    BaseRule.ErrorMessage = GetErrorMessage(mMaxValue & "以下である必要があります")
                    ValidateCore = False
                    Exit Function
                End If
            Else
                If numValue >= mMaxValue Then
                    BaseRule.ErrorMessage = GetErrorMessage(mMaxValue & "未満である必要があります")
                    ValidateCore = False
                    Exit Function
                End If
            End If
        End If
    End If
    
    ValidateCore = True
End Function

'@Description("小数点以下の桁数を取得します")
Private Function GetDecimalPlaces(ByVal value As Double) As Long
    Dim strValue As String
    strValue = CStr(value)
    
    Dim decimalPos As Long
    decimalPos = InStr(strValue, ".")
    
    If decimalPos = 0 Then
        GetDecimalPlaces = 0
    Else
        GetDecimalPlaces = Len(strValue) - decimalPos
    End If
End Function

'@Description("エラーメッセージを更新します")
Private Sub UpdateErrorMessage()
    Dim msg As String
    
    If Len(mFieldName) > 0 Then
        msg = mFieldName & "は"
    End If
    
    If mMinValue <> 0 Or mMaxValue <> 0 Then
        If mMinValue <> 0 And mMaxValue <> 0 Then
            msg = msg & mMinValue & (IIf(mIncludeMin, "以上", "より大きく")) & _
                  "、" & mMaxValue & (IIf(mIncludeMax, "以下", "未満")) & "である必要があります"
        ElseIf mMinValue <> 0 Then
            msg = msg & mMinValue & (IIf(mIncludeMin, "以上", "より大きい")) & "必要があります"
        ElseIf mMaxValue <> 0 Then
            msg = msg & mMaxValue & (IIf(mIncludeMax, "以下", "未満")) & "である必要があります"
        End If
    End If
    
    If mPrecision >= 0 Then
        If Len(msg) > 0 Then msg = msg & "、"
        msg = msg & "小数点以下" & mPrecision & "桁までしか許可されていません"
    End If
    
    BaseRule.ErrorMessage = msg
End Function

'@Description("エラーメッセージを生成します")
Private Function GetErrorMessage(ByVal detail As String) As String
    If Len(mFieldName) > 0 Then
        GetErrorMessage = mFieldName & ": " & detail
    Else
        GetErrorMessage = detail
    End If
End Function

' BaseRuleのイベントハンドラ
Private Sub BaseRule_RuleExecuted(ByVal ruleName As String, ByVal success As Boolean, ByVal executionTime As Double)
    ' 必要に応じてイベントを処理
End Sub

' ValidationRuleBaseのプロパティとメソッドの委譲
Public Property Get Name() As String
    Name = BaseRule.Name
End Property

Public Property Let Name(ByVal value As String)
    BaseRule.Name = value
End Property

Public Property Get Description() As String
    Description = BaseRule.Description
End Property

Public Property Let Description(ByVal value As String)
    BaseRule.Description = value
End Property

Public Property Get ErrorMessage() As String
    ErrorMessage = BaseRule.ErrorMessage
End Property

Public Property Let ErrorMessage(ByVal value As String)
    BaseRule.ErrorMessage = value
End Property

Public Property Get IsEnabled() As Boolean
    IsEnabled = BaseRule.IsEnabled
End Property

Public Property Let IsEnabled(ByVal value As Boolean)
    BaseRule.IsEnabled = value
End Property

Public Property Get Priority() As ValidationPriority
    Priority = BaseRule.Priority
End Property

Public Property Let Priority(ByVal value As ValidationPriority)
    BaseRule.Priority = value
End Property

Public Property Get Dependencies() As Collection
    Set Dependencies = BaseRule.Dependencies
End Property

Public Sub AddDependency(ByVal rule As IValidationRule)
    BaseRule.AddDependency rule
End Sub

Public Property Get ValidationContext() As IValidationContext
    Set ValidationContext = BaseRule.ValidationContext
End Property

Public Property Set ValidationContext(ByVal value As IValidationContext)
    Set BaseRule.ValidationContext = value
End Property

Public Property Get ExecutionTime() As Double
    ExecutionTime = BaseRule.ExecutionTime
End Property

Public Property Get LastExecutionResult() As Boolean
    LastExecutionResult = BaseRule.LastExecutionResult
End Property

Public Property Get ErrorDetails() As Collection
    Set ErrorDetails = BaseRule.ErrorDetails
End Property

Public Function ToJSON() As String
    ToJSON = BaseRule.ToJSON
End Function
