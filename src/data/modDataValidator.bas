Attribute VB_Name = "modDataValidator"
Option Explicit

'*******************************************************************************
' モジュール: modDataValidator
' 目的：     データ入力の検証機能の提供
' 作成日：   2025/01/17
'*******************************************************************************

' 検証結果の定義
Private Type ValidationResult
    IsValid As Boolean
    ErrorMessage As String
End Type

' 一般的な検証パターン
Private Const PATTERN_EMAIL As String = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
Private Const PATTERN_DATE As String = "^\d{4}/\d{2}/\d{2}$"
Private Const PATTERN_TIME As String = "^\d{2}:\d{2}(:\d{2})?$"
Private Const PATTERN_PHONE As String = "^[0-9-()+ ]{10,}$"

'*******************************************************************************
' 目的：    必須入力の検証
' 引数：    value - 検証対象の値
'           fieldName - フィールド名（エラーメッセージ用）
' 戻り値：  ValidationResult
'*******************************************************************************
Public Function ValidateRequired(ByVal value As String, _
                               ByVal fieldName As String) As ValidationResult
                               
    Dim result As ValidationResult
    
    ' 空文字列や空白文字のみの場合はエラー
    If Trim(value) = "" Then
        result.IsValid = False
        result.ErrorMessage = fieldName & "は必須項目です。"
    Else
        result.IsValid = True
    End If
    
    ValidateRequired = result
End Function

'*******************************************************************************
' 目的：    文字列の長さの検証
' 引数：    value - 検証対象の値
'           fieldName - フィールド名（エラーメッセージ用）
'           minLength - 最小長（オプション）
'           maxLength - 最大長（オプション）
' 戻り値：  ValidationResult
'*******************************************************************************
Public Function ValidateLength(ByVal value As String, _
                             ByVal fieldName As String, _
                             Optional ByVal minLength As Long = 0, _
                             Optional ByVal maxLength As Long = -1) As ValidationResult
                             
    Dim result As ValidationResult
    Dim valueLength As Long
    
    valueLength = Len(value)
    
    ' 最小長のチェック
    If valueLength < minLength Then
        result.IsValid = False
        result.ErrorMessage = fieldName & "は" & minLength & "文字以上で入力してください。"
        ValidateLength = result
        Exit Function
    End If
    
    ' 最大長のチェック（-1は無制限）
    If maxLength > 0 And valueLength > maxLength Then
        result.IsValid = False
        result.ErrorMessage = fieldName & "は" & maxLength & "文字以下で入力してください。"
        ValidateLength = result
        Exit Function
    End If
    
    result.IsValid = True
    ValidateLength = result
End Function

'*******************************************************************************
' 目的：    数値範囲の検証
' 引数：    value - 検証対象の値
'           fieldName - フィールド名（エラーメッセージ用）
'           minValue - 最小値（オプション）
'           maxValue - 最大値（オプション）
' 戻り値：  ValidationResult
'*******************************************************************************
Public Function ValidateNumberRange(ByVal value As Variant, _
                                  ByVal fieldName As String, _
                                  Optional ByVal minValue As Double = -1E+308, _
                                  Optional ByVal maxValue As Double = 1E+308) As ValidationResult
                                  
    Dim result As ValidationResult
    
    ' 数値形式の確認
    If Not IsNumeric(value) Then
        result.IsValid = False
        result.ErrorMessage = fieldName & "は数値で入力してください。"
        ValidateNumberRange = result
        Exit Function
    End If
    
    Dim numValue As Double
    numValue = CDbl(value)
    
    ' 最小値のチェック
    If numValue < minValue Then
        result.IsValid = False
        result.ErrorMessage = fieldName & "は" & minValue & "以上で入力してください。"
        ValidateNumberRange = result
        Exit Function
    End If
    
    ' 最大値のチェック
    If numValue > maxValue Then
        result.IsValid = False
        result.ErrorMessage = fieldName & "は" & maxValue & "以下で入力してください。"
        ValidateNumberRange = result
        Exit Function
    End If
    
    result.IsValid = True
    ValidateNumberRange = result
End Function

'*******************************************************************************
' 目的：    日付形式の検証
' 引数：    value - 検証対象の値
'           fieldName - フィールド名（エラーメッセージ用）
'           minDate - 最小日付（オプション）
'           maxDate - 最大日付（オプション）
' 戻り値：  ValidationResult
'*******************************************************************************
Public Function ValidateDate(ByVal value As String, _
                           ByVal fieldName As String, _
                           Optional ByVal minDate As Date, _
                           Optional ByVal maxDate As Date) As ValidationResult
                           
    Dim result As ValidationResult
    
    ' 日付形式の確認
    If Not value Like PATTERN_DATE Then
        result.IsValid = False
        result.ErrorMessage = fieldName & "はYYYY/MM/DD形式で入力してください。"
        ValidateDate = result
        Exit Function
    End If
    
    Dim dateValue As Date
    On Error Resume Next
    dateValue = CDate(value)
    If Err.Number <> 0 Then
        result.IsValid = False
        result.ErrorMessage = fieldName & "は有効な日付を入力してください。"
        ValidateDate = result
        Exit Function
    End If
    On Error GoTo 0
    
    ' 最小日付のチェック
    If Not IsEmpty(minDate) And dateValue < minDate Then
        result.IsValid = False
        result.ErrorMessage = fieldName & "は" & Format(minDate, "yyyy/mm/dd") & "以降の日付を入力してください。"
        ValidateDate = result
        Exit Function
    End If
    
    ' 最大日付のチェック
    If Not IsEmpty(maxDate) And dateValue > maxDate Then
        result.IsValid = False
        result.ErrorMessage = fieldName & "は" & Format(maxDate, "yyyy/mm/dd") & "以前の日付を入力してください。"
        ValidateDate = result
        Exit Function
    End If
    
    result.IsValid = True
    ValidateDate = result
End Function

'*******************************************************************************
' 目的：    メールアドレス形式の検証
' 引数：    value - 検証対象の値
'           fieldName - フィールド名（エラーメッセージ用）
' 戻り値：  ValidationResult
'*******************************************************************************
Public Function ValidateEmail(ByVal value As String, _
                            ByVal fieldName As String) As ValidationResult
                            
    Dim result As ValidationResult
    
    ' メールアドレス形式の確認
    If Not value Like PATTERN_EMAIL Then
        result.IsValid = False
        result.ErrorMessage = fieldName & "は有効なメールアドレス形式で入力してください。"
    Else
        result.IsValid = True
    End If
    
    ValidateEmail = result
End Function

'*******************************************************************************
' 目的：    電話番号形式の検証
' 引数：    value - 検証対象の値
'           fieldName - フィールド名（エラーメッセージ用）
' 戻り値：  ValidationResult
'*******************************************************************************
Public Function ValidatePhone(ByVal value As String, _
                            ByVal fieldName As String) As ValidationResult
                            
    Dim result As ValidationResult
    
    ' 電話番号形式の確認
    If Not value Like PATTERN_PHONE Then
        result.IsValid = False
        result.ErrorMessage = fieldName & "は有効な電話番号形式で入力してください。"
    Else
        result.IsValid = True
    End If
    
    ValidatePhone = result
End Function

'*******************************************************************************
' 目的：    正規表現パターンによる検証
' 引数：    value - 検証対象の値
'           pattern - 正規表現パターン
'           fieldName - フィールド名（エラーメッセージ用）
'           errorMessage - カスタムエラーメッセージ（オプション）
' 戻り値：  ValidationResult
'*******************************************************************************
Public Function ValidatePattern(ByVal value As String, _
                              ByVal pattern As String, _
                              ByVal fieldName As String, _
                              Optional ByVal errorMessage As String = "") As ValidationResult
                              
    Dim result As ValidationResult
    Dim regex As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = pattern
    End With
    
    ' パターンマッチング
    If Not regex.Test(value) Then
        result.IsValid = False
        If errorMessage = "" Then
            result.ErrorMessage = fieldName & "は正しい形式で入力してください。"
        Else
            result.ErrorMessage = errorMessage
        End If
    Else
        result.IsValid = True
    End If
    
    ValidatePattern = result
End Function

'*******************************************************************************
' 目的：    データ型の検証
' 引数：    value - 検証対象の値
'           expectedType - 期待するデータ型（"Number", "Date", "Boolean"など）
'           fieldName - フィールド名（エラーメッセージ用）
' 戻り値：  ValidationResult
'*******************************************************************************
Public Function ValidateDataType(ByVal value As Variant, _
                               ByVal expectedType As String, _
                               ByVal fieldName As String) As ValidationResult
                               
    Dim result As ValidationResult
    
    Select Case LCase(expectedType)
        Case "number"
            result.IsValid = IsNumeric(value)
            If Not result.IsValid Then
                result.ErrorMessage = fieldName & "は数値で入力してください。"
            End If
            
        Case "date"
            result.IsValid = IsDate(value)
            If Not result.IsValid Then
                result.ErrorMessage = fieldName & "は日付形式で入力してください。"
            End If
            
        Case "boolean"
            result.IsValid = IsBoolean(value)
            If Not result.IsValid Then
                result.ErrorMessage = fieldName & "は真偽値で入力してください。"
            End If
            
        Case Else
            result.IsValid = False
            result.ErrorMessage = "未対応のデータ型です: " & expectedType
    End Select
    
    ValidateDataType = result
End Function

'*******************************************************************************
' 目的：    複数の検証結果の組み合わせ
' 引数：    results - ValidationResultの配列
' 戻り値：  ValidationResult（エラーメッセージは改行で結合）
'*******************************************************************************
Public Function CombineResults(ParamArray results() As Variant) As ValidationResult
    Dim result As ValidationResult
    Dim i As Long
    Dim messages As String
    
    result.IsValid = True
    
    For i = LBound(results) To UBound(results)
        If Not results(i).IsValid Then
            result.IsValid = False
            If messages = "" Then
                messages = results(i).ErrorMessage
            Else
                messages = messages & vbNewLine & results(i).ErrorMessage
            End If
        End If
    Next i
    
    result.ErrorMessage = messages
    CombineResults = result
End Function
