VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestStringValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "TestStringValidator"

' ======================
' プライベート変数
' ======================
Private mValidator As StringValidator
Private mMockError As MockError

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    Set mMockError = New MockError
End Sub

Private Sub Class_Terminate()
    Set mValidator = Nothing
    Set mMockError = Nothing
End Sub

' ======================
' テストメソッド
' ======================
Public Sub TestMinLengthValidation()
    ' 準備
    Set mValidator = New StringValidator
    Set mValidator.Error = mMockError
    
    Dim minLengthRule As MinLengthRule
    Set minLengthRule = New MinLengthRule
    minLengthRule.MinLength = 5
    mValidator.AddRule minLengthRule
    
    ' 実行と検証
    Debug.Assert Not mValidator.Validate("abc"), "最小長より短い文字列は失敗するべき"
    Debug.Assert mValidator.Validate("abcde"), "最小長と同じ長さの文字列は成功するべき"
    Debug.Assert mValidator.Validate("abcdef"), "最小長より長い文字列は成功するべき"
End Sub

Public Sub TestMaxLengthValidation()
    ' 準備
    Set mValidator = New StringValidator
    Set mValidator.Error = mMockError
    
    Dim maxLengthRule As MaxLengthRule
    Set maxLengthRule = New MaxLengthRule
    maxLengthRule.MaxLength = 5
    mValidator.AddRule maxLengthRule
    
    ' 実行と検証
    Debug.Assert mValidator.Validate("abc"), "最大長より短い文字列は成功するべき"
    Debug.Assert mValidator.Validate("abcde"), "最大長と同じ長さの文字列は成功するべき"
    Debug.Assert Not mValidator.Validate("abcdef"), "最大長より長い文字列は失敗するべき"
End Sub

Public Sub TestPatternValidation()
    ' 準備
    Set mValidator = New StringValidator
    Set mValidator.Error = mMockError
    
    Dim patternRule As PatternRule
    Set patternRule = New PatternRule
    patternRule.Pattern = "^[0-9]+$"
    mValidator.AddRule patternRule
    
    ' 実行と検証
    Debug.Assert mValidator.Validate("12345"), "数字のみの文字列は成功するべき"
    Debug.Assert Not mValidator.Validate("abc"), "数字以外の文字列は失敗するべき"
    Debug.Assert Not mValidator.Validate("123abc"), "数字と文字が混在する文字列は失敗するべき"
End Sub

Public Sub TestEmailValidation()
    ' 準備
    Set mValidator = CreateEmailValidator(mMockError)
    
    ' 実行と検証
    Debug.Assert mValidator.Validate("test@example.com"), "正しい形式のメールアドレスは成功するべき"
    Debug.Assert Not mValidator.Validate("invalid-email"), "@を含まないメールアドレスは失敗するべき"
    Debug.Assert Not mValidator.Validate("test@"), "ドメインがないメールアドレスは失敗するべき"
    Debug.Assert Not mValidator.Validate("@example.com"), "ローカル部がないメールアドレスは失敗するべき"
End Sub

Public Sub TestPhoneValidation()
    ' 準備
    Set mValidator = CreatePhoneValidator(mMockError)
    
    ' 実行と検証
    Debug.Assert mValidator.Validate("03-1234-5678"), "正しい形式の電話番号は成功するべき"
    Debug.Assert mValidator.Validate("0312345678"), "ハイフンなしの電話番号は成功するべき"
    Debug.Assert Not mValidator.Validate("abc-defg-hijk"), "数字以外の文字列は失敗するべき"
    Debug.Assert Not mValidator.Validate("03-1234"), "短すぎる電話番号は失敗するべき"
    Debug.Assert Not mValidator.Validate("03-1234-5678-9012"), "長すぎる電話番号は失敗するべき"
End Sub

Public Sub TestMultipleRules()
    ' 準備
    Set mValidator = New StringValidator
    Set mValidator.Error = mMockError
    
    Dim minLengthRule As MinLengthRule
    Set minLengthRule = New MinLengthRule
    minLengthRule.MinLength = 5
    mValidator.AddRule minLengthRule
    
    Dim maxLengthRule As MaxLengthRule
    Set maxLengthRule = New MaxLengthRule
    maxLengthRule.MaxLength = 10
    mValidator.AddRule maxLengthRule
    
    Dim patternRule As PatternRule
    Set patternRule = New PatternRule
    patternRule.Pattern = "^[a-zA-Z]+$"
    mValidator.AddRule patternRule
    
    ' 実行と検証
    Debug.Assert Not mValidator.Validate("abc"), "短すぎる文字列は失敗するべき"
    Debug.Assert Not mValidator.Validate("abcdefghijk"), "長すぎる文字列は失敗するべき"
    Debug.Assert Not mValidator.Validate("abc123"), "数字を含む文字列は失敗するべき"
    Debug.Assert mValidator.Validate("abcdef"), "すべてのルールを満たす文字列は成功するべき"
End Sub