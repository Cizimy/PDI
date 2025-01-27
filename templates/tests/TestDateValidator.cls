VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestDateValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "TestDateValidator"

' ======================
' プライベート変数
' ======================
Private mValidator As DateValidator
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
Public Sub TestMinDateValidation()
    ' 準備
    Set mValidator = New DateValidator
    Set mValidator.Error = mMockError
    
    Dim minDateRule As MinDateRule
    Set minDateRule = New MinDateRule
    minDateRule.MinDate = DateSerial(2000, 1, 1)
    mValidator.AddRule minDateRule
    
    ' 実行と検証
    Debug.Assert Not mValidator.Validate(DateSerial(1999, 12, 31)), "最小日付より前の日付は失敗するべき"
    Debug.Assert mValidator.Validate(DateSerial(2000, 1, 1)), "最小日付と同じ日付は成功するべき"
    Debug.Assert mValidator.Validate(DateSerial(2000, 1, 2)), "最小日付より後の日付は成功するべき"
End Sub

Public Sub TestMaxDateValidation()
    ' 準備
    Set mValidator = New DateValidator
    Set mValidator.Error = mMockError
    
    Dim maxDateRule As MaxDateRule
    Set maxDateRule = New MaxDateRule
    maxDateRule.MaxDate = DateSerial(2000, 12, 31)
    mValidator.AddRule maxDateRule
    
    ' 実行と検証
    Debug.Assert mValidator.Validate(DateSerial(2000, 12, 30)), "最大日付より前の日付は成功するべき"
    Debug.Assert mValidator.Validate(DateSerial(2000, 12, 31)), "最大日付と同じ日付は成功するべき"
    Debug.Assert Not mValidator.Validate(DateSerial(2001, 1, 1)), "最大日付より後の日付は失敗するべき"
End Sub

Public Sub TestAllowFutureDateValidation()
    ' 準備
    Set mValidator = New DateValidator
    Set mValidator.Error = mMockError
    
    Dim futureDateRule As AllowFutureDateRule
    Set futureDateRule = New AllowFutureDateRule
    futureDateRule.AllowFutureDate = False
    mValidator.AddRule futureDateRule
    
    ' 実行と検証
    Debug.Assert mValidator.Validate(DateAdd("d", -1, Date)), "過去の日付は成功するべき"
    Debug.Assert mValidator.Validate(Date), "現在の日付は成功するべき"
    Debug.Assert Not mValidator.Validate(DateAdd("d", 1, Date)), "未来の日付は失敗するべき"
End Sub

Public Sub TestAllowPastDateValidation()
    ' 準備
    Set mValidator = New DateValidator
    Set mValidator.Error = mMockError
    
    Dim pastDateRule As AllowPastDateRule
    Set pastDateRule = New AllowPastDateRule
    pastDateRule.AllowPastDate = False
    mValidator.AddRule pastDateRule
    
    ' 実行と検証
    Debug.Assert Not mValidator.Validate(DateAdd("d", -1, Date)), "過去の日付は失敗するべき"
    Debug.Assert mValidator.Validate(Date), "現在の日付は成功するべき"
    Debug.Assert mValidator.Validate(DateAdd("d", 1, Date)), "未来の日付は成功するべき"
End Sub

Public Sub TestBirthDateValidation()
    ' 準備
    Set mValidator = CreateBirthDateValidator(mMockError)
    
    ' 実行と検証
    Debug.Assert Not mValidator.Validate(DateSerial(1899, 12, 31)), "1900年以前の日付は失敗するべき"
    Debug.Assert mValidator.Validate(DateSerial(1900, 1, 1)), "1900年1月1日は成功するべき"
    Debug.Assert mValidator.Validate(DateAdd("d", -1, Date)), "昨日の日付は成功するべき"
    Debug.Assert mValidator.Validate(Date), "今日の日付は成功するべき"
    Debug.Assert Not mValidator.Validate(DateAdd("d", 1, Date)), "未来の日付は失敗するべき"
End Sub

Public Sub TestExpiryDateValidation()
    ' 準備
    Set mValidator = CreateExpiryDateValidator(mMockError)
    
    ' 実行と検証
    Debug.Assert Not mValidator.Validate(DateAdd("d", -1, Date)), "過去の日付は失敗するべき"
    Debug.Assert mValidator.Validate(Date), "今日の日付は成功するべき"
    Debug.Assert mValidator.Validate(DateAdd("yyyy", 50, Date)), "50年後の日付は成功するべき"
    Debug.Assert mValidator.Validate(DateAdd("yyyy", 100, Date)), "100年後の日付は成功するべき"
    Debug.Assert Not mValidator.Validate(DateAdd("yyyy", 101, Date)), "100年より先の日付は失敗するべき"
End Sub

Public Sub TestMultipleRules()
    ' 準備
    Set mValidator = New DateValidator
    Set mValidator.Error = mMockError
    
    Dim minDateRule As MinDateRule
    Set minDateRule = New MinDateRule
    minDateRule.MinDate = DateSerial(2000, 1, 1)
    mValidator.AddRule minDateRule
    
    Dim maxDateRule As MaxDateRule
    Set maxDateRule = New MaxDateRule
    maxDateRule.MaxDate = DateSerial(2100, 12, 31)
    mValidator.AddRule maxDateRule
    
    Dim futureDateRule As AllowFutureDateRule
    Set futureDateRule = New AllowFutureDateRule
    futureDateRule.AllowFutureDate = True
    mValidator.AddRule futureDateRule
    
    Dim pastDateRule As AllowPastDateRule
    Set pastDateRule = New AllowPastDateRule
    pastDateRule.AllowPastDate = True
    mValidator.AddRule pastDateRule
    
    ' 実行と検証
    Debug.Assert Not mValidator.Validate(DateSerial(1999, 12, 31)), "範囲外の過去の日付は失敗するべき"
    Debug.Assert mValidator.Validate(DateSerial(2000, 1, 1)), "範囲内の過去の日付は成功するべき"
    Debug.Assert mValidator.Validate(Date), "現在の日付は成功するべき"
    Debug.Assert mValidator.Validate(DateSerial(2100, 12, 31)), "範囲内の未来の日付は成功するべき"
    Debug.Assert Not mValidator.Validate(DateSerial(2101, 1, 1)), "範囲外の未来の日付は失敗するべき"
End Sub