Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "modMathUtils"

' ======================
' 定数定義
' ======================
Private Const EPSILON As Double = 0.0000000001 ' 浮動小数点比較用の許容誤差

' ======================
' プライベート変数
' ======================
Private mPerformanceMonitor As clsPerformanceMonitor
Private mIsInitialized As Boolean

' ======================
' 初期化・終了処理
' ======================
Public Sub InitializeModule()
    If mIsInitialized Then Exit Sub
    
    Set mPerformanceMonitor = New clsPerformanceMonitor
    mIsInitialized = True
End Sub

Public Sub TerminateModule()
    If Not mIsInitialized Then Exit Sub
    
    Set mPerformanceMonitor = Nothing
    mIsInitialized = False
End Sub

' ======================
' 公開関数
' ======================

''' <summary>
''' 安全な除算を行います
''' </summary>
''' <param name="numerator">分子</param>
''' <param name="denominator">分母</param>
''' <param name="defaultValue">分母が0の場合の戻り値</param>
''' <returns>除算結果、またはデフォルト値</returns>
Public Function SafeDivide(ByVal numerator As Double, ByVal denominator As Double, _
                         Optional ByVal defaultValue As Variant = 0) As Variant
    If Not mIsInitialized Then InitializeModule
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.StartMeasurement "SafeDivide"
    End If
    
    On Error GoTo ErrorHandler
    
    If Abs(denominator) < EPSILON Then
        ' 分母が0の場合の警告を出力
        Dim errDetail As typErrorDetail
        With errDetail
            .ErrorCode = ERR_DIVISION_BY_ZERO
            .Description = "分母が0のため、デフォルト値" & CStr(defaultValue) & "を返します。(分子: " & CStr(numerator) & ")"
            .Category = ECGeneral
            .Source = MODULE_NAME
            .ProcedureName = "SafeDivide"
            .StackTrace = modStackTrace.GetStackTrace()
            .OccurredAt = Now
        End With
        modError.HandleError errDetail
        
        SafeDivide = defaultValue
    Else
        SafeDivide = numerator / denominator
    End If
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "SafeDivide"
    End If
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "除算中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "SafeDivide"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "SafeDivide"
    End If
    SafeDivide = defaultValue
End Function

''' <summary>
''' 数値が指定された範囲内かどうかを確認します
''' </summary>
''' <param name="value">確認する値</param>
''' <param name="minValue">最小値</param>
''' <param name="maxValue">最大値</param>
''' <returns>範囲内の場合True</returns>
Public Function IsInRange(ByVal value As Double, ByVal minValue As Double, _
                        ByVal maxValue As Double) As Boolean
    If Not mIsInitialized Then InitializeModule
    
    IsInRange = (value >= minValue And value <= maxValue)
End Function

''' <summary>
''' 値を指定された範囲内に収めます
''' </summary>
''' <param name="value">対象の値</param>
''' <param name="minValue">最小値</param>
''' <param name="maxValue">最大値</param>
''' <returns>範囲内に収められた値</returns>
Public Function Clamp(ByVal value As Double, ByVal minValue As Double, _
                     ByVal maxValue As Double) As Double
    If Not mIsInitialized Then InitializeModule
    
    If value < minValue Then
        Clamp = minValue
    ElseIf value > maxValue Then
        Clamp = maxValue
    Else
        Clamp = value
    End If
End Function

''' <summary>
''' 指定された精度で四捨五入します
''' </summary>
''' <param name="value">対象の値</param>
''' <param name="decimals">小数点以下の桁数</param>
''' <returns>四捨五入された値</returns>
Public Function Round(ByVal value As Double, Optional ByVal decimals As Long = 0) As Double
    If Not mIsInitialized Then InitializeModule
    
    Dim factor As Double
    factor = 10 ^ decimals
    Round = Fix(value * factor + 0.5) / factor
End Function

' ======================
' テストサポート機能
' 警告: これらのメソッドは開発時のテスト目的でのみ使用し、
' 本番環境では使用しないでください。
' ======================
#If DEBUG Then
    ''' <summary>
    ''' モジュールの状態を初期化（テスト用）
    ''' </summary>
    Private Sub ResetModule()
        TerminateModule
        InitializeModule
    End Sub
    
    ''' <summary>
    ''' パフォーマンスモニターの参照を取得（テスト用）
    ''' </summary>
    Private Function GetPerformanceMonitor() As clsPerformanceMonitor
        Set GetPerformanceMonitor = mPerformanceMonitor
    End Function
#End If