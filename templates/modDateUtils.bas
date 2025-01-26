Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "modDateUtils"

' ======================
' 定数定義
' ======================
Private Const DEFAULT_DATE_FORMAT As String = "yyyy/mm/dd"
Private Const DEFAULT_TIME_FORMAT As String = "hh:nn:ss"
Private Const DEFAULT_DATETIME_FORMAT As String = "yyyy/mm/dd hh:nn:ss"

' ======================
' プライベート変数
' ======================
Private performanceMonitor As clsPerformanceMonitor
Private isInitialized As Boolean

' ======================
' 初期化・終了処理
' ======================
Public Sub InitializeModule()
    If isInitialized Then Exit Sub
    
    Set performanceMonitor = New clsPerformanceMonitor
    isInitialized = True
End Sub

Public Sub TerminateModule()
    If Not isInitialized Then Exit Sub
    
    Set performanceMonitor = Nothing
    isInitialized = False
End Sub

' ======================
' 公開関数
' ======================

''' <summary>
''' 日付の妥当性を確認します
''' </summary>
''' <param name="testDate">確認する日付</param>
''' <returns>有効な日付の場合True</returns>
Public Function IsValidDate(ByVal testDate As Variant) As Boolean
    If Not isInitialized Then InitializeModule
    
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.StartMeasurement "IsValidDate"
    End If
    
    On Error Resume Next
    IsValidDate = IsDate(testDate)
    On Error GoTo 0
    
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.EndMeasurement "IsValidDate"
    End If
End Function

''' <summary>
''' 日付に指定された期間を加算します
''' </summary>
''' <param name="interval">期間の単位</param>
''' <param name="number">加算する数</param>
''' <param name="dateValue">対象の日付</param>
''' <returns>加算後の日付</returns>
Public Function DateAdd(ByVal interval As String, ByVal number As Double, _
                      ByVal dateValue As Date) As Date
    If Not isInitialized Then InitializeModule
    
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.StartMeasurement "DateAdd"
    End If
    
    On Error GoTo ErrorHandler
    
    DateAdd = VBA.DateAdd(interval, number, dateValue)
    
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.EndMeasurement "DateAdd"
    End If
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "日付の加算中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "DateAdd"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.EndMeasurement "DateAdd"
    End If
    DateAdd = dateValue
End Function

''' <summary>
''' 2つの日付の差分を計算します
''' </summary>
''' <param name="interval">期間の単位</param>
''' <param name="date1">日付1</param>
''' <param name="date2">日付2</param>
''' <returns>日付の差分</returns>
Public Function DateDiff(ByVal interval As String, ByVal date1 As Date, _
                       ByVal date2 As Date) As Long
    If Not isInitialized Then InitializeModule
    
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.StartMeasurement "DateDiff"
    End If
    
    On Error GoTo ErrorHandler
    
    DateDiff = VBA.DateDiff(interval, date1, date2)
    
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.EndMeasurement "DateDiff"
    End If
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "日付の差分計算中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "DateDiff"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.EndMeasurement "DateDiff"
    End If
    DateDiff = 0
End Function

''' <summary>
''' 日付を指定された形式でフォーマットします
''' </summary>
''' <param name="dateValue">対象の日付</param>
''' <param name="format">フォーマット文字列（オプション）</param>
''' <returns>フォーマットされた日付文字列</returns>
Public Function FormatDate(ByVal dateValue As Date, _
                         Optional ByVal format As String = DEFAULT_DATE_FORMAT) As String
    If Not isInitialized Then InitializeModule
    
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.StartMeasurement "FormatDate"
    End If
    
    On Error GoTo ErrorHandler
    
    FormatDate = Format$(dateValue, format)
    
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.EndMeasurement "FormatDate"
    End If
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "日付のフォーマット中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "FormatDate"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.EndMeasurement "FormatDate"
    End If
    FormatDate = Format$(dateValue, DEFAULT_DATE_FORMAT)
End Function

''' <summary>
''' 現在の日時を取得します
''' </summary>
''' <returns>現在の日時</returns>
Public Function GetCurrentDateTime() As Date
    If Not isInitialized Then InitializeModule
    GetCurrentDateTime = Now
End Function

''' <summary>
''' 指定された日付が営業日かどうかを確認します
''' </summary>
''' <param name="dateValue">確認する日付</param>
''' <returns>営業日の場合True</returns>
Public Function IsBusinessDay(ByVal dateValue As Date) As Boolean
    If Not isInitialized Then InitializeModule
    
    ' 土曜日(7)または日曜日(1)の場合はFalse
    IsBusinessDay = Not (Weekday(dateValue, vbSunday) = 1 Or _
                        Weekday(dateValue, vbSunday) = 7)
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
        Set GetPerformanceMonitor = performanceMonitor
    End Function
#End If