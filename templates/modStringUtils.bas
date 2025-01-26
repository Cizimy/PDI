Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "modStringUtils"

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
''' 文字列を左側からパディングします
''' </summary>
''' <param name="baseStr">対象の文字列</param>
''' <param name="totalWidth">目標の長さ</param>
''' <param name="padChar">パディング文字（オプション）</param>
''' <returns>パディングされた文字列</returns>
Public Function PadLeft(ByVal baseStr As String, ByVal totalWidth As Long, _
                      Optional ByVal padChar As String = " ") As String
    If Not isInitialized Then InitializeModule
    
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.StartMeasurement "PadLeft"
    End If
    
    On Error GoTo ErrorHandler
    
    If Len(baseStr) >= totalWidth Then
        PadLeft = baseStr
    Else
        PadLeft = String(totalWidth - Len(baseStr), Left$(padChar, 1)) & baseStr
    End If
    
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.EndMeasurement "PadLeft"
    End If
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "文字列のパディング中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "PadLeft"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.EndMeasurement "PadLeft"
    End If
    PadLeft = baseStr
End Function

''' <summary>
''' 文字列を右側からパディングします
''' </summary>
''' <param name="baseStr">対象の文字列</param>
''' <param name="totalWidth">目標の長さ</param>
''' <param name="padChar">パディング文字（オプション）</param>
''' <returns>パディングされた文字列</returns>
Public Function PadRight(ByVal baseStr As String, ByVal totalWidth As Long, _
                       Optional ByVal padChar As String = " ") As String
    If Not isInitialized Then InitializeModule
    
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.StartMeasurement "PadRight"
    End If
    
    On Error GoTo ErrorHandler
    
    If Len(baseStr) >= totalWidth Then
        PadRight = baseStr
    Else
        PadRight = baseStr & String(totalWidth - Len(baseStr), Left$(padChar, 1))
    End If
    
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.EndMeasurement "PadRight"
    End If
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "文字列のパディング中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "PadRight"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.EndMeasurement "PadRight"
    End If
    PadRight = baseStr
End Function

''' <summary>
''' 文字列の前後の空白を削除します
''' </summary>
''' <param name="str">対象の文字列</param>
''' <returns>トリムされた文字列</returns>
Public Function TrimString(ByVal str As String) As String
    If Not isInitialized Then InitializeModule
    
    TrimString = Trim$(str)
End Function

''' <summary>
''' 文字列を指定された区切り文字で分割します
''' </summary>
''' <param name="str">対象の文字列</param>
''' <param name="delimiter">区切り文字</param>
''' <returns>分割された文字列の配列</returns>
Public Function SplitString(ByVal str As String, ByVal delimiter As String) As Variant
    If Not isInitialized Then InitializeModule
    
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.StartMeasurement "SplitString"
    End If
    
    On Error GoTo ErrorHandler
    
    SplitString = Split(str, delimiter)
    
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.EndMeasurement "SplitString"
    End If
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "文字列の分割中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "SplitString"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.EndMeasurement "SplitString"
    End If
    SplitString = Array()
End Function

''' <summary>
''' 文字列配列を指定された区切り文字で結合します
''' </summary>
''' <param name="arr">文字列配列</param>
''' <param name="delimiter">区切り文字</param>
''' <returns>結合された文字列</returns>
Public Function JoinStrings(ByRef arr As Variant, Optional ByVal delimiter As String = "") As String
    If Not isInitialized Then InitializeModule
    
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.StartMeasurement "JoinStrings"
    End If
    
    On Error GoTo ErrorHandler
    
    JoinStrings = Join(arr, delimiter)
    
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.EndMeasurement "JoinStrings"
    End If
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "文字列の結合中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "JoinStrings"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    If Not performanceMonitor Is Nothing Then
        performanceMonitor.EndMeasurement "JoinStrings"
    End If
    JoinStrings = ""
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