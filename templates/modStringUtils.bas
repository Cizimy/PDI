Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "modStringUtils"

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
''' 文字列を左側からパディングします
''' </summary>
''' <param name="baseStr">対象の文字列</param>
''' <param name="totalWidth">目標の長さ</param>
''' <param name="padChar">パディング文字（オプション）</param>
''' <returns>パディングされた文字列</returns>
Public Function PadLeft(ByVal baseStr As String, ByVal totalWidth As Long, _
                      Optional ByVal padChar As String = " ") As String
    If Not mIsInitialized Then InitializeModule
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.StartMeasurement "PadLeft"
    End If
    
    On Error GoTo ErrorHandler
    
    If Len(baseStr) >= totalWidth Then
        PadLeft = baseStr
    Else
        PadLeft = String(totalWidth - Len(baseStr), Left$(padChar, 1)) & baseStr
    End If
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "PadLeft"
    End If
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "文字列のパディング中にエラーが発生しました: " & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "PadLeft"
    End With
    modError.HandleError errDetail
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
    If Not mIsInitialized Then InitializeModule
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.StartMeasurement "PadRight"
    End If
    
    On Error GoTo ErrorHandler
    
    If Len(baseStr) >= totalWidth Then
        PadRight = baseStr
    Else
        PadRight = baseStr & String(totalWidth - Len(baseStr), Left$(padChar, 1))
    End If
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "PadRight"
    End If
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "文字列のパディング中にエラーが発生しました: " & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "PadRight"
    End With
    modError.HandleError errDetail
    PadRight = baseStr
End Function

''' <summary>
''' 文字列の前後の空白を削除します
''' </summary>
''' <param name="str">対象の文字列</param>
''' <returns>トリムされた文字列</returns>
Public Function TrimString(ByVal str As String) As String
    If Not mIsInitialized Then InitializeModule
    
    TrimString = Trim$(str)
End Function

''' <summary>
''' 文字列を指定された区切り文字で分割します
''' </summary>
''' <param name="str">対象の文字列</param>
''' <param name="delimiter">区切り文字</param>
''' <returns>分割された文字列の配列</returns>
Public Function SplitString(ByVal str As String, ByVal delimiter As String) As Variant
    If Not mIsInitialized Then InitializeModule
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.StartMeasurement "SplitString"
    End If
    
    On Error GoTo ErrorHandler
    
    SplitString = Split(str, delimiter)
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "SplitString"
    End If
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "文字列の分割中にエラーが発生しました: " & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "SplitString"
    End With
    modError.HandleError errDetail
    SplitString = Array()
End Function

''' <summary>
''' 文字列配列を指定された区切り文字で結合します
''' </summary>
''' <param name="arr">文字列配列</param>
''' <param name="delimiter">区切り文字</param>
''' <returns>結合された文字列</returns>
Public Function JoinStrings(ByRef arr As Variant, Optional ByVal delimiter As String = "") As String
    If Not mIsInitialized Then InitializeModule
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.StartMeasurement "JoinStrings"
    End If
    
    On Error GoTo ErrorHandler
    
    JoinStrings = Join(arr, delimiter)
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "JoinStrings"
    End If
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_UNEXPECTED
        .Description = "文字列の結合中にエラーが発生しました: " & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "JoinStrings"
    End With
    modError.HandleError errDetail
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
        Set GetPerformanceMonitor = mPerformanceMonitor
    End Function
#End If