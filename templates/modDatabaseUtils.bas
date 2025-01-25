Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "modDatabaseUtils"

' ======================
' 定数定義
' ======================
Private Const MAX_RETRY_COUNT As Long = 3
Private Const RETRY_INTERVAL_MS As Long = 1000

' ======================
' プライベート変数
' ======================
Private mPerformanceMonitor As clsPerformanceMonitor
Private mIsInitialized As Boolean
Private mDefaultConnection As Object ' ADODB.Connection

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
    
    CloseConnection
    Set mPerformanceMonitor = Nothing
    mIsInitialized = False
End Sub

' ======================
' 公開関数
' ======================

''' <summary>
''' データベース接続文字列を取得します
''' </summary>
''' <returns>接続文字列</returns>
Public Function GetConnectionString() As String
    If Not mIsInitialized Then InitializeModule
    
    On Error GoTo ErrorHandler
    
    ' 設定から接続文字列を取得
    GetConnectionString = modConfig.Settings.DatabaseConnectionString
    
    If GetConnectionString = "" Then
        Dim errDetail As typErrorDetail
        With errDetail
            .ErrorCode = ERR_DATABASE_CONNECTION_FAILED
            .Description = "データベース接続文字列が設定されていません。"
            .Source = MODULE_NAME
            .ProcedureName = "GetConnectionString"
        End With
        modError.HandleError errDetail
    End If
    Exit Function

ErrorHandler:
    Dim errDetail2 As typErrorDetail
    With errDetail2
        .ErrorCode = ERR_DATABASE_CONNECTION_FAILED
        .Description = "接続文字列の取得中にエラーが発生しました: " & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "GetConnectionString"
    End With
    modError.HandleError errDetail2
    GetConnectionString = ""
End Function

''' <summary>
''' データベース接続を取得します
''' </summary>
''' <returns>データベース接続オブジェクト</returns>
Public Function GetConnection() As Object ' ADODB.Connection
    If Not mIsInitialized Then InitializeModule
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.StartMeasurement "GetConnection"
    End If
    
    On Error GoTo ErrorHandler
    
    ' 既存の接続を確認
    If Not mDefaultConnection Is Nothing Then
        If mDefaultConnection.State = 1 Then ' adStateOpen
            Set GetConnection = mDefaultConnection
            GoTo CleanExit
        End If
    End If
    
    ' 新しい接続を作成
    Dim connStr As String
    connStr = GetConnectionString()
    If connStr = "" Then Exit Function
    
    Set mDefaultConnection = CreateObject("ADODB.Connection")
    mDefaultConnection.ConnectionString = connStr
    mDefaultConnection.Open
    
    Set GetConnection = mDefaultConnection
    
CleanExit:
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "GetConnection"
    End If
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_DATABASE_CONNECTION_FAILED
        .Description = "データベース接続の取得中にエラーが発生しました: " & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "GetConnection"
    End With
    modError.HandleError errDetail
    Set GetConnection = Nothing
End Function

''' <summary>
''' データベース接続を閉じます
''' </summary>
Public Sub CloseConnection()
    If Not mDefaultConnection Is Nothing Then
        On Error Resume Next
        If mDefaultConnection.State = 1 Then ' adStateOpen
            mDefaultConnection.Close
        End If
        Set mDefaultConnection = Nothing
        On Error GoTo 0
    End If
End Sub

''' <summary>
''' データベース接続をテストします
''' </summary>
''' <returns>接続成功の場合True</returns>
Public Function TestConnection() As Boolean
    If Not mIsInitialized Then InitializeModule
    
    Dim conn As Object
    Set conn = GetConnection()
    
    TestConnection = Not (conn Is Nothing)
    
    If Not conn Is Nothing Then
        If conn.State = 1 Then ' adStateOpen
            TestConnection = True
        End If
    End If
End Function

''' <summary>
''' SQLクエリを実行し、結果を取得します
''' </summary>
''' <param name="sql">SQLクエリ</param>
''' <param name="params">パラメータ配列（オプション）</param>
''' <returns>レコードセット</returns>
Public Function ExecuteQuery(ByVal sql As String, _
                           Optional ByRef params As Variant) As Object ' ADODB.Recordset
    If Not mIsInitialized Then InitializeModule
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.StartMeasurement "ExecuteQuery"
    End If
    
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Set conn = GetConnection()
    If conn Is Nothing Then Exit Function
    
    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        Set .ActiveConnection = conn
        .CommandText = sql
        .CommandType = 1 ' adCmdText
        
        ' パラメータの設定
        If Not IsMissing(params) Then
            If IsArray(params) Then
                Dim i As Long
                For i = LBound(params) To UBound(params)
                    .Parameters.Append .CreateParameter("p" & i, GetParameterType(params(i)), 1, , params(i))
                Next i
            End If
        End If
        
        Set ExecuteQuery = .Execute
    End With
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "ExecuteQuery"
    End If
    Exit Function

ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_DATABASE_QUERY_FAILED
        .Description = "クエリの実行中にエラーが発生しました: " & Err.Description
        .Source = MODULE_NAME
        .ProcedureName = "ExecuteQuery"
    End With
    modError.HandleError errDetail
    Set ExecuteQuery = Nothing
End Function

' ======================
' プライベート関数
' ======================
Private Function GetParameterType(ByVal value As Variant) As Integer
    Select Case VarType(value)
        Case vbInteger, vbLong
            GetParameterType = 3 ' adInteger
        Case vbSingle, vbDouble
            GetParameterType = 5 ' adDouble
        Case vbString
            GetParameterType = 200 ' adVarChar
        Case vbDate
            GetParameterType = 7 ' adDate
        Case vbBoolean
            GetParameterType = 11 ' adBoolean
        Case Else
            GetParameterType = 12 ' adVariant
    End Select
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