Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "modDatabaseUtils"

' ======================
' 定数定義
' ======================
Private Const ERR_MODULE_NOT_INITIALIZED As String = "モジュールが初期化されていません。"
Private Const DEFAULT_CONNECTION_STRING As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=default.accdb;"

' ======================
' プライベート変数
' ======================
Private mPerformanceMonitor As clsPerformanceMonitor
Private mIsInitialized As Boolean
Private mLock As clsLock
Private mDefaultConnection As Object ' ADODB.Connection
Private mConfig As IDatabaseConfig ' データベース設定
Private mConnectionPool As ConnectionPool ' コネクションプール

' ======================
' 初期化・終了処理
' ======================
''' <summary>モジュールを初期化します</summary>
''' <param name="config">データベース設定を提供するインターフェース（必須）</param>
Public Sub InitializeModule(ByVal config As IDatabaseConfig)
    If mIsInitialized Then Exit Sub
    
    Set mPerformanceMonitor = New clsPerformanceMonitor
    If config Is Nothing Then
        Err.Raise vbObjectError + 1001, MODULE_NAME, _
            "データベース設定が指定されていません。"
    End If
    Set mConfig = config
    Set mLock = New clsLock
    Set mConnectionPool = New ConnectionPool
    
    ' コネクションプールの初期化
    With mConnectionPool
        .MinPoolSize = CLng(mConfig.GetDatabaseSetting("MinPoolSize"))
        .MaxPoolSize = CLng(mConfig.GetDatabaseSetting("MaxPoolSize"))
        .ConnectionTimeout = mConfig.ConnectionTimeout
    End With
    
    mIsInitialized = True
End Sub

Public Sub TerminateModule()
    If Not mIsInitialized Then Exit Sub
    
    CloseConnection
    Set mPerformanceMonitor = Nothing
    Set mLock = Nothing
    Set mConfig = Nothing
    Set mConnectionPool = Nothing
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
    If Not mIsInitialized Then Err.Raise vbObjectError + 1002, MODULE_NAME, ERR_MODULE_NOT_INITIALIZED
    
    On Error GoTo ErrorHandler

    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.StartMeasurement "GetConnectionString"
    End If
    
    mLock.AcquireLock
    
    ' IDatabaseConfigから接続文字列を取得
    GetConnectionString = mConfig.GetConnectionString
    
    mLock.ReleaseLock
    
    ' 接続文字列が空の場合、デフォルト値を使用
    If GetConnectionString = "" Then
        ' デフォルト接続文字列を使用する前に警告をログ
        LogWarning "接続文字列が設定されていません。デフォルト値を使用します。", _
                  "GetConnectionString"
        
        GetConnectionString = DEFAULT_CONNECTION_STRING
    End If
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "GetConnectionString"
    End If
    Exit Function

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrDatabaseConnectionFailed
        .Description = "接続文字列の取得中にエラーが発生しました: " & Err.Description
        .Category = ECDatabase
        .Source = MODULE_NAME
        .ProcedureName = "GetConnectionString"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "GetConnectionString"
    End If
    GetConnectionString = DEFAULT_CONNECTION_STRING
End Function

''' <summary>
''' データベース接続を取得します
''' </summary>
''' <returns>データベース接続オブジェクト</returns>
Public Function GetConnection() As Object ' ADODB.Connection
    If Not mIsInitialized Then Err.Raise vbObjectError + 1002, MODULE_NAME, ERR_MODULE_NOT_INITIALIZED
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.StartMeasurement "GetConnection"
    End If
    
    On Error GoTo ErrorHandler
    
    mLock.AcquireLock
    
    ' コネクションプールから接続を取得
    Set GetConnection = mConnectionPool.GetConnection(GetConnectionString())
    
    GoTo CleanupAndExit

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrDatabaseConnectionFailed
        .Description = "データベース接続の取得中にエラーが発生しました: " & Err.Description
        .Category = ECDatabase
        .Source = MODULE_NAME
        .ProcedureName = "GetConnection"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    Set GetConnection = Nothing

CleanupAndExit:
    mLock.ReleaseLock
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "GetConnection"
    End If
End Function

''' <summary>
''' データベース接続を閉じます
''' </summary>
Public Sub CloseConnection()
    If Not mConnectionPool Is Nothing Then
        mLock.AcquireLock
        
        On Error Resume Next
        mConnectionPool.ReleaseAllConnections
        
        mLock.ReleaseLock
        On Error GoTo 0
    End If
End Sub

''' <summary>
''' データベース接続をテストします
''' </summary>
''' <returns>接続成功の場合True</returns>
Public Function TestConnection() As Boolean
    If Not mIsInitialized Then Err.Raise vbObjectError + 1002, MODULE_NAME, ERR_MODULE_NOT_INITIALIZED
    
    Dim conn As Object
    Set conn = GetConnection()
    
    TestConnection = Not (conn Is Nothing)
    
    If Not conn Is Nothing Then
        If conn.State = 1 Then ' adStateOpen
            TestConnection = True
            mConnectionPool.ReleaseConnection conn
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
    If Not mIsInitialized Then Err.Raise vbObjectError + 1002, MODULE_NAME, ERR_MODULE_NOT_INITIALIZED
    
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
        .CommandTimeout = mConfig.CommandTimeout
        
        ' パラメータの設定
        If Not IsMissing(params) Then
            ' 単一値のパラメータを配列に変換
            Dim paramArray As Variant
            If IsArray(params) Then
                paramArray = params
            Else
                ReDim paramArray(0)
                paramArray(0) = params
            End If
            
            ' パラメータのバリデーション
            ValidateParameters paramArray
            
            ' パラメータの追加
            Dim i As Long
            For i = LBound(paramArray) To UBound(paramArray)
                Dim paramValue As Variant
                paramValue = paramArray(i)
                If Not IsNull(paramValue) Then
                    .Parameters.Append .CreateParameter("p" & i, GetParameterType(paramValue), 1, , paramValue)
                End If
            Next i
        End If
        
        Set ExecuteQuery = .Execute
    End With
    
    ' 接続をプールに返却
    mConnectionPool.ReleaseConnection conn
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "ExecuteQuery"
    End If
    Exit Function

ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrDatabaseQueryFailed
        .Description = "クエリの実行中にエラーが発生しました: " & Err.Description
        .Category = ECDatabase
        .Source = MODULE_NAME
        .ProcedureName = "ExecuteQuery"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "ExecuteQuery"
    End If
    Set ExecuteQuery = Nothing
End Function

' ======================
' プライベート関数
' ======================
Private Function GetParameterType(ByVal Value As Variant) As Integer
    Select Case VarType(Value)
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

Private Sub ValidateParameters(ByRef params As Variant)
    If Not IsArray(params) Then Exit Sub
    
    Dim i As Long
    For i = LBound(params) To UBound(params)
        If Not IsNull(params(i)) Then
            Select Case VarType(params(i))
                Case vbInteger, vbLong, vbSingle, vbDouble, vbString, vbDate, vbBoolean
                    ' サポートされている型
                Case Else
                    Err.Raise vbObjectError + 1003, MODULE_NAME, _
                        "サポートされていないパラメータ型です: " & TypeName(params(i))
            End Select
        End If
    Next i
End Sub

Private Sub LogWarning(ByVal message As String, ByVal procedureName As String)
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrDatabaseWarning
        .Description = message
        .Category = ECDatabase
        .Source = MODULE_NAME
        .ProcedureName = procedureName
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
End Sub

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
        InitializeModule mConfig
    End Sub
    
    ''' <summary>
    ''' パフォーマンスモニターの参照を取得（テスト用）
    ''' </summary>
    Private Function GetPerformanceMonitor() As clsPerformanceMonitor
        Set GetPerformanceMonitor = mPerformanceMonitor
    End Function
    
    ''' <summary>
    ''' コネクションプールの参照を取得（テスト用）
    ''' </summary>
    Private Function GetConnectionPool() As ConnectionPool
        Set GetConnectionPool = mConnectionPool
    End Function
#End If