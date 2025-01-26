VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConnectionPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "ConnectionPool"

' ======================
' 定数定義
' ======================
Private Const MAX_POOL_SIZE As Long = 5
Private Const CONNECTION_TIMEOUT As Long = 30 ' 秒

' ======================
' メンバ変数
' ======================
Private Type ConnectionInfo
    Connection As Object
    InUse As Boolean
    LastUsed As Date
End Type

Private mConnections() As ConnectionInfo
Private mConnectionString As String
Private mPoolSize As Long
Private mLock As clsLock
Private mPerformanceMonitor As clsPerformanceMonitor
Private mConfig As IDatabaseConfig

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    ReDim mConnections(1 To MAX_POOL_SIZE)
    mPoolSize = 0
    Set mLock = New clsLock
    Set mPerformanceMonitor = New clsPerformanceMonitor
End Sub

Private Sub Class_Terminate()
    Dim i As Long
    For i = 1 To MAX_POOL_SIZE
        If Not mConnections(i).Connection Is Nothing Then
            If mConnections(i).Connection.State = 1 Then ' adStateOpen
                mConnections(i).Connection.Close
            End If
            Set mConnections(i).Connection = Nothing
        End If
    Next i
    Set mLock = Nothing
    Set mConfig = Nothing
    Set mPerformanceMonitor = Nothing
End Sub

' ======================
' パブリックメソッド
' ======================

''' <summary>
''' コネクションプールを初期化します
''' </summary>
''' <param name="config">データベース設定を提供するインターフェース</param>
Public Sub Initialize(ByVal config As IDatabaseConfig)
    Set mConfig = config
    mConnectionString = mConfig.GetConnectionString
End Sub

''' <summary>
''' コネクションを取得します
''' </summary>
''' <returns>データベース接続オブジェクト</returns>
Public Function GetConnection() As Object
    On Error GoTo ErrorHandler
    mPerformanceMonitor.StartMeasurement "Get Database Connection"
    
    mLock.AcquireLock
    
    ' 未使用のコネクションを探す
    Dim i As Long
    For i = 1 To mPoolSize
        If Not mConnections(i).InUse Then
            ' コネクションの状態をチェック
            If ValidateConnection(mConnections(i).Connection) Then
                mConnections(i).InUse = True
                mConnections(i).LastUsed = Now
                Set GetConnection = mConnections(i).Connection
                GoTo Cleanup
            Else
                ' 無効なコネクションを破棄して再作成
                Set mConnections(i).Connection = CreateNewConnection()
                If Not mConnections(i).Connection Is Nothing Then
                    mConnections(i).InUse = True
                    mConnections(i).LastUsed = Now
                    Set GetConnection = mConnections(i).Connection
                    GoTo Cleanup
                End If
            End If
        End If
    Next i
    
    ' 新しいコネクションを作成（プールサイズが上限に達していない場合）
    If mPoolSize < MAX_POOL_SIZE Then
        mPoolSize = mPoolSize + 1
        Set mConnections(mPoolSize).Connection = CreateNewConnection()
        If Not mConnections(mPoolSize).Connection Is Nothing Then
            mConnections(mPoolSize).InUse = True
            mConnections(mPoolSize).LastUsed = Now
            Set GetConnection = mConnections(mPoolSize).Connection
            GoTo Cleanup
        End If
    End If
    
    ' 利用可能なコネクションがない場合はエラー
    Err.Raise vbObjectError + 1, MODULE_NAME, "利用可能なコネクションがありません。"
    
Cleanup:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "Get Database Connection"
    Exit Function
    
ErrorHandler:
    Dim errDetail As typErrorDetail
    With errDetail
        .ErrorCode = ERR_DATABASE_CONNECTION_FAILED
        .Description = "コネクション取得中にエラーが発生しました: " & Err.Description
        .Category = ECDatabase
        .Source = MODULE_NAME
        .ProcedureName = "GetConnection"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    
    If Not mLock Is Nothing Then mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "Get Database Connection"
    Set GetConnection = Nothing
End Function

''' <summary>
''' コネクションを解放します
''' </summary>
''' <param name="connection">解放するコネクション</param>
Public Sub ReleaseConnection(ByVal connection As Object)
    If connection Is Nothing Then Exit Sub
    
    mLock.AcquireLock
    
    Dim i As Long
    For i = 1 To mPoolSize
        If mConnections(i).Connection Is connection Then
            mConnections(i).InUse = False
            mConnections(i).LastUsed = Now
            Exit For
        End If
    Next i
    
    mLock.ReleaseLock
End Sub

' ======================
' プライベートメソッド
' ======================

''' <summary>
''' 新しいデータベース接続を作成します
''' </summary>
Private Function CreateNewConnection() As Object
    On Error GoTo ErrorHandler
    mPerformanceMonitor.StartMeasurement "Create Database Connection"
    
    mLock.AcquireLock
    
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = mConnectionString
    conn.ConnectionTimeout = CONNECTION_TIMEOUT
    
    ' データベース接続の作成
    conn.Open
    
    Set CreateNewConnection = conn
    
Cleanup:
    mLock.ReleaseLock
    mPerformanceMonitor.EndMeasurement "Create Database Connection"
    Exit Function
    
ErrorHandler:
    Set CreateNewConnection = Nothing
    Resume Cleanup
End Function

''' <summary>
''' コネクションが有効かどうかを確認します
''' </summary>
Private Function ValidateConnection(ByVal connection As Object) As Boolean
    If connection Is Nothing Then
        ValidateConnection = False
        Exit Function
    End If
    
    On Error Resume Next
    ValidateConnection = (connection.State = 1) ' adStateOpen
    If Err.Number <> 0 Then
        ValidateConnection = False
    End If
    On Error GoTo 0
End Function

' ======================
' テストサポート機能
' 警告: これらのメソッドは開発時のテスト目的でのみ使用し、
' 本番環境では使用しないでください。
' ======================
#If DEBUG Then
    ''' <summary>
    ''' コネクションプールの現在のサイズを取得（テスト用）
    ''' </summary>
    Private Property Get PoolSize() As Long
        PoolSize = mPoolSize
    End Property
    
    ''' <summary>
    ''' アクティブなコネクション数を取得（テスト用）
    ''' </summary>
    Private Property Get ActiveConnections() As Long
        Dim count As Long
        Dim i As Long
        For i = 1 To mPoolSize
            If mConnections(i).InUse Then count = count + 1
        Next i
        ActiveConnections = count
    End Property
    
    ''' <summary>
    ''' コネクションプールをクリア（テスト用）
    ''' </summary>
    Private Sub ClearPool()
        Class_Terminate
        Class_Initialize
    End Sub
#End If