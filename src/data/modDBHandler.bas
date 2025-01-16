Attribute VB_Name = "modDBHandler"
Option Explicit

'*******************************************************************************
' モジュール: modDBHandler
' 目的：     データベース操作の中央管理
' 作成日：   2025/01/17
'*******************************************************************************

Private Type DBConnection
    ConnectionString As String
    Connection As Object      ' ADODB.Connection
    IsConnected As Boolean
    LastError As String
End Type

Private mConnection As DBConnection

'*******************************************************************************
' 目的：    データベース接続の初期化
' 引数：    なし
' 戻り値：  なし
'*******************************************************************************
Public Sub Initialize()
    On Error GoTo ErrorHandler
    
    ' 設定からデータベース接続情報を取得
    Dim server As String
    Dim database As String
    Dim username As String
    Dim password As String
    
    server = modConfigManager.GetValue("database", "server", "")
    database = modConfigManager.GetValue("database", "database", "")
    username = modConfigManager.GetValue("database", "username", "")
    password = modConfigManager.GetValue("database", "password", "")
    
    ' 接続文字列の構築
    mConnection.ConnectionString = "Provider=SQLOLEDB;" & _
                                 "Data Source=" & server & ";" & _
                                 "Initial Catalog=" & database & ";" & _
                                 "User ID=" & username & ";" & _
                                 "Password=" & password
    
    modLogger.Info "データベース設定を初期化しました。", "DBHandler.Initialize"
    Exit Sub
    
ErrorHandler:
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "DBHandler.Initialize", _
                               etDatabase
End Sub

'*******************************************************************************
' 目的：    データベースへの接続
' 引数：    なし
' 戻り値：  成功時 True
'*******************************************************************************
Public Function Connect() As Boolean
    On Error GoTo ErrorHandler
    
    ' 既に接続されている場合は何もしない
    If mConnection.IsConnected Then
        Connect = True
        Exit Function
    End If
    
    ' 接続オブジェクトの作成と接続
    Set mConnection.Connection = CreateObject("ADODB.Connection")
    With mConnection.Connection
        .ConnectionString = mConnection.ConnectionString
        .ConnectionTimeout = 30
        .Open
    End With
    
    mConnection.IsConnected = True
    Connect = True
    
    modLogger.Info "データベースに接続しました。", "DBHandler.Connect"
    Exit Function
    
ErrorHandler:
    mConnection.LastError = Err.Description
    mConnection.IsConnected = False
    Connect = False
    
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "DBHandler.Connect", _
                               etDatabase
End Function

'*******************************************************************************
' 目的：    データベース接続の切断
' 引数：    なし
' 戻り値：  なし
'*******************************************************************************
Public Sub Disconnect()
    On Error GoTo ErrorHandler
    
    If mConnection.IsConnected Then
        mConnection.Connection.Close
        Set mConnection.Connection = Nothing
        mConnection.IsConnected = False
        modLogger.Info "データベース接続を切断しました。", "DBHandler.Disconnect"
    End If
    
    Exit Sub
    
ErrorHandler:
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "DBHandler.Disconnect", _
                               etDatabase
End Sub

'*******************************************************************************
' 目的：    SQLクエリの実行（読み取り）
' 引数：    sql - 実行するSQL文
'           params - SQLパラメータ（オプション）
' 戻り値：  レコードセット
'*******************************************************************************
Public Function ExecuteQuery(ByVal sql As String, _
                           Optional ByRef params As Variant) As Object
    On Error GoTo ErrorHandler
    
    Dim cmd As Object    ' ADODB.Command
    Dim rs As Object     ' ADODB.Recordset
    
    ' 接続確認
    If Not mConnection.IsConnected Then
        If Not Connect Then
            Exit Function
        End If
    End If
    
    ' コマンドオブジェクトの作成
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        Set .ActiveConnection = mConnection.Connection
        .CommandText = sql
        .CommandType = 1  ' adCmdText
        
        ' パラメータの設定
        If Not IsMissing(params) Then
            If IsArray(params) Then
                Dim i As Long
                For i = LBound(params) To UBound(params)
                    .Parameters.Append .CreateParameter("p" & i, _
                                                     GetParamType(params(i)), _
                                                     1, _  ' adParamInput
                                                     Len(params(i)), _
                                                     params(i))
                Next i
            End If
        End If
    End With
    
    ' クエリの実行
    Set rs = cmd.Execute
    Set ExecuteQuery = rs
    
    modLogger.Debug "SQLクエリを実行しました: " & sql, "DBHandler.ExecuteQuery"
    Exit Function
    
ErrorHandler:
    mConnection.LastError = Err.Description
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "DBHandler.ExecuteQuery", _
                               etDatabase, _
                               "SQL: " & sql
End Function

'*******************************************************************************
' 目的：    SQLコマンドの実行（更新系）
' 引数：    sql - 実行するSQL文
'           params - SQLパラメータ（オプション）
' 戻り値：  影響を受けた行数
'*******************************************************************************
Public Function ExecuteCommand(ByVal sql As String, _
                             Optional ByRef params As Variant) As Long
    On Error GoTo ErrorHandler
    
    Dim cmd As Object    ' ADODB.Command
    Dim affectedRows As Long
    
    ' 接続確認
    If Not mConnection.IsConnected Then
        If Not Connect Then
            Exit Function
        End If
    End If
    
    ' コマンドオブジェクトの作成
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        Set .ActiveConnection = mConnection.Connection
        .CommandText = sql
        .CommandType = 1  ' adCmdText
        
        ' パラメータの設定
        If Not IsMissing(params) Then
            If IsArray(params) Then
                Dim i As Long
                For i = LBound(params) To UBound(params)
                    .Parameters.Append .CreateParameter("p" & i, _
                                                     GetParamType(params(i)), _
                                                     1, _  ' adParamInput
                                                     Len(params(i)), _
                                                     params(i))
                Next i
            End If
        End If
    End With
    
    ' コマンドの実行
    affectedRows = cmd.Execute
    ExecuteCommand = affectedRows
    
    modLogger.Debug "SQLコマンドを実行しました: " & sql & _
                   " (影響行数: " & affectedRows & ")", _
                   "DBHandler.ExecuteCommand"
    Exit Function
    
ErrorHandler:
    mConnection.LastError = Err.Description
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "DBHandler.ExecuteCommand", _
                               etDatabase, _
                               "SQL: " & sql
End Function

'*******************************************************************************
' 目的：    トランザクションの開始
' 引数：    なし
' 戻り値：  なし
'*******************************************************************************
Public Sub BeginTransaction()
    On Error GoTo ErrorHandler
    
    If mConnection.IsConnected Then
        mConnection.Connection.BeginTrans
        modLogger.Debug "トランザクションを開始しました。", "DBHandler.BeginTransaction"
    End If
    
    Exit Sub
    
ErrorHandler:
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "DBHandler.BeginTransaction", _
                               etDatabase
End Sub

'*******************************************************************************
' 目的：    トランザクションのコミット
' 引数：    なし
' 戻り値：  なし
'*******************************************************************************
Public Sub CommitTransaction()
    On Error GoTo ErrorHandler
    
    If mConnection.IsConnected Then
        mConnection.Connection.CommitTrans
        modLogger.Debug "トランザクションをコミットしました。", "DBHandler.CommitTransaction"
    End If
    
    Exit Sub
    
ErrorHandler:
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "DBHandler.CommitTransaction", _
                               etDatabase
End Sub

'*******************************************************************************
' 目的：    トランザクションのロールバック
' 引数：    なし
' 戻り値：  なし
'*******************************************************************************
Public Sub RollbackTransaction()
    On Error GoTo ErrorHandler
    
    If mConnection.IsConnected Then
        mConnection.Connection.RollbackTrans
        modLogger.Debug "トランザクションをロールバックしました。", "DBHandler.RollbackTransaction"
    End If
    
    Exit Sub
    
ErrorHandler:
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "DBHandler.RollbackTransaction", _
                               etDatabase
End Sub

'*******************************************************************************
' 目的：    最後のエラーメッセージの取得
' 引数：    なし
' 戻り値：  エラーメッセージ
'*******************************************************************************
Public Function GetLastError() As String
    GetLastError = mConnection.LastError
End Function

'*******************************************************************************
' 目的：    接続状態の確認
' 引数：    なし
' 戻り値：  接続されている場合True
'*******************************************************************************
Public Function IsConnected() As Boolean
    IsConnected = mConnection.IsConnected
End Function

'*******************************************************************************
' 目的：    パラメータの型を判定
' 引数：    value - パラメータ値
' 戻り値：  ADODBのパラメータ型
'*******************************************************************************
Private Function GetParamType(ByVal value As Variant) As Integer
    ' ADODBのデータ型定数
    ' adEmpty = 0, adSmallInt = 2, adInteger = 3, adSingle = 4, adDouble = 5,
    ' adCurrency = 6, adDate = 7, adBSTR = 8, adIDispatch = 9, adError = 10,
    ' adBoolean = 11, adVariant = 12, adIUnknown = 13, adDecimal = 14,
    ' adTinyInt = 16, adUnsignedTinyInt = 17, adUnsignedSmallInt = 18,
    ' adUnsignedInt = 19, adBigInt = 20, adUnsignedBigInt = 21, adVarChar = 200
    
    Select Case VarType(value)
        Case vbInteger
            GetParamType = 3    ' adInteger
        Case vbLong
            GetParamType = 3    ' adInteger
        Case vbSingle
            GetParamType = 4    ' adSingle
        Case vbDouble
            GetParamType = 5    ' adDouble
        Case vbCurrency
            GetParamType = 6    ' adCurrency
        Case vbDate
            GetParamType = 7    ' adDate
        Case vbString
            GetParamType = 200  ' adVarChar
        Case vbBoolean
            GetParamType = 11   ' adBoolean
        Case Else
            GetParamType = 12   ' adVariant
    End Select
End Function
