VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockDatabaseConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDatabaseConfig

' モック用の内部状態
Private mConnectionString As String
Private mMaxPoolSize As Long
Private mRetryCount As Long
Private mRetryInterval As Long
Private mConnectionTimeout As Long

' メソッド呼び出し回数のカウンター
Private mGetConnectionStringCallCount As Long
Private mGetMaxPoolSizeCallCount As Long
Private mGetRetryCountCallCount As Long
Private mGetRetryIntervalCallCount As Long
Private mGetConnectionTimeoutCallCount As Long

'*******************************************************************************
' 初期化
'*******************************************************************************
Private Sub Class_Initialize()
    ' デフォルト値の設定
    mConnectionString = "Provider=SQLOLEDB;Data Source=TestServer;"
    mMaxPoolSize = 5
    mRetryCount = 3
    mRetryInterval = 1000
    mConnectionTimeout = 30000
    
    ' カウンターの初期化
    mGetConnectionStringCallCount = 0
    mGetMaxPoolSizeCallCount = 0
    mGetRetryCountCallCount = 0
    mGetRetryIntervalCallCount = 0
    mGetConnectionTimeoutCallCount = 0
End Sub

'*******************************************************************************
' IDatabaseConfig インターフェースの実装
'*******************************************************************************
Private Function IDatabaseConfig_GetConnectionString() As String
    mGetConnectionStringCallCount = mGetConnectionStringCallCount + 1
    IDatabaseConfig_GetConnectionString = mConnectionString
End Function

Private Function IDatabaseConfig_GetMaxPoolSize() As Long
    mGetMaxPoolSizeCallCount = mGetMaxPoolSizeCallCount + 1
    IDatabaseConfig_GetMaxPoolSize = mMaxPoolSize
End Function

Private Function IDatabaseConfig_GetRetryCount() As Long
    mGetRetryCountCallCount = mGetRetryCountCallCount + 1
    IDatabaseConfig_GetRetryCount = mRetryCount
End Function

Private Function IDatabaseConfig_GetRetryInterval() As Long
    mGetRetryIntervalCallCount = mGetRetryIntervalCallCount + 1
    IDatabaseConfig_GetRetryInterval = mRetryInterval
End Function

Private Function IDatabaseConfig_GetConnectionTimeout() As Long
    mGetConnectionTimeoutCallCount = mGetConnectionTimeoutCallCount + 1
    IDatabaseConfig_GetConnectionTimeout = mConnectionTimeout
End Function

'*******************************************************************************
' モック用のセッターメソッド
'*******************************************************************************
Public Sub SetConnectionString(ByVal value As String)
    mConnectionString = value
End Sub

Public Sub SetMaxPoolSize(ByVal value As Long)
    mMaxPoolSize = value
End Sub

Public Sub SetRetryCount(ByVal value As Long)
    mRetryCount = value
End Sub

Public Sub SetRetryInterval(ByVal value As Long)
    mRetryInterval = value
End Sub

Public Sub SetConnectionTimeout(ByVal value As Long)
    mConnectionTimeout = value
End Sub

'*******************************************************************************
' 呼び出し回数の取得メソッド
'*******************************************************************************
Public Property Get GetConnectionStringCallCount() As Long
    GetConnectionStringCallCount = mGetConnectionStringCallCount
End Property

Public Property Get GetMaxPoolSizeCallCount() As Long
    GetMaxPoolSizeCallCount = mGetMaxPoolSizeCallCount
End Property

Public Property Get GetRetryCountCallCount() As Long
    GetRetryCountCallCount = mGetRetryCountCallCount
End Property

Public Property Get GetRetryIntervalCallCount() As Long
    GetRetryIntervalCallCount = mGetRetryIntervalCallCount
End Property

Public Property Get GetConnectionTimeoutCallCount() As Long
    GetConnectionTimeoutCallCount = mGetConnectionTimeoutCallCount
End Property

'*******************************************************************************
' カウンターのリセット
'*******************************************************************************
Public Sub ResetCallCounts()
    mGetConnectionStringCallCount = 0
    mGetMaxPoolSizeCallCount = 0
    mGetRetryCountCallCount = 0
    mGetRetryIntervalCallCount = 0
    mGetConnectionTimeoutCallCount = 0
End Sub