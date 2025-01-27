VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorHandlerManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("エラーハンドラーの管理を行うクラス")

'==========================
' 定数
'==========================
Private Const MODULE_NAME As String = "ErrorHandlerManager"

'==========================
' メンバ変数
'==========================
Private mHandlers As Collection
Private mLock As clsLock
Private mIsInitialized As Boolean

'==========================
' 初期化・終了処理
'==========================
Private Sub Class_Initialize()
    Set mHandlers = New Collection
    Set mLock = New clsLock
    InitializeManager
End Sub

Private Sub Class_Terminate()
    Set mHandlers = Nothing
    Set mLock = Nothing
    mIsInitialized = False
End Sub

'==========================
' パブリックメソッド
'==========================

''' <summary>
''' エラーハンドラーマネージャーを初期化します
''' </summary>
Public Sub InitializeManager()
    If mIsInitialized Then Exit Sub
    
    RegisterDefaultHandlers
    mIsInitialized = True
End Sub

''' <summary>
''' エラーコードに対応するエラーハンドラーを登録します
''' </summary>
''' <param name="errorCode">エラーコード</param>
''' <param name="handler">エラーハンドラー</param>
Public Sub RegisterHandler(ByVal errorCode As ErrorCode, ByVal handler As IErrorHandler)
    mLock.AcquireLock
    
    On Error Resume Next
    mHandlers.Remove CStr(errorCode)
    On Error GoTo 0
    
    mHandlers.Add handler, CStr(errorCode)
    
    mLock.ReleaseLock
End Sub

''' <summary>
''' エラーコードに対応するエラーハンドラーの登録を解除します
''' </summary>
''' <param name="errorCode">エラーコード</param>
Public Sub UnregisterHandler(ByVal errorCode As ErrorCode)
    mLock.AcquireLock
    
    On Error Resume Next
    mHandlers.Remove CStr(errorCode)
    On Error GoTo 0
    
    mLock.ReleaseLock
End Sub

''' <summary>
''' エラーコードに対応するエラーハンドラーを取得します
''' </summary>
''' <param name="errorCode">エラーコード</param>
''' <returns>エラーハンドラー。未登録の場合はデフォルトハンドラー</returns>
Public Function GetErrorHandler(ByVal errorCode As ErrorCode) As IErrorHandler
    Dim handler As IErrorHandler
    
    mLock.AcquireLock
    
    On Error Resume Next
    Set handler = mHandlers(CStr(errorCode))
    If Err.Number <> 0 Then
        ' 該当するハンドラが見つからない場合は、エラーカテゴリに基づいてデフォルトハンドラを返す
        Set handler = GetDefaultHandlerForCategory(modErrorCodes.GetErrorCategory(errorCode))
    End If
    On Error GoTo 0
    
    mLock.ReleaseLock
    
    Set GetErrorHandler = handler
End Function

'==========================
' プライベートメソッド
'==========================

''' <summary>
''' デフォルトのエラーハンドラーを登録します
''' </summary>
Private Sub RegisterDefaultHandlers()
    ' データベース接続エラー用ハンドラ
    RegisterHandler ErrDbConnectionFailed, New DatabaseConnectionErrorHandler
    
    ' ファイル不在エラー用ハンドラ
    RegisterHandler ErrFileNotFound, New FileNotFoundErrorHandler
    
    ' 入力検証エラー用ハンドラ
    RegisterHandler ErrInvalidInput, New InvalidInputErrorHandler
End Sub

''' <summary>
''' エラーカテゴリに対応するデフォルトのエラーハンドラーを取得します
''' </summary>
''' <param name="category">エラーカテゴリ</param>
''' <returns>デフォルトのエラーハンドラー</returns>
Private Function GetDefaultHandlerForCategory(ByVal category As ErrorCodeCategory) As IErrorHandler
    Select Case category
        Case ECDatabase
            Set GetDefaultHandlerForCategory = New DatabaseConnectionErrorHandler
        Case ECFileIO
            Set GetDefaultHandlerForCategory = New FileNotFoundErrorHandler
        Case Else
            Set GetDefaultHandlerForCategory = New InvalidInputErrorHandler
    End Select
End Function

'==========================
' テストサポート（開発環境用）
'==========================
#If DEBUG Then
    Public Function GetHandlerCount() As Long
        mLock.AcquireLock
        GetHandlerCount = mHandlers.Count
        mLock.ReleaseLock
    End Function
    
    Public Sub ClearHandlers()
        mLock.AcquireLock
        Set mHandlers = New Collection
        mLock.ReleaseLock
    End Sub
#End If