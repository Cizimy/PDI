Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "modCommon"

' ======================
' アプリケーション定数
' ======================
Public Const APPLICATION_NAME As String = "MyApp"
Public Const APPLICATION_VERSION As String = "1.0.0"
Public Const DEFAULT_LOG_FILE As String = "app.log"
Public Const DEFAULT_DATE_FORMAT As String = "yyyy/MM/dd"
Public Const DEFAULT_DATETIME_FORMAT As String = "yyyy/MM/dd HH:nn:ss"
Public Const MAX_RETRY_COUNT As Integer = 3

' ======================
' ログ関連の定義
' ======================
Public Enum LogLevel
    LevelDebug
    LevelInfo
    LevelWarning
    LevelError
    LevelFatal
End Enum

Public Enum LogDestination
    DestNone
    DestFile
    DestDatabase
    DestEventLog
    DestConsole
    DestEmail
End Enum

' ======================
' セキュリティレベル
' ======================
Public Enum SecurityLevel
    LevelLow = 1
    LevelMedium = 2
    LevelHigh = 3
    LevelExtreme = 4
End Enum

' ======================
' ファイルアクセスモード
' ======================
Public Enum FileAccessMode
    ModeReadOnly = 1
    ModeReadWrite = 2
    ModeAppend = 3
    ModeExclusive = 4
End Enum

' ======================
' 型定義
' ======================
Public Type ErrorInfo
    Code As ErrorCode
    Category As ErrorCodeCategory
    Description As String
    Source As String
    ProcedureName As String
    StackTrace As String
    OccurredAt As Date
    AdditionalInfo As String
End Type

Public Type FileInfo
    Name As String
    Path As String
    Size As Long
    Created As Date
    LastModified As Date
    FileType As String
    Attributes As Long
End Type

' ======================
' モジュール変数
' ======================
Private mPerformanceMonitor As clsPerformanceMonitor
Private mIsInitialized As Boolean

' ======================
' 初期化・終了処理
' ======================
Public Property Get IsInitialized() As Boolean
    IsInitialized = mIsInitialized
End Property

Public Sub InitializeModule()
    If mIsInitialized Then Exit Sub
    
    ' スタックトレースモジュールの初期化
    modStackTrace.InitializeModule
    
    ' パフォーマンスモニターの初期化
    Set mPerformanceMonitor = New clsPerformanceMonitor
    
    ' 設定の初期化
    modConfig.InitializeModule
    
    mIsInitialized = True
End Sub

Public Sub TerminateModule()
    If Not mIsInitialized Then Exit Sub
    
    ' 設定の終了処理
    modConfig.TerminateModule
    
    ' パフォーマンスモニターの解放
    Set mPerformanceMonitor = Nothing
    
    mIsInitialized = False
End Sub

' ======================
' エラーハンドリング
' ======================
Public Sub HandleError(ByRef errInfo As ErrorInfo)
    ' エラー情報の補完
    With errInfo
        If .OccurredAt = #12:00:00 AM# Then .OccurredAt = Now
        If .Category = 0 Then .Category = modErrorCodes.GetErrorCategory(.Code)
        If Len(.StackTrace) = 0 Then .StackTrace = modStackTrace.GetStackTrace()
    End With
    
    ' エラーハンドラの取得
    Dim handler As IErrorHandler
    Set handler = GetErrorHandler(errInfo.Code)
    
    ' パフォーマンスモニタリング（エラー発生時の状態記録）
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.StartMeasurement "ErrorHandling_" & errInfo.Code
    End If
    
    ' エラーハンドラによる処理
    Dim proceed As Boolean
    proceed = handler.HandleError(errInfo)
    
    ' エラー処理の結果に基づいて処理を継続するかどうかを判断
    If Not proceed Then
        Err.Raise errInfo.Code, errInfo.Source, errInfo.Description
    End If
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "ErrorHandling_" & errInfo.Code
    End If
End Sub

' ======================
' テストサポート機能
' ======================
#If DEBUG Then
    Public Sub ResetModule()
        TerminateModule
        InitializeModule
    End Sub
    
    Public Function GetPerformanceMonitor() As clsPerformanceMonitor
        Set GetPerformanceMonitor = mPerformanceMonitor
    End Function
#End If