VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MsgBoxNotifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredefinedId = False
Attribute VB_Exposed = False
Option Explicit

' ======================
' モジュール情報
' ======================
Private Const MODULE_NAME As String = "MsgBoxNotifier"

' ======================
' インターフェース実装
' ======================
Implements IUserNotifier

' ======================
' プライベート変数
' ======================
Private mNotificationStyle As VbMsgBoxStyle
Private mDefaultTitle As String
Private mLock As clsLock
Private mPerformanceMonitor As clsPerformanceMonitor

' ======================
' 型定義
' ======================
Private Type NotificationContext
    ErrorDetail As ErrorInfo
    Buttons As VbMsgBoxStyle
    Title As String
    Message As String
    Icon As VbMsgBoxStyle
    IsLocked As Boolean
    Result As VbMsgBoxResult
    ExecutionTime As Double
End Type

' ======================
' 初期化・終了処理
' ======================
Private Sub Class_Initialize()
    mNotificationStyle = vbOKOnly Or vbInformation
    mDefaultTitle = "エラー"
    Set mLock = New clsLock
    Set mPerformanceMonitor = New clsPerformanceMonitor
End Sub

Private Sub Class_Terminate()
    Set mLock = Nothing
    Set mPerformanceMonitor = Nothing
End Sub

' ======================
' IUserNotifier実装
' ======================
Private Function IUserNotifier_Notify(ByRef errorDetail As ErrorInfo, _
                                   Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
                                   Optional ByVal title As String = "エラー") As VbMsgBoxResult
    ' パフォーマンス計測開始
    mPerformanceMonitor.StartMeasurement "Notify"
    
    ' 通知コンテキストの初期化
    Dim context As NotificationContext
    context.ErrorDetail = errorDetail
    context.Buttons = buttons
    context.Title = IIf(title = "", mDefaultTitle, title)

    ' ロック取得
    On Error Resume Next
    mLock.AcquireLock
    context.IsLocked = (Err.Number = 0)
    On Error GoTo 0

    ' 通知処理の実行
    If TryPrepareNotification(context) Then
        context.Result = ShowNotification(context)
    Else
        ' 通知準備に失敗した場合のフォールバック
        context.Result = vbAbort
    End If

    ' クリーンアップ処理
    If context.IsLocked Then
        mLock.ReleaseLock
    End If

    ' パフォーマンス計測終了
    mPerformanceMonitor.EndMeasurement "Notify"
    
    IUserNotifier_Notify = context.Result
End Function

Private Property Let IUserNotifier_NotificationStyle(ByVal style As VbMsgBoxStyle)
    mNotificationStyle = style
End Property

Private Property Get IUserNotifier_NotificationStyle() As VbMsgBoxStyle
    IUserNotifier_NotificationStyle = mNotificationStyle
End Property

Private Property Let IUserNotifier_DefaultTitle(ByVal title As String)
    mDefaultTitle = title
End Property

Private Property Get IUserNotifier_DefaultTitle() As String
    IUserNotifier_DefaultTitle = mDefaultTitle
End Property

' ======================
' プライベートメソッド
' ======================
Private Function TryPrepareNotification(ByRef context As NotificationContext) As Boolean
    On Error GoTo ErrorHandler
    
    mPerformanceMonitor.StartMeasurement "PrepareNotification"
    
    ' アイコンの設定
    context.Icon = GetIconForErrorCategory(context.ErrorDetail.Category)
    
    ' メッセージの組み立て
    context.Message = BuildErrorMessage(context.ErrorDetail)
    
    mPerformanceMonitor.EndMeasurement "PrepareNotification"
    TryPrepareNotification = True
    Exit Function
    
ErrorHandler:
    ' エラー情報のログ記録
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "通知の準備中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "TryPrepareNotification"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "PrepareNotification"
    End If
    
    TryPrepareNotification = False
End Function

Private Function ShowNotification(ByRef context As NotificationContext) As VbMsgBoxResult
    On Error GoTo ErrorHandler
    
    mPerformanceMonitor.StartMeasurement "ShowNotification"
    
    ' メッセージボックスの表示（NotificationStyleを適用）
    ShowNotification = MsgBox(context.Message, _
                            mNotificationStyle Or context.Buttons Or context.Icon, _
                            context.Title)
    
    mPerformanceMonitor.EndMeasurement "ShowNotification"
    Exit Function
    
ErrorHandler:
    Dim errDetail As ErrorInfo
    With errDetail
        .Code = ErrUnexpected
        .Description = "メッセージボックスの表示中にエラーが発生しました: " & Err.Description
        .Category = ECGeneral
        .Source = MODULE_NAME
        .ProcedureName = "ShowNotification"
        .StackTrace = modStackTrace.GetStackTrace()
        .OccurredAt = Now
    End With
    modError.HandleError errDetail
    
    If Not mPerformanceMonitor Is Nothing Then
        mPerformanceMonitor.EndMeasurement "ShowNotification"
    End If
    
    ShowNotification = vbAbort
End Function

Private Function GetIconForErrorCategory(ByVal category As ErrorCodeCategory) As VbMsgBoxStyle
    Select Case category
        Case ECSystem, ECSecurity
            GetIconForErrorCategory = vbCritical
        Case ECDatabase, ECNetwork
            GetIconForErrorCategory = vbExclamation
        Case Else
            GetIconForErrorCategory = vbInformation
    End Select
End Function

Private Function BuildErrorMessage(ByRef errorDetail As ErrorInfo) As String
    Dim msg As String
    
    msg = "エラーが発生しました。" & vbCrLf & vbCrLf & _
          "説明: " & errorDetail.Description & vbCrLf & _
          "発生場所: " & errorDetail.Source & "." & errorDetail.ProcedureName
    
    ' スタックトレースが存在する場合は追加
    If Len(errorDetail.StackTrace) > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "スタックトレース:" & vbCrLf & _
              errorDetail.StackTrace
    End If
    
    BuildErrorMessage = msg
End Function

' ======================
' テストサポート機能
' ======================
#If DEBUG Then
    ' === 通知テスト ===
    Public Sub TestNotification()
        ' テスト用エラー情報の作成
        Dim testError As ErrorInfo
        With testError
            .Code = ErrUnexpected
            .Description = "テスト通知"
            .Category = ECGeneral
            .Source = MODULE_NAME
            .ProcedureName = "TestNotification"
            .StackTrace = "テストスタックトレース"
            .OccurredAt = Now
        End With
        
        ' 各種スタイルでのテスト
        TestNotificationStyle testError, vbOKOnly, "OKのみ"
        TestNotificationStyle testError, vbOKCancel, "OK/キャンセル"
        TestNotificationStyle testError, vbYesNo, "はい/いいえ"
    End Sub
    
    Private Sub TestNotificationStyle(ByRef testError As ErrorInfo, _
                                    ByVal style As VbMsgBoxStyle, _
                                    ByVal description As String)
        mPerformanceMonitor.StartMeasurement "NotificationTest_" & description
        
        Dim result As VbMsgBoxResult
        result = IUserNotifier_Notify(testError, style, "テスト: " & description)
        
        mPerformanceMonitor.EndMeasurement "NotificationTest_" & description
        
        Debug.Print description & " テスト結果:"
        Debug.Print "  - 実行時間: " & _
                   mPerformanceMonitor.GetMeasurement("NotificationTest_" & description)
        Debug.Print "  - ユーザー応答: " & result
    End Sub
    
    ' === リソース管理テスト ===
    Public Sub TestResourceManagement()
        Dim lockCountBefore As Long
        lockCountBefore = GetActiveLockCount()
        
        Dim testError As ErrorInfo
        With testError
            .Code = ErrUnexpected
            .Description = "リソース管理テスト"
            .Category = ECGeneral
            .Source = MODULE_NAME
            .ProcedureName = "TestResourceManagement"
            .StackTrace = ""
            .OccurredAt = Now
        End With
        
        mPerformanceMonitor.StartMeasurement "ResourceTest"
        
        On Error Resume Next
        IUserNotifier_Notify testError, vbOKOnly, "リソーステスト"
        
        mPerformanceMonitor.EndMeasurement "ResourceTest"
        
        Dim lockCountAfter As Long
        lockCountAfter = GetActiveLockCount()
        
        Debug.Print "リソース管理テスト結果:"
        Debug.Print "  - 実行時間: " & _
                   mPerformanceMonitor.GetMeasurement("ResourceTest")
        
        If lockCountBefore <> lockCountAfter Then
            Debug.Print "  警告: リソースリークの可能性があります"
            Debug.Print "  - ロック数 Before: " & lockCountBefore
            Debug.Print "  - ロック数 After: " & lockCountAfter
        End If
    End Sub
    
    ' === パフォーマンスレポート ===
    Public Function GetPerformanceReport() As String
        If Not mPerformanceMonitor Is Nothing Then
            GetPerformanceReport = mPerformanceMonitor.GetAllMeasurements()
        Else
            GetPerformanceReport = "パフォーマンスモニターが初期化されていません。"
        End If
    End Function
    
    ' === 内部状態取得 ===
    Private Function GetActiveLockCount() As Long
        Dim result As Long
        result = 0
        
        If Not mLock Is Nothing Then
            If mLock.IsLocked Then
                result = result + 1
            End If
        End If
        
        GetActiveLockCount = result
    End Function
#End If