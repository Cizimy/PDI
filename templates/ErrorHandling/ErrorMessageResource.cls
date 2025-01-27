VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorMessageResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("エラーメッセージの国際化リソースを管理するクラス")

' プライベート変数
Private mMessages As Object ' Dictionary
Private mCurrentLocale As String
Private mDefaultLocale As String

Private Sub Class_Initialize()
    Set mMessages = CreateObject("Scripting.Dictionary")
    mDefaultLocale = "ja-JP"
    mCurrentLocale = mDefaultLocale
    InitializeMessages
End Sub

Private Sub Class_Terminate()
    Set mMessages = Nothing
End Sub

'@Description("メッセージリソースを初期化します")
Private Sub InitializeMessages()
    ' 日本語メッセージ
    Dim jaMessages As Object
    Set jaMessages = CreateObject("Scripting.Dictionary")
    
    With jaMessages
        .Add ecInvalidInput, "不正な入力です"
        .Add ecDatabaseError, "データベースエラーが発生しました"
        .Add ecNetworkError, "ネットワークエラーが発生しました"
        .Add ecFileNotFound, "ファイルが見つかりません"
        .Add ecFileAccessDenied, "ファイルアクセスが拒否されました"
        .Add ecOutOfMemory, "メモリが不足しています"
        .Add ecTimeout, "タイムアウトが発生しました"
        .Add ecDivisionByZero, "ゼロ除算が発生しました"
        .Add ecUnknownError, "不明なエラーが発生しました"
    End With
    mMessages.Add "ja-JP", jaMessages
    
    ' 英語メッセージ
    Dim enMessages As Object
    Set enMessages = CreateObject("Scripting.Dictionary")
    
    With enMessages
        .Add ecInvalidInput, "Invalid input"
        .Add ecDatabaseError, "Database error occurred"
        .Add ecNetworkError, "Network error occurred"
        .Add ecFileNotFound, "File not found"
        .Add ecFileAccessDenied, "File access denied"
        .Add ecOutOfMemory, "Out of memory"
        .Add ecTimeout, "Timeout occurred"
        .Add ecDivisionByZero, "Division by zero occurred"
        .Add ecUnknownError, "Unknown error occurred"
    End With
    mMessages.Add "en-US", enMessages
End Sub

'@Description("現在のロケールを設定します")
Public Property Let CurrentLocale(ByVal locale As String)
    If mMessages.Exists(locale) Then
        mCurrentLocale = locale
    Else
        mCurrentLocale = mDefaultLocale
    End If
End Property

'@Description("現在のロケールを取得します")
Public Property Get CurrentLocale() As String
    CurrentLocale = mCurrentLocale
End Property

'@Description("エラーメッセージを取得します")
Public Function GetMessage(ByVal errorCode As ErrorCode, Optional ByVal locale As String = "") As String
    On Error GoTo ErrorHandler
    
    Dim targetLocale As String
    targetLocale = IIf(locale = "", mCurrentLocale, locale)
    
    If Not mMessages.Exists(targetLocale) Then
        targetLocale = mDefaultLocale
    End If
    
    If mMessages(targetLocale).Exists(errorCode) Then
        GetMessage = mMessages(targetLocale)(errorCode)
    Else
        GetMessage = mMessages(mDefaultLocale)(ecUnknownError)
    End If
    Exit Function
    
ErrorHandler:
    GetMessage = "エラーメッセージの取得に失敗しました"
End Function

'@Description("新しいロケールとメッセージを追加します")
Public Sub AddLocale(ByVal locale As String, ByVal messages As Object)
    If Not mMessages.Exists(locale) Then
        mMessages.Add locale, messages
    End If
End Sub

'@Description("指定したロケールのメッセージを更新します")
Public Sub UpdateMessage(ByVal locale As String, ByVal errorCode As ErrorCode, ByVal message As String)
    If mMessages.Exists(locale) Then
        If mMessages(locale).Exists(errorCode) Then
            mMessages(locale)(errorCode) = message
        Else
            mMessages(locale).Add errorCode, message
        End If
    End If
End Sub