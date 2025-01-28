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

' メッセージテンプレート情報を格納する型
Private Type MessageTemplate
    Template As String           ' メッセージテンプレート
    LocaleID As String          ' ロケールID
    Category As ErrorCategory   ' エラーカテゴリ
    Version As Long             ' テンプレートバージョン
    LastModified As Date        ' 最終更新日時
    CustomFormatter As IMessageFormatter ' カスタムフォーマッター
End Type

' プライベート変数
Private mMessages As Object ' Dictionary<String, Dictionary<ErrorCode, MessageTemplate>>
Private mCurrentLocale As String
Private mDefaultLocale As String
Private mDefaultFormatter As IMessageFormatter
Private mTemplateCache As Object ' Dictionary
Private mCategoryMessages As Object ' Dictionary<ErrorCategory, Collection>

' イベント
Public Event MessageUpdated(ByVal locale As String, ByVal errorCode As ErrorCode)
Public Event LocaleAdded(ByVal locale As String)
Public Event TemplateRegistered(ByVal errorCode As ErrorCode, ByVal template As String)
Public Event FormatterChanged(ByVal errorCode As ErrorCode)

Private Sub Class_Initialize()
    Set mMessages = CreateObject("Scripting.Dictionary")
    Set mTemplateCache = CreateObject("Scripting.Dictionary")
    Set mCategoryMessages = CreateObject("Scripting.Dictionary")
    Set mDefaultFormatter = New DefaultMessageFormatter
    
    mDefaultLocale = "ja-JP"
    mCurrentLocale = mDefaultLocale
    
    InitializeMessages
End Sub

Private Sub Class_Terminate()
    Set mMessages = Nothing
    Set mTemplateCache = Nothing
    Set mCategoryMessages = Nothing
    Set mDefaultFormatter = Nothing
End Sub

'@Description("メッセージリソースを初期化します")
Private Sub InitializeMessages()
    ' 日本語メッセージ
    Dim jaMessages As Object
    Set jaMessages = CreateObject("Scripting.Dictionary")
    
    With jaMessages
        .Add ecInvalidInput, CreateTemplate("不正な入力です: {0}", "ja-JP", ecInputError)
        .Add ecDatabaseError, CreateTemplate("データベースエラーが発生しました: {0}", "ja-JP", ecDatabaseError)
        .Add ecNetworkError, CreateTemplate("ネットワークエラーが発生しました: {0}", "ja-JP", ecNetworkError)
        .Add ecFileNotFound, CreateTemplate("ファイル '{0}' が見つかりません", "ja-JP", ecFileIOError)
        .Add ecFileAccessDenied, CreateTemplate("ファイル '{0}' へのアクセスが拒否されました", "ja-JP", ecFileIOError)
        .Add ecOutOfMemory, CreateTemplate("メモリが不足しています。必要: {0}MB, 利用可能: {1}MB", "ja-JP", ecSystemError)
        .Add ecTimeout, CreateTemplate("タイムアウトが発生しました。制限時間: {0}秒", "ja-JP", ecSystemError)
        .Add ecDivisionByZero, CreateTemplate("ゼロ除算が発生しました", "ja-JP", ecLogicError)
        .Add ecUnknownError, CreateTemplate("不明なエラーが発生しました: {0}", "ja-JP", ecUnknownCategory)
    End With
    mMessages.Add "ja-JP", jaMessages
    
    ' 英語メッセージ
    Dim enMessages As Object
    Set enMessages = CreateObject("Scripting.Dictionary")
    
    With enMessages
        .Add ecInvalidInput, CreateTemplate("Invalid input: {0}", "en-US", ecInputError)
        .Add ecDatabaseError, CreateTemplate("Database error occurred: {0}", "en-US", ecDatabaseError)
        .Add ecNetworkError, CreateTemplate("Network error occurred: {0}", "en-US", ecNetworkError)
        .Add ecFileNotFound, CreateTemplate("File '{0}' not found", "en-US", ecFileIOError)
        .Add ecFileAccessDenied, CreateTemplate("Access denied to file '{0}'", "en-US", ecFileIOError)
        .Add ecOutOfMemory, CreateTemplate("Out of memory. Required: {0}MB, Available: {1}MB", "en-US", ecSystemError)
        .Add ecTimeout, CreateTemplate("Operation timed out. Limit: {0} seconds", "en-US", ecSystemError)
        .Add ecDivisionByZero, CreateTemplate("Division by zero occurred", "en-US", ecLogicError)
        .Add ecUnknownError, CreateTemplate("Unknown error occurred: {0}", "en-US", ecUnknownCategory)
    End With
    mMessages.Add "en-US", enMessages
    
    ' カテゴリ別メッセージの初期化
    InitializeCategoryMessages
End Sub

'@Description("カテゴリ別メッセージを初期化します")
Private Sub InitializeCategoryMessages()
    Dim category As ErrorCategory
    For category = ecInputError To ecUnknownCategory
        Set mCategoryMessages(category) = New Collection
    Next category
    
    ' メッセージをカテゴリごとに分類
    Dim locale As Variant
    Dim errorCode As Variant
    Dim template As MessageTemplate
    
    For Each locale In mMessages.Keys
        For Each errorCode In mMessages(locale).Keys
            template = mMessages(locale)(errorCode)
            mCategoryMessages(template.Category).Add errorCode
        Next errorCode
    Next locale
End Sub

'@Description("メッセージテンプレートを作成します")
Private Function CreateTemplate(ByVal template As String, ByVal locale As String, _
                              ByVal category As ErrorCategory) As MessageTemplate
    Dim result As MessageTemplate
    
    With result
        .Template = template
        .LocaleID = locale
        .Category = category
        .Version = 1
        .LastModified = Now
        Set .CustomFormatter = Nothing ' デフォルトフォーマッターを使用
    End With
    
    CreateTemplate = result
End Function

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
Public Function GetMessage(ByVal errorCode As ErrorCode, Optional ByVal locale As String = "", _
                         ParamArray args() As Variant) As String
    On Error GoTo ErrorHandler
    
    Dim targetLocale As String
    targetLocale = IIf(locale = "", mCurrentLocale, locale)
    
    If Not mMessages.Exists(targetLocale) Then
        targetLocale = mDefaultLocale
    End If
    
    If mMessages(targetLocale).Exists(errorCode) Then
        Dim template As MessageTemplate
        template = mMessages(targetLocale)(errorCode)
        
        ' カスタムフォーマッターがある場合はそれを使用
        If Not template.CustomFormatter Is Nothing Then
            GetMessage = template.CustomFormatter.FormatMessage(template.Template, args)
        Else
            GetMessage = mDefaultFormatter.FormatMessage(template.Template, args)
        End If
    Else
        GetMessage = FormatUnknownError(errorCode, args)
    End If
    Exit Function
    
ErrorHandler:
    GetMessage = "エラーメッセージの取得に失敗しました"
End Function

'@Description("不明なエラーのメッセージをフォーマットします")
Private Function FormatUnknownError(ByVal errorCode As ErrorCode, ByRef args() As Variant) As String
    Dim template As MessageTemplate
    template = mMessages(mDefaultLocale)(ecUnknownError)
    
    If Not template.CustomFormatter Is Nothing Then
        FormatUnknownError = template.CustomFormatter.FormatMessage(template.Template, Array(CStr(errorCode)))
    Else
        FormatUnknownError = mDefaultFormatter.FormatMessage(template.Template, Array(CStr(errorCode)))
    End If
End Function

'@Description("新しいロケールとメッセージを追加します")
Public Sub AddLocale(ByVal locale As String, ByVal messages As Object)
    If Not mMessages.Exists(locale) Then
        ValidateMessages messages
        mMessages.Add locale, messages
        RaiseEvent LocaleAdded(locale)
    End If
End Sub

'@Description("メッセージの検証を行います")
Private Sub ValidateMessages(ByVal messages As Object)
    If messages Is Nothing Then
        Err.Raise vbObjectError + 1, "ErrorMessageResource", "Messages cannot be null"
    End If
    
    If Not TypeOf messages Is Object Then
        Err.Raise vbObjectError + 2, "ErrorMessageResource", "Messages must be a Dictionary"
    End If
End Sub

'@Description("指定したロケールのメッセージを更新します")
Public Sub UpdateMessage(ByVal locale As String, ByVal errorCode As ErrorCode, _
                        ByVal template As String, Optional ByVal formatter As IMessageFormatter = Nothing)
    If mMessages.Exists(locale) Then
        If mMessages(locale).Exists(errorCode) Then
            Dim msgTemplate As MessageTemplate
            msgTemplate = mMessages(locale)(errorCode)
            
            ' テンプレートの更新
            msgTemplate.Template = template
            msgTemplate.Version = msgTemplate.Version + 1
            msgTemplate.LastModified = Now
            
            ' フォーマッターの更新
            If Not formatter Is Nothing Then
                Set msgTemplate.CustomFormatter = formatter
                RaiseEvent FormatterChanged(errorCode)
            End If
            
            mMessages(locale)(errorCode) = msgTemplate
            ClearTemplateCache errorCode
            
            RaiseEvent MessageUpdated(locale, errorCode)
        Else
            mMessages(locale).Add errorCode, CreateTemplate(template, locale, GetDefaultCategory(errorCode))
        End If
    End If
End Sub

'@Description("テンプレートキャッシュをクリアします")
Private Sub ClearTemplateCache(ByVal errorCode As ErrorCode)
    Dim key As Variant
    For Each key In mTemplateCache.Keys
        If InStr(key, CStr(errorCode)) > 0 Then
            mTemplateCache.Remove key
        End If
    Next key
End Sub

'@Description("デフォルトのエラーカテゴリを取得します")
Private Function GetDefaultCategory(ByVal errorCode As ErrorCode) As ErrorCategory
    Select Case errorCode
        Case ecInvalidInput: GetDefaultCategory = ecInputError
        Case ecDatabaseError: GetDefaultCategory = ecDatabaseError
        Case ecNetworkError: GetDefaultCategory = ecNetworkError
        Case ecFileNotFound, ecFileAccessDenied: GetDefaultCategory = ecFileIOError
        Case Else: GetDefaultCategory = ecUnknownCategory
    End Select
End Function

'@Description("カテゴリ別のエラーコードを取得します")
Public Function GetErrorCodesByCategory(ByVal category As ErrorCategory) As Collection
    If mCategoryMessages.Exists(category) Then
        Set GetErrorCodesByCategory = mCategoryMessages(category)
    Else
        Set GetErrorCodesByCategory = New Collection
    End If
End Function

'@Description("カスタムフォーマッターを設定します")
Public Sub SetCustomFormatter(ByVal errorCode As ErrorCode, ByVal formatter As IMessageFormatter)
    Dim locale As Variant
    For Each locale In mMessages.Keys
        If mMessages(locale).Exists(errorCode) Then
            Dim template As MessageTemplate
            template = mMessages(locale)(errorCode)
            Set template.CustomFormatter = formatter
            mMessages(locale)(errorCode) = template
        End If
    Next locale
    
    RaiseEvent FormatterChanged(errorCode)
End Sub

'@Description("メッセージテンプレートを登録します")
Public Sub RegisterTemplate(ByVal errorCode As ErrorCode, ByVal template As String, _
                          Optional ByVal category As ErrorCategory = ecUnknownCategory)
    Dim locale As Variant
    For Each locale In mMessages.Keys
        UpdateMessage CStr(locale), errorCode, template
    Next locale
    
    RaiseEvent TemplateRegistered(errorCode, template)
End Sub