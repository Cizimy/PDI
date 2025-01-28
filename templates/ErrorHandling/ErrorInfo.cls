VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("エラー情報を格納するクラス")

' エラー詳細を格納する型
Private Type ErrorDetails
    Message As String        ' 基本メッセージ
    Description As String    ' 詳細説明
    Context As Dictionary    ' 型安全な追加情報
End Type

' プライベート変数
Private mDetails As ErrorDetails
Private mErrorCode As ErrorCode
Private mStackTrace As Collection  ' スタックトレース情報
Private mTimestamp As Date        ' エラー発生時刻
Private mSeverity As ErrorSeverity ' エラーの重大度
Private mCategory As ErrorCategory ' エラーのカテゴリ
Private mStatus As ErrorStatus     ' エラーの状態
Private mRetryInfo As Dictionary   ' リトライ関連情報
Private mSourceInfo As Dictionary  ' エラー発生源情報
Private mRelatedResources As Collection ' 関連リソース
Private mTags As Collection        ' タグ情報

' プロパティ：エラーコード
Public Property Get ErrorCode() As ErrorCode
    ErrorCode = mErrorCode
End Property

Public Property Let ErrorCode(ByVal value As ErrorCode)
    mErrorCode = value
End Property

' プロパティ：メッセージ
Public Property Get Message() As String
    Message = mDetails.Message
End Property

Public Property Let Message(ByVal value As String)
    mDetails.Message = value
End Property

' プロパティ：説明
Public Property Get Description() As String
    Description = mDetails.Description
End Property

Public Property Let Description(ByVal value As String)
    mDetails.Description = value
End Property

' プロパティ：コンテキスト
Public Property Get Context() As Dictionary
    Set Context = mDetails.Context
End Property

' プロパティ：発生時刻
Public Property Get Timestamp() As Date
    Timestamp = mTimestamp
End Property

Public Property Let Timestamp(ByVal value As Date)
    mTimestamp = value
End Property

' プロパティ：重大度
Public Property Get Severity() As ErrorSeverity
    Severity = mSeverity
End Property

Public Property Let Severity(ByVal value As ErrorSeverity)
    mSeverity = value
End Property

' プロパティ：カテゴリ
Public Property Get Category() As ErrorCategory
    Category = mCategory
End Property

Public Property Let Category(ByVal value As ErrorCategory)
    mCategory = value
End Property

' プロパティ：状態
Public Property Get Status() As ErrorStatus
    Status = mStatus
End Property

Public Property Let Status(ByVal value As ErrorStatus)
    mStatus = value
End Property

' プロパティ：スタックトレース
Public Property Get StackTrace() As Collection
    Set StackTrace = mStackTrace
End Property

' プロパティ：リトライ情報
Public Property Get RetryCount() As Long
    If mRetryInfo.Exists("count") Then
        RetryCount = mRetryInfo("count")
    Else
        RetryCount = 0
    End If
End Property

Public Property Let RetryCount(ByVal value As Long)
    mRetryInfo("count") = value
End Property

Public Property Get LastRetryTimestamp() As Date
    If mRetryInfo.Exists("lastAttempt") Then
        LastRetryTimestamp = mRetryInfo("lastAttempt")
    Else
        LastRetryTimestamp = #12:00:00 AM#
    End If
End Property

Public Property Let LastRetryTimestamp(ByVal value As Date)
    mRetryInfo("lastAttempt") = value
End Property

' プロパティ：関連リソース
Public Property Get RelatedResources() As Collection
    Set RelatedResources = mRelatedResources
End Property

' プロパティ：タグ
Public Property Get Tags() As Collection
    Set Tags = mTags
End Property

' プロパティ：ソース情報
Public Property Get SourceModule() As String
    If mSourceInfo.Exists("module") Then
        SourceModule = mSourceInfo("module")
    Else
        SourceModule = ""
    End If
End Property

Public Property Let SourceModule(ByVal value As String)
    mSourceInfo("module") = value
End Property

Public Property Get SourceProcedure() As String
    If mSourceInfo.Exists("procedure") Then
        SourceProcedure = mSourceInfo("procedure")
    Else
        SourceProcedure = ""
    End If
End Property

Public Property Let SourceProcedure(ByVal value As String)
    mSourceInfo("procedure") = value
End Property

Public Property Get SourceLine() As Long
    If mSourceInfo.Exists("line") Then
        SourceLine = mSourceInfo("line")
    Else
        SourceLine = 0
    End If
End Property

Public Property Let SourceLine(ByVal value As Long)
    mSourceInfo("line") = value
End Property

Private Sub Class_Initialize()
    ' コレクションの初期化
    Set mDetails.Context = CreateObject("Scripting.Dictionary")
    Set mStackTrace = New Collection
    Set mRetryInfo = CreateObject("Scripting.Dictionary")
    Set mSourceInfo = CreateObject("Scripting.Dictionary")
    Set mRelatedResources = New Collection
    Set mTags = New Collection
    
    ' 初期値の設定
    mTimestamp = Now
    mStatus = esNew
    mRetryInfo("count") = 0
    mRetryInfo("lastAttempt") = #12:00:00 AM#
    
    ' ソース情報の初期化
    mSourceInfo("module") = ""
    mSourceInfo("procedure") = ""
    mSourceInfo("line") = 0
End Sub

Private Sub Class_Terminate()
    ' オブジェクトの解放
    Set mDetails.Context = Nothing
    Set mStackTrace = Nothing
    Set mRetryInfo = Nothing
    Set mSourceInfo = Nothing
    Set mRelatedResources = Nothing
    Set mTags = Nothing
End Sub

'@Description("エラー情報を文字列形式で取得します")
Public Function ToString() As String
    Dim result As String
    
    result = "Error Information:" & vbCrLf & _
             "----------------" & vbCrLf & _
             "Code: " & ErrorCode & vbCrLf & _
             "Message: " & Message & vbCrLf & _
             "Description: " & Description & vbCrLf & _
             "Time: " & Format(Timestamp, "yyyy/mm/dd hh:nn:ss") & vbCrLf & _
             "Module: " & SourceModule & vbCrLf & _
             "Procedure: " & SourceProcedure & vbCrLf & _
             "Line: " & SourceLine & vbCrLf & _
             "Severity: " & GetSeverityText(Severity) & vbCrLf & _
             "Category: " & GetCategoryText(Category) & vbCrLf & _
             "Status: " & GetStatusText(Status) & vbCrLf & _
             "Retry Count: " & RetryCount & vbCrLf & _
             "Last Retry: " & Format(LastRetryTimestamp, "yyyy/mm/dd hh:nn:ss") & vbCrLf & _
             "Context Keys: " & GetContextKeysText()
    
    ToString = result
End Function

'@Description("重大度の文字列表現を取得")
Private Function GetSeverityText(ByVal severity As ErrorSeverity) As String
    Select Case severity
        Case esInfo: GetSeverityText = "情報"
        Case esWarning: GetSeverityText = "警告"
        Case esError: GetSeverityText = "エラー"
        Case esCritical: GetSeverityText = "重大"
        Case esFatal: GetSeverityText = "致命的"
        Case Else: GetSeverityText = "不明"
    End Select
End Function

'@Description("カテゴリの文字列表現を取得")
Private Function GetCategoryText(ByVal category As ErrorCategory) As String
    Select Case category
        Case ecInputError: GetCategoryText = "入力エラー"
        Case ecSystemError: GetCategoryText = "システムエラー"
        Case ecDatabaseError: GetCategoryText = "データベースエラー"
        Case ecNetworkError: GetCategoryText = "ネットワークエラー"
        Case ecFileIOError: GetCategoryText = "ファイルI/Oエラー"
        Case ecSecurityError: GetCategoryText = "セキュリティエラー"
        Case ecUnknownCategory: GetCategoryText = "不明"
        Case Else: GetCategoryText = "その他"
    End Select
End Function

'@Description("状態の文字列表現を取得")
Private Function GetStatusText(ByVal status As ErrorStatus) As String
    Select Case status
        Case esNew: GetStatusText = "新規"
        Case esPending: GetStatusText = "保留中"
        Case esInProgress: GetStatusText = "処理中"
        Case esRetrying: GetStatusText = "リトライ中"
        Case esResolved: GetStatusText = "解決済み"
        Case esFailed: GetStatusText = "失敗"
        Case esIgnored: GetStatusText = "無視"
        Case esCancelled: GetStatusText = "キャンセル"
        Case Else: GetStatusText = "不明"
    End Select
End Function

'@Description("コンテキストキーの文字列表現を取得")
Private Function GetContextKeysText() As String
    If mDetails.Context Is Nothing Then
        GetContextKeysText = "なし"
        Exit Function
    End If
    
    If mDetails.Context.Count = 0 Then
        GetContextKeysText = "なし"
        Exit Function
    End If
    
    Dim keys() As String
    ReDim keys(0 To mDetails.Context.Count - 1)
    
    Dim i As Long
    i = 0
    Dim key As Variant
    For Each key In mDetails.Context.Keys
        keys(i) = CStr(key)
        i = i + 1
    Next
    
    GetContextKeysText = Join(keys, ", ")
End Function

'@Description("エラー情報をクローンします")
Public Function Clone() As ErrorInfo
    Dim newInfo As ErrorInfo
    Set newInfo = New ErrorInfo
    
    ' 基本プロパティのコピー
    With newInfo
        .ErrorCode = Me.ErrorCode
        .Message = Me.Message
        .Description = Me.Description
        .Timestamp = Me.Timestamp
        .Severity = Me.Severity
        .Category = Me.Category
        .Status = Me.Status
        .SourceModule = Me.SourceModule
        .SourceProcedure = Me.SourceProcedure
        .SourceLine = Me.SourceLine
        .RetryCount = Me.RetryCount
        .LastRetryTimestamp = Me.LastRetryTimestamp
    End With
    
    ' コンテキストのコピー
    Dim key As Variant
    For Each key In mDetails.Context.Keys
        newInfo.Context.Add key, mDetails.Context(key)
    Next
    
    ' スタックトレースのコピー
    Dim item As Variant
    For Each item In mStackTrace
        newInfo.StackTrace.Add item
    Next
    
    ' 関連リソースのコピー
    For Each item In mRelatedResources
        newInfo.RelatedResources.Add item
    Next
    
    ' タグのコピー
    For Each item In mTags
        newInfo.Tags.Add item
    Next
    
    Set Clone = newInfo
End Function

'@Description("エラー情報を検証します")
Public Function Validate() As Boolean
    ' 必須フィールドの検証
    If mErrorCode = 0 Then Exit Function
    If Len(mDetails.Message) = 0 Then Exit Function
    If mSourceInfo("module") = "" Then Exit Function
    If mSourceInfo("procedure") = "" Then Exit Function
    
    ' コンテキストの検証
    If mDetails.Context Is Nothing Then Exit Function
    
    ' 日時の検証
    If mTimestamp = #12:00:00 AM# Then Exit Function
    
    Validate = True
End Function

'@Description("エラー情報をマージします")
Public Sub MergeWith(ByVal other As ErrorInfo)
    If other Is Nothing Then Exit Sub
    
    ' 基本情報のマージ（空の場合のみ更新）
    If Len(mDetails.Description) = 0 Then mDetails.Description = other.Description
    
    ' コンテキストのマージ
    Dim key As Variant
    For Each key In other.Context.Keys
        If Not mDetails.Context.Exists(key) Then
            mDetails.Context.Add key, other.Context(key)
        End If
    Next
    
    ' スタックトレースのマージ
    Dim item As Variant
    For Each item In other.StackTrace
        If Not Contains(mStackTrace, item) Then
            mStackTrace.Add item
        End If
    Next
    
    ' 関連リソースのマージ
    For Each item In other.RelatedResources
        If Not Contains(mRelatedResources, item) Then
            mRelatedResources.Add item
        End If
    Next
    
    ' タグのマージ
    For Each item In other.Tags
        If Not Contains(mTags, item) Then
            mTags.Add item
        End If
    Next
End Sub

'@Description("コレクションに要素が含まれているかチェックします")
Private Function Contains(ByVal col As Collection, ByVal item As Variant) As Boolean
    Dim existing As Variant
    For Each existing In col
        If existing = item Then
            Contains = True
            Exit Function
        End If
    Next
    Contains = False
End Function